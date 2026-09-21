"""SQLite authority: transactions, per-object CAS, filtered snapshots and guarded undo.

WebSocket transports only filtered snapshots; HTTP commands are idempotent.
No game rules, filesystem resource edits, or user-provided SQL/path names.
"""
from __future__ import annotations
import copy
import hashlib
import hmac
import json
import secrets
import sqlite3
import time
from contextlib import contextmanager
from pathlib import Path
from fastapi import HTTPException
from .models import Command, Piece, Scene

TTL = 30 * 86400

def uid(): return secrets.token_hex(12)
def digest(value): return hashlib.sha256(value.encode()).hexdigest()
def pack(value): return json.dumps(value, ensure_ascii=False, separators=(',', ':'), allow_nan=False)
def deny(code, text): raise HTTPException(code, text)

class Store:
    def __init__(self, folder: Path):
        self.folder = folder
        folder.mkdir(parents=True, exist_ok=True)
        self.path = folder / 'tabletop.sqlite'
        with self.db() as db:
            db.executescript('''
            PRAGMA journal_mode=WAL;
            CREATE TABLE IF NOT EXISTS rooms(id TEXT PRIMARY KEY, state TEXT NOT NULL);
            CREATE TABLE IF NOT EXISTS sessions(hash TEXT PRIMARY KEY, room TEXT, member TEXT, expires REAL);
            CREATE TABLE IF NOT EXISTS receipts(room TEXT, actor TEXT, id TEXT, hash TEXT, revision INTEGER, PRIMARY KEY(room,actor,id));
            CREATE TABLE IF NOT EXISTS undo(room TEXT, actor TEXT, revision INTEGER, scene TEXT, before TEXT, after TEXT, used INTEGER DEFAULT 0);
            CREATE TABLE IF NOT EXISTS assets(id TEXT PRIMARY KEY, room TEXT, data BLOB NOT NULL, width INTEGER, height INTEGER);
            CREATE TABLE IF NOT EXISTS rates(key TEXT PRIMARY KEY, start REAL, hits INTEGER);
            CREATE TABLE IF NOT EXISTS sockets(id TEXT PRIMARY KEY, room TEXT, member TEXT, expires REAL);
            ''')

    @contextmanager
    def db(self, write=False):
        db = sqlite3.connect(self.path, timeout=10)
        db.row_factory = sqlite3.Row
        try:
            db.execute('PRAGMA synchronous=FULL')
            db.execute('BEGIN IMMEDIATE' if write else 'BEGIN')
            yield db
            db.commit()
        except BaseException:
            db.rollback()
            raise
        finally: db.close()

    def _room(self, db, room):
        row = db.execute('SELECT state FROM rooms WHERE id=?', (room,)).fetchone()
        if not row: deny(404, '房间不存在')
        return json.loads(row['state'])

    def _save(self, db, state):
        text = pack(state)
        if len(text.encode()) > 8 * 1024 * 1024: deny(413, '房间内容已达到 8MB 上限，请拆分房间')
        db.execute('UPDATE rooms SET state=? WHERE id=?', (text, state['id']))

    def _auth(self, db, room, token):
        row = db.execute('SELECT member FROM sessions WHERE hash=? AND room=? AND expires>?', (digest(token or ''), room, time.time())).fetchone()
        state = self._room(db, room)
        member = state['members'].get(row['member']) if row else None
        if not member or member.get('revoked'): deny(401, '房间身份已失效，请重新加入')
        return state, member

    def _rate(self, db, key, limit=60, seconds=10):
        now = time.time()
        row = db.execute('SELECT * FROM rates WHERE key=?', (key,)).fetchone()
        if row and now-row['start'] < seconds:
            if row['hits'] >= limit: deny(429, '操作太频繁，请稍后重试')
            db.execute('UPDATE rates SET hits=hits+1 WHERE key=?', (key,))
        else: db.execute('INSERT OR REPLACE INTO rates VALUES(?,?,1)', (key, now))
        db.execute('DELETE FROM rates WHERE start<?', (now-86400,))

    def _session(self, db, room, member):
        token = secrets.token_urlsafe(32)
        db.execute('DELETE FROM sessions WHERE expires<?', (time.time(),))
        db.execute('INSERT INTO sessions VALUES(?,?,?,?)', (digest(token), room, member, time.time()+TTL))
        return token

    def create(self, name, nick, address):
        with self.db(True) as db:
            self._rate(db, 'create:'+address, 8, 3600)
            if db.execute('SELECT count(*) FROM rooms').fetchone()[0] >= 200: deny(409, '服务器房间容量已满')
            room, member, scene = uid(), uid(), uid()
            state = {'id': room, 'name': name, 'revision': 1, 'active': scene,
                     'members': {member: {'id': member, 'name': nick, 'role': 'gm'}},
                     'invites': {}, 'scenes': {scene: Scene(id=scene, v=1).model_dump()}}
            db.execute('INSERT INTO rooms VALUES(?,?)', (room, pack(state)))
            return room, self._session(db, room, member)

    def join(self, room, nick, invite, address):
        with self.db(True) as db:
            self._rate(db, 'join:'+address, 30, 60)
            state = self._room(db, room)
            role = next((r for r, key in state['invites'].items() if hmac.compare_digest(key, digest(invite))), None)
            if role not in ('player', 'spectator'): deny(403, '邀请已失效或不正确')
            if len(state['members']) >= 24: deny(409, '每房间最多 24 个身份（含已撤销身份），请另建房间')
            member = uid()
            state['members'][member] = {'id': member, 'name': nick, 'role': role}
            state['revision'] += 1
            self._save(db, state)
            return self._session(db, room, member)

    def invite(self, room, token, role):
        with self.db(True) as db:
            state, member = self._auth(db, room, token)
            if member['role'] != 'gm': deny(403, '仅主持人可以生成邀请')
            self._rate(db, room+':'+member['id'])
            secret = secrets.token_urlsafe(24)
            state['invites'][role] = digest(secret)
            state['revision'] += 1
            self._save(db, state)
            return secret

    @staticmethod
    def visible(piece, member, pieces):
        def one(p):
            return member['role'] == 'gm' or p['visibility'] == 'all' or (p['visibility'] == 'selected' and member['id'] in p['viewers'])
        if not one(piece): return False
        # No hidden-anchor position leaks through an otherwise public aura.
        if piece['follow']:
            target = pieces.get(piece['follow'])
            return bool(target and one(target))
        return True

    def _view(self, db, state, member):
        gm = member['role'] == 'gm'
        scene = copy.deepcopy(state['scenes'][state['active']])
        pieces = scene['pieces']
        visible = {key: obj for key, obj in pieces.items() if self.visible(obj, member, pieces)}
        for obj in visible.values():
            if not gm:
                obj.pop('gm_note', None)
                obj.pop('viewers', None)
            obj['editable'] = gm or (member['role'] == 'player' and member['id'] in obj['owners'] and not obj['locked'])
        scene['pieces'] = visible
        scene['order'] = [key for key in scene['order'] if key in visible]
        if scene['turn'] not in visible: scene['turn'] = ''
        online = {row[0] for row in db.execute('SELECT DISTINCT member FROM sockets WHERE room=? AND expires>?', (state['id'], time.time()))}
        return {'id': state['id'], 'name': state['name'], 'revision': state['revision'], 'me': member['id'], 'role': member['role'],
                'scene': scene, 'scenes': [{'id': s['id'], 'name': s['name']} for s in state['scenes'].values() if gm or s['id'] == state['active']],
                'members': [{**m, 'online': m['id'] in online} for m in state['members'].values() if not m.get('revoked')],
                'ping': state.get('ping') if state.get('ping',{}).get('expires',0)>time.time() else None,
                'canUndo': bool(db.execute('SELECT 1 FROM undo WHERE room=? AND actor=? AND used=0 LIMIT 1', (state['id'], member['id'])).fetchone())}

    def view(self, room, token):
        with self.db() as db:
            state, member = self._auth(db, room, token)
            return self._view(db, state, member)

    def socket_lease(self, room, token, connection):
        with self.db(True) as db:
            _, member = self._auth(db, room, token)
            db.execute('DELETE FROM sockets WHERE expires<?', (time.time(),))
            if not db.execute('SELECT 1 FROM sockets WHERE id=?', (connection,)).fetchone():
                if db.execute('SELECT count(*) FROM sockets WHERE room=?', (room,)).fetchone()[0] >= 48: deny(429, '房间连接数量过多')
                if db.execute('SELECT count(*) FROM sockets').fetchone()[0] >= 256: deny(429, '服务连接容量已满')
            db.execute('INSERT OR REPLACE INTO sockets VALUES(?,?,?,?)', (connection, room, member['id'], time.time()+25))

    def release(self, connection):
        with self.db(True) as db: db.execute('DELETE FROM sockets WHERE id=?', (connection,))

    def _asset_exists(self, db, room, asset):
        if asset and not db.execute('SELECT 1 FROM assets WHERE id=? AND room=?', (asset, room)).fetchone(): deny(400, '图片不属于当前房间')

    def command(self, room, token, cmd: Command):
        with self.db(True) as db:
            state, member = self._auth(db, room, token)
            if member['role'] == 'spectator': deny(403, '观战者不能修改战场')
            actor, gm = member['id'], member['role'] == 'gm'
            fingerprint = digest(pack(cmd.model_dump()))
            receipt = db.execute('SELECT * FROM receipts WHERE room=? AND actor=? AND id=?', (room, actor, cmd.id)).fetchone()
            if receipt:
                if receipt['hash'] != fingerprint: deny(409, '操作编号已被不同内容使用')
                return self._view(db, state, member)
            self._rate(db, room+':'+actor)
            if cmd.kind != 'pieces' and cmd.kind != 'undo' and not gm: deny(403, '仅主持人可进行此操作')
            rev = state['revision']+1
            scene = state['scenes'].get(cmd.scene)
            if cmd.kind in ('pieces', 'scene', 'scene.delete') and not scene: deny(404, '场景不存在')
            if not gm and cmd.scene != state['active']: deny(403, '只允许操作当前场景')
            if cmd.kind == 'pieces':
                if not cmd.edits or len({e.id for e in cmd.edits}) != len(cmd.edits): deny(400, '无效的批量操作')
                before, after = {}, {}
                all_pieces = scene['pieces']
                for edit in cmd.edits:
                    old = all_pieces.get(edit.id)
                    if edit.expected != (old['v'] if old else 0): deny(409, '对象已被他人修改，已加载最新状态；请核对后重试')
                    if not gm and old and (not self.visible(old, member, all_pieces) or actor not in old['owners'] or old['locked']): deny(403, '没有这个对象的编辑权限')
                    value = edit.value.model_dump() if edit.value else None
                    if value:
                        if value['id'] != edit.id: deny(400, '对象编号不一致')
                        if any(len(s) > 100 for s in value['statuses']): deny(400, '状态标签过长')
                        if not gm:
                            for field in ('owners', 'visibility', 'viewers', 'locked', 'gm_note', 'asset', 'kind'):
                                required = old[field] if old else {'owners': [actor], 'visibility': 'all', 'viewers': [], 'locked': False, 'gm_note': '', 'asset': '', 'kind': value['kind']}[field]
                                # Private fields never leave the server and cannot be patched by players.
                                if field in ('gm_note', 'viewers'):
                                    if value[field]: deny(403, '不能写入主持人私密字段')
                                    value[field] = copy.deepcopy(required)
                                elif value[field] != required: deny(403, '不能修改控制权、隐私或图片权限')
                        for mid in value['owners']+value['viewers']:
                            if mid not in state['members'] or state['members'][mid].get('revoked'): deny(400, '成员不存在或已撤销')
                        if value['kind'] == 'polygon' and len(value['points']) < 3: deny(400, '多边形至少需要三个顶点')
                        if value['kind'] == 'line' and len(value['points']) < 2: deny(400, '路径至少需要两个顶点')
                        self._asset_exists(db, room, value['asset'])
                        value['v'] = rev
                    before[edit.id], after[edit.id] = copy.deepcopy(old), value
                updated = {**all_pieces}
                for key, value in after.items():
                    if value is None: updated.pop(key, None)
                    else: updated[key] = value
                if len(updated) > 500: deny(413, '每场景最多 500 个对象')
                for obj in updated.values():
                    if obj['follow']:
                        anchor = updated.get(obj['follow'])
                        if not anchor or anchor['kind'] != 'token' or obj['kind'] == 'token': deny(400, '跟随范围必须绑定现有棋子；删除棋子前先解除范围绑定')
                        if obj['id'] in after and not gm and not self.visible(anchor, member, updated): deny(403, '不能绑定不可见棋子')
                scene['pieces'] = updated
                db.execute('INSERT INTO undo(room,actor,revision,scene,before,after) VALUES(?,?,?,?,?,?)', (room, actor, rev, cmd.scene, pack(before), pack(after)))
            elif cmd.kind == 'undo':
                entry = db.execute('SELECT rowid,* FROM undo WHERE room=? AND actor=? AND used=0 ORDER BY revision DESC LIMIT 1', (room, actor)).fetchone()
                if not entry: deny(409, '没有可撤销的对象操作')
                undo_scene = state['scenes'].get(entry['scene'])
                if not undo_scene or (not gm and entry['scene'] != state['active']): deny(409, '请切回原场景再撤销')
                before, after = json.loads(entry['before']), json.loads(entry['after'])
                for key, expected in after.items():
                    current = undo_scene['pieces'].get(key)
                    if (expected is None and current is not None) or (expected is not None and (not current or current['v'] != expected['v'])): deny(409, '后续有人修改过这些对象，撤销已阻止，未覆盖他人操作')
                    if current and not gm and (current['locked'] or actor not in current['owners'] or not self.visible(current, member, undo_scene['pieces'])): deny(403, '当前没有撤销权限')
                proposed = copy.deepcopy(undo_scene['pieces'])
                for key, old in before.items():
                    if old is None: proposed.pop(key, None)
                    else: proposed[key] = {**old, 'v': rev}
                if len(proposed) > 500: deny(409, '撤销将超过对象上限')
                for obj in proposed.values():
                    if obj['follow'] and (obj['follow'] not in proposed or proposed[obj['follow']]['kind'] != 'token'): deny(409, '后续建立了跟随关系，请先解除关联再撤销')
                undo_scene['pieces'] = proposed
                db.execute('UPDATE undo SET used=1 WHERE rowid=?', (entry['rowid'],))
                # Rebase only the immediately preceding own history touching these
                # objects, and only if it expected the exact restored revision.
                previous = db.execute('SELECT rowid,after FROM undo WHERE room=? AND actor=? AND used=0 AND scene=? ORDER BY revision DESC LIMIT 1', (room, actor, entry['scene'])).fetchone()
                if previous:
                    values = json.loads(previous['after'])
                    for key, old in before.items():
                        if old and values.get(key) and values[key]['v'] == old['v']: values[key]['v'] = rev
                    db.execute('UPDATE undo SET after=? WHERE rowid=?', (pack(values), previous['rowid']))
            elif cmd.kind == 'scene.add':
                if len(state['scenes']) >= 20: deny(409, '每房间最多 20 个场景')
                new = Scene(id=uid(), name=str(cmd.data.get('name', '新场景')), v=rev).model_dump()
                state['scenes'][new['id']] = new
                state['active'] = new['id']
            elif cmd.kind == 'scene.switch':
                if cmd.scene not in state['scenes']: deny(404, '场景不存在')
                state['active'] = cmd.scene
            elif cmd.kind == 'scene.delete':
                if cmd.expected != state['revision']: deny(409, '战场已改变，请核对后重新删除场景')
                if len(state['scenes']) <= 1: deny(409, '至少保留一个场景')
                del state['scenes'][cmd.scene]
                if state['active'] == cmd.scene: state['active'] = next(iter(state['scenes']))
            elif cmd.kind == 'scene':
                if cmd.expected != scene['v']: deny(409, '场景设置已被修改，请刷新后重试')
                if set(cmd.data)-{'name','grid','snap','grid_size','background','map_width','map_height','round','turn','order'}: deny(400, '无效场景字段')
                new = Scene.model_validate({**scene, **cmd.data, 'v': rev}).model_dump()
                self._asset_exists(db, room, new['background'])
                if new['turn'] and new['turn'] not in new['pieces']: deny(400, '行动棋子不存在')
                if any(key not in new['pieces'] for key in new['order']): deny(400, '行动列表含不存在棋子')
                state['scenes'][cmd.scene] = new
            elif cmd.kind == 'member':
                target = state['members'].get(cmd.data.get('id'))
                if not target or target['role'] == 'gm': deny(400, '不能撤销主持人')
                target['revoked'] = True
                for current_scene in state['scenes'].values():
                    for obj in current_scene['pieces'].values():
                        if target['id'] in obj['owners'] or target['id'] in obj['viewers']:
                            obj['owners'] = [m for m in obj['owners'] if m != target['id']]
                            obj['viewers'] = [m for m in obj['viewers'] if m != target['id']]
                            obj['v'] = rev
                db.execute('DELETE FROM sessions WHERE room=? AND member=?', (room, target['id']))
                db.execute('DELETE FROM sockets WHERE room=? AND member=?', (room, target['id']))
            elif cmd.kind == 'import':
                if cmd.data.get('format') != 'realm-tabletop-1': deny(400, '不支持的备份格式')
                sources = cmd.data.get('scenes')
                if not isinstance(sources, list) or not sources or len(state['scenes'])+len(sources)>20: deny(400, '导入后场景不能超过 20 个')
                import base64
                from .media import sanitize
                assets=cmd.data.get('assets', [])
                if not isinstance(assets,list) or len(assets)>80: deny(400,'备份素材数量过多')
                count,size=db.execute('SELECT count(*),coalesce(sum(length(data)),0) FROM assets WHERE room=?',(room,)).fetchone()
                if count+len(assets)>80: deny(413,'导入后素材超过80份')
                asset_map={}
                for item in assets:
                    if not isinstance(item,dict) or not isinstance(item.get('id'),str) or item['id'] in asset_map: deny(400,'备份图片编号无效')
                    raw,w,h=sanitize(base64.b64decode(item.get('data',''),validate=True))
                    size+=len(raw)
                    if size>80*1024*1024: deny(413,'导入后素材超过80MB')
                    asset_id=uid(); asset_map[item['id']]=asset_id
                    db.execute('INSERT INTO assets VALUES(?,?,?,?,?)',(asset_id,room,raw,w,h))
                for source in sources:
                    imported = Scene.model_validate(source).model_dump()
                    imported['background']=asset_map.get(imported['background'], imported['background'])
                    self._asset_exists(db, room, imported['background'])
                    mapping = {key: uid() for key in imported['pieces']}
                    remapped = {}
                    for key, obj in imported['pieces'].items():
                        obj['asset']=asset_map.get(obj['asset'], obj['asset'])
                        self._asset_exists(db, room, obj['asset'])
                        obj.update(id=mapping[key], v=rev, owners=[], viewers=[], visibility='gm')
                        if obj['follow'] and obj['follow'] not in mapping: deny(400, '备份跟随目标丢失')
                        obj['follow'] = mapping.get(obj['follow'], '')
                        remapped[obj['id']] = obj
                    if any(p['follow'] and (remapped[p['follow']]['kind']!='token' or p['kind']=='token') for p in remapped.values()): deny(400, '备份跟随关系无效')
                    imported.update(id=uid(), pieces=remapped, v=rev, order=[mapping[k] for k in imported['order'] if k in mapping], turn=mapping.get(imported['turn'], ''))
                    state['scenes'][imported['id']] = imported
            elif cmd.kind == 'ping':
                x, y = float(cmd.data.get('x', 0)), float(cmd.data.get('y', 0))
                if not (-1e7<=x<=1e7 and -1e7<=y<=1e7): deny(400, '无效坐标')
                # Deliberate public GM indicator, never inferred from hidden objects.
                state['ping'] = {'x': x, 'y': y, 'scene': state['active'], 'expires': time.time()+5}
            state['revision'] = rev
            self._save(db, state)
            db.execute('INSERT INTO receipts VALUES(?,?,?,?,?)', (room, actor, cmd.id, fingerprint, rev))
            # Persist dedupe for the room's lifespan; history alone is bounded.
            db.execute('DELETE FROM undo WHERE room=? AND revision<?', (room, rev-200))
            return self._view(db, state, member)

    def upload(self, room, token, data, width, height):
        with self.db(True) as db:
            state, member = self._auth(db, room, token)
            if member['role']!='gm': deny(403, '仅主持人可上传素材')
            self._rate(db, 'asset:'+room, 20, 60)
            count, size = db.execute('SELECT count(*), coalesce(sum(length(data)),0) FROM assets WHERE room=?', (room,)).fetchone()
            if count>=80 or size+len(data)>80*1024*1024: deny(413, '房间素材容量已满（80份/80MB）')
            key=uid()
            db.execute('INSERT INTO assets VALUES(?,?,?,?,?)', (key, room, data, width, height))
            return {'id':key, 'width':width, 'height':height}

    def asset(self, room, token, key):
        with self.db() as db:
            state, member = self._auth(db, room, token)
            if member['role']!='gm':
                scene = self._view(db, state, member)['scene']
                allowed = {scene['background']} | {p['asset'] for p in scene['pieces'].values()}
                if key not in allowed: deny(404, '图片不存在或不可见')
            row = db.execute('SELECT data FROM assets WHERE room=? AND id=?', (room,key)).fetchone()
            if not row: deny(404, '图片不存在或不可见')
            return row['data']

    def export(self, room, token):
        with self.db() as db:
            state, member = self._auth(db, room, token)
            if member['role']!='gm': deny(403, '仅主持人可导出备份')
            # Portable content backup: does not expose sessions/invites/member secrets.
            import base64
            scenes=copy.deepcopy(list(state['scenes'].values()))
            for scene in scenes:
                for obj in scene['pieces'].values(): obj.update(owners=[],viewers=[],visibility='gm')
            return {'format':'realm-tabletop-1', 'scenes':scenes,
                    'assets':[{'id':r['id'],'data':base64.b64encode(r['data']).decode()} for r in db.execute('SELECT id,data FROM assets WHERE room=?',(room,))]}
