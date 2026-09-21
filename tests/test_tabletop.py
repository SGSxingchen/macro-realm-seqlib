import base64
import concurrent.futures
import json
from io import BytesIO
import pytest
from fastapi import HTTPException
from PIL import Image
from web.backend.app.tabletop.store import Store, uid
from web.backend.app.tabletop.models import Command, Piece
from web.backend.app.tabletop.media import sanitize

@pytest.fixture
def desk(tmp_path):
    db=Store(tmp_path/'desk')
    room,host=db.create('测试战场','主持人','local')
    invite=db.invite(room,host,'player')
    player=db.join(room,'玩家',invite,'local')
    observer=db.join(room,'观战',db.invite(room,host,'spectator'),'local')
    return db,room,host,player,observer

def state(d,session): return d[0].view(d[1],session)
def send(d,session,kind='pieces',**kw):
    return d[0].command(d[1],session,Command(id=uid(),kind=kind,scene=state(d,session)['scene']['id'],**kw))
def create(d,session,**kw):
    p=Piece(id=uid(),**kw).model_dump()
    send(d,session,edits=[{'id':p['id'],'expected':0,'value':p}])
    return state(d,session)['scene']['pieces'][p['id']]
def clean(p): return {k:v for k,v in p.items() if k!='editable'}
def patch(d,session,p,**kw):
    return send(d,session,edits=[{'id':p['id'],'expected':p['v'],'value':{**clean(p),**kw}}])

def test_permission_visibility_and_revocation(desk):
    d=desk; h,p,o=d[2:]; mid=state(d,p)['me']
    secret=create(d,h,name='秘密敌人',visibility='gm',gm_note='只有KP知道',x=400)
    assert '秘密敌人' not in json.dumps(state(d,p),ensure_ascii=False)
    aura=create(d,h,kind='circle',follow=secret['id'])
    assert aura['id'] not in state(d,p)['scene']['pieces']
    own=create(d,h,owners=[mid],gm_note='秘密弱点')
    player_piece=state(d,p)['scene']['pieces'][own['id']]
    assert 'gm_note' not in player_piece
    patch(d,p,player_piece,x=8)
    assert state(d,h)['scene']['pieces'][own['id']]['gm_note']=='秘密弱点'
    with pytest.raises(HTTPException) as e: patch(d,o,own,x=10)
    assert e.value.status_code==403
    with pytest.raises(HTTPException): patch(d,p,secret,x=2)
    patch(d,h,secret,visibility='selected',viewers=[mid])
    assert secret['id'] in state(d,p)['scene']['pieces']
    assert secret['id'] not in state(d,o)['scene']['pieces']
    send(d,h,'member',data={'id':mid})
    with pytest.raises(HTTPException) as e: state(d,p)
    assert e.value.status_code==401

def test_player_cannot_escalate_or_steal_assets(desk):
    d=desk; h,p,o=d[2:]; mid=state(d,p)['me']
    piece=create(d,h,owners=[mid])
    for kw in [{'visibility':'gm'},{'owners':[]},{'gm_note':'x'},{'locked':True},{'asset':'secret'},{'kind':'note'}]:
        with pytest.raises(HTTPException): patch(d,p,piece,**kw)
    for k in ('scene.add','scene','scene.switch','member','import'):
        with pytest.raises(HTTPException): send(d,p,k)

def test_batch_atomic_idempotent_and_conflict(desk):
    d=desk; h=d[2]
    a=create(d,h); b=create(d,h)
    cmd=Command(id=uid(),kind='pieces',scene=state(d,h)['scene']['id'],edits=[{'id':a['id'],'expected':a['v'],'value':{**clean(a),'x':3}}])
    first=d[0].command(d[1],h,cmd)
    again=d[0].command(d[1],h,cmd)
    assert first['revision']==again['revision']
    with pytest.raises(HTTPException): patch(d,h,a,x=5)
    with pytest.raises(HTTPException): send(d,h,edits=[{'id':b['id'],'expected':b['v'],'value':{**clean(b),'x':50}}, {'id':a['id'],'expected':a['v'],'value':clean(a)}])
    assert state(d,h)['scene']['pieces'][b['id']]['x']==0
    with pytest.raises(HTTPException): d[0].command(d[1],h,cmd.model_copy(update={'kind':'undo'}))

def test_undo_never_overwrites_other_edits(desk):
    d=desk; h,p=d[2:4]; mid=state(d,p)['me']
    a=create(d,h,owners=[mid])
    s=patch(d,h,a,x=10); a=s['scene']['pieces'][a['id']]
    patch(d,p,a,x=20)
    with pytest.raises(HTTPException) as e: send(d,h,'undo')
    assert e.value.status_code==409
    assert state(d,h)['scene']['pieces'][a['id']]['x']==20

def test_persistence_restart_and_scene_isolation(desk):
    d=desk; h,p=d[2:4]
    a=create(d,h,name='原场景')
    send(d,h,'scene.add',data={'name':'隐藏备团场景'})
    new=state(d,h)['scene']['id']
    create(d,h,name='新场景')
    assert '原场景' not in json.dumps(state(d,p),ensure_ascii=False)
    fresh=Store(d[0].folder)
    assert fresh.view(d[1],h)['scene']['id']==new
    send(d,h,'scene.switch',data={}, **{})
    assert a['id'] not in state(d,p)['scene']['pieces']

def png():
    b=BytesIO();Image.new('RGB',(40,30),(40,80,110)).save(b,format='PNG');return b.getvalue()

def test_media_authorization_and_portable_backup(desk):
    d=desk; h,p=d[2:4]; db,room=d[:2]
    raw,w,hgt=sanitize(png())
    asset=db.upload(room,h,raw,w,hgt)['id']
    secret=create(d,h,asset=asset,visibility='gm')
    with pytest.raises(HTTPException): db.asset(room,p,asset)
    patch(d,h,secret,visibility='all')
    assert db.asset(room,p,asset)==raw
    exported=db.export(room,h)
    assert 'invites' not in exported and 'sessions' not in exported
    room2,host2=db.create('新房间','新KP','local')
    with pytest.raises(HTTPException): db.asset(room2,host2,asset)
    db.command(room2,host2,Command(id=uid(),kind='import',data=exported))
    st=db.view(room2,host2)
    assert len(st['scenes'])==2
    db.command(room2,host2,Command(id=uid(),kind='scene.switch',scene=st['scenes'][1]['id']))
    piece=next(iter(db.view(room2,host2)['scene']['pieces'].values()))
    assert piece['asset']!=asset and piece['visibility']=='gm'
    assert db.asset(room2,host2,piece['asset'])==raw
    with pytest.raises(HTTPException): sanitize(b'<svg onload="alert(1)"/>')

def test_real_concurrent_writers_one_wins(desk):
    d=desk; h=d[2]; a=create(d,h)
    def write(x):
        try: patch(d,h,a,x=x); return True
        except HTTPException as e: assert e.status_code==409; return False
    with concurrent.futures.ThreadPoolExecutor(2) as pool: results=list(pool.map(write,[1,2]))
    assert sorted(results)==[False,True]

def test_invite_rotation_and_same_nick_not_identity(desk):
    db,room,h,p,_=desk
    old=db.invite(room,h,'player'); new=db.invite(room,h,'player')
    with pytest.raises(HTTPException): db.join(room,'玩家',old,'local')
    session=db.join(room,'玩家',new,'local')
    assert db.view(room,session)['me']!=db.view(room,p)['me']
    with pytest.raises(HTTPException): db.invite(room,p,'player')


def test_revoke_removes_object_permissions_without_revealing(desk):
    d=desk; h,p=d[2:4]; mid=state(d,p)['me']
    obj=create(d,h,owners=[mid],visibility='selected',viewers=[mid])
    result=send(d,h,'member',data={'id':mid})
    owned=result['scene']['pieces'][obj['id']]
    assert not owned['owners'] and not owned['viewers']
    assert owned['visibility']=='selected'

def test_scene_delete_uses_room_revision(desk):
    d=desk;h=d[2]
    s=send(d,h,'scene.add',data={'name':'second'})
    create(d,h)
    with pytest.raises(HTTPException) as e:
        send(d,h,'scene.delete',expected=s['revision'])
    assert e.value.status_code==409

def test_http_origin_auth_validation_and_cookie(tmp_path,monkeypatch):
    from fastapi.testclient import TestClient
    from web.backend.app.main import app
    from web.backend.app.tabletop import store as service_store
    monkeypatch.setenv('TABLETOP_DATA_DIR',str(tmp_path/'http'))
    monkeypatch.setenv('TABLETOP_CREATE_KEY','test-only')
    service_store.cache_clear()
    with TestClient(app) as client:
        payload={'name':'test','nick':'GM','key':'test-only'}
        assert client.post('/api/tabletop/rooms',json=payload).status_code==403
        assert client.post('/api/tabletop/rooms',json=payload,headers={'Origin':'https://evil.example'}).status_code==403
        response=client.post('/api/tabletop/rooms',json=payload,headers={'Origin':'http://testserver'})
        assert response.status_code==200
        room=response.json()['room']
        assert 'HttpOnly' in response.headers['set-cookie'] and 'SameSite=strict' in response.headers['set-cookie']
        assert f'Path=/api/tabletop/{room}' in response.headers['set-cookie']
        assert client.get(f'/api/tabletop/{room}/state').json()['role']=='gm'
        assert client.post(f'/api/tabletop/{room}/commands',json={'id':'bad','kind':'evil'},headers={'Origin':'http://testserver'}).status_code==422
        with client.websocket_connect(f'/api/tabletop/{room}/ws',headers={'Origin':'http://testserver'}) as ws:
            snapshot=ws.receive_json()
            assert snapshot['type']=='snapshot'
            assert snapshot['state']['role']=='gm'
        unauthorized=TestClient(app)
        assert unauthorized.get(f'/api/tabletop/{room}/state').status_code==401
    service_store.cache_clear()
