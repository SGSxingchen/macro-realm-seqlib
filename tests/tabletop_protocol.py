"""Real two-worker HTTP/WebSocket smoke test without browser or mocked APIs."""
import asyncio
import io
import json
import os
from pathlib import Path
import signal
import subprocess
import sys
import tempfile
import time
import uuid
import httpx
from PIL import Image
from websockets.asyncio.client import connect
from websockets.exceptions import ConnectionClosed

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'web/frontend/test-results/tabletop'
OUT.mkdir(parents=True,exist_ok=True)
BASE='http://127.0.0.1:8769'
server=None
log=None
checks=[]

def start(folder):
    global server,log
    log=open(OUT/'protocol-server.log','a')
    env={**os.environ,'TABLETOP_CREATE_KEY':'protocol-test-only','TABLETOP_DATA_DIR':folder}
    server=subprocess.Popen([sys.executable,'-m','uvicorn','tests.tabletop_server:app','--host','127.0.0.1','--port','8769','--workers','2','--ws-max-size','1024'],cwd=ROOT,env=env,stdout=log,stderr=log,start_new_session=True)
    for _ in range(150):
        try:
            if httpx.get(BASE+'/api/tabletop/health',trust_env=False,timeout=1).status_code==200:return
        except httpx.HTTPError:pass
        time.sleep(.1)
    raise RuntimeError('Server not ready')

def stop():
    global server,log
    if server and server.poll() is None:
        os.killpg(server.pid,signal.SIGTERM)
        try:server.wait(10)
        except subprocess.TimeoutExpired:os.killpg(server.pid,signal.SIGKILL);server.wait()
    if log:log.close()
    server=log=None

def new_client():return httpx.AsyncClient(base_url=BASE,headers={'Origin':BASE},trust_env=False,timeout=10)
async def call(client,path,data):
    r=await client.post('/api/tabletop'+path,json=data)
    assert r.status_code==200,(r.status_code,r.text)
    return r.json()
async def view(client,room):
    r=await client.get(f'/api/tabletop/{room}/state');assert r.status_code==200,r.text;return r.json()
async def command(client,room,kind='pieces',**kw):
    st=await view(client,room)
    return await call(client,f'/{room}/commands',{'id':uuid.uuid4().hex,'kind':kind,'scene':st['scene']['id'],**kw})
async def receive(ws,revision):
    async with asyncio.timeout(12):
        while True:
            result=json.loads(await ws.recv())
            if result['type']=='snapshot' and result['state']['revision']>=revision:return result['state']
async def socket(client,room):
    return await connect(BASE.replace('http:','ws:')+f'/api/tabletop/{room}/ws',origin=BASE,additional_headers={'Cookie':'realm_tabletop='+client.cookies.get('realm_tabletop')},max_size=12*1024*1024,proxy=None)
def raw(p):return {k:v for k,v in p.items() if k!='editable'}

def passed(name):checks.append(name);print('PASS:',name,flush=True)
async def exercise(folder):
    async with new_client() as gm,new_client() as pl,new_client() as observer,new_client() as other:
        # Every role has its own cookie/session, not shared browser credentials.
        room=(await call(gm,'/rooms',{'name':'协议验收','nick':'KP','key':'protocol-test-only'}))['room']
        invite=(await call(gm,f'/{room}/invite/player',{}))['invite']
        await call(pl,f'/{room}/join',{'nick':'玩家','invite':invite})
        invite=(await call(gm,f'/{room}/invite/spectator',{}))['invite']
        await call(observer,f'/{room}/join',{'nick':'旁观','invite':invite})
        mid=(await view(pl,room))['me']
        gws,pws,ows=await socket(gm,room),await socket(pl,room),await socket(observer,room)
        try:
            for ws in [gws,pws,ows]:await receive(ws,1)
            passed('three authenticated real WebSocket sessions')
            token,hidden,aura=[uuid.uuid4().hex for _ in range(3)]
            st=await command(gm,room,edits=[
                {'id':token,'expected':0,'value':{'id':token,'name':'owned','owners':[mid],'gm_note':'SECRET-NOTE'}},
                {'id':hidden,'expected':0,'value':{'id':hidden,'name':'SECRET-TOKEN','visibility':'gm'}},
                {'id':aura,'expected':0,'value':{'id':aura,'kind':'circle','name':'SECRET-AURA','follow':hidden}},
            ])
            player_view=await receive(pws,st['revision'])
            text=json.dumps(player_view)
            assert 'SECRET-' not in text and hidden not in text and aura not in text
            assert len(player_view['scene']['pieces'])==1
            passed('private object, anchor area and GM note removed from actual network frame')
            piece=player_view['scene']['pieces'][token]
            st=await command(pl,room,edits=[{'id':token,'expected':piece['v'],'value':{**raw(piece),'x':23,'y':12}}])
            host_view=await receive(gws,st['revision']);assert host_view['scene']['pieces'][token]['x']==23
            assert host_view['scene']['pieces'][token]['gm_note']=='SECRET-NOTE'
            passed('player-owned movement syncs to GM without overwriting private fields')
            r=await observer.post(f'/api/tabletop/{room}/commands',json={'id':uuid.uuid4().hex,'kind':'scene.add','data':{'name':'unauthorized'}})
            assert r.status_code==403
            r=await pl.post(f'/api/tabletop/{room}/commands',headers={'Origin':'https://invalid.example'},json={'id':uuid.uuid4().hex,'kind':'undo'})
            assert r.status_code==403
            passed('observer denial and server-side origin checks')
            # Two competing real HTTP requests, sharing a base version.
            current=(await view(pl,room))['scene']['pieces'][token]
            payload={'kind':'pieces','scene':st['scene']['id'],'edits':[{'id':token,'expected':current['v'],'value':{**raw(current),'x':30}}]}
            result=await asyncio.gather(*(c.post(f'/api/tabletop/{room}/commands',json={**payload,'id':uuid.uuid4().hex}) for c in (gm,pl)))
            assert sorted(r.status_code for r in result)==[200,409]
            passed('concurrent network edits reject stale version atomically')
            p=(await view(gm,room))['scene']['pieces'][token]
            value={**raw(p),'x':35}
            payload={'id':uuid.uuid4().hex,'kind':'pieces','scene':st['scene']['id'],'edits':[{'id':token,'expected':p['v'],'value':value}]}
            one=await call(gm,f'/{room}/commands',payload);two=await call(gm,f'/{room}/commands',payload)
            assert one['revision']==two['revision']
            changed=(await view(pl,room))['scene']['pieces'][token]
            await command(pl,room,edits=[{'id':token,'expected':changed['v'],'value':{**raw(changed),'x':40}}])
            response=await gm.post(f'/api/tabletop/{room}/commands',json={'id':uuid.uuid4().hex,'kind':'undo'})
            assert response.status_code==409
            passed('idempotent write receipt and undo cannot erase another actor\'s edit')
            image=io.BytesIO();Image.new('RGB',(60,40),(30,50,70)).save(image,format='PNG')
            response=await gm.post(f'/api/tabletop/{room}/assets',content=image.getvalue());assert response.status_code==200
            asset=response.json()['id']
            assert (await pl.get(f'/api/tabletop/{room}/assets/{asset}')).status_code==404
            st=await view(gm,room);st=await command(gm,room,'scene',expected=st['scene']['v'],data={'background':asset})
            assert (await pl.get(f'/api/tabletop/{room}/assets/{asset}')).status_code==200
            assert (await pl.get(f'/api/tabletop/{room}/assets/{asset}')).headers['cache-control']=='no-store'
            room2=(await call(other,'/rooms',{'name':'other','nick':'other','key':'protocol-test-only'}))['room']
            assert (await other.get(f'/api/tabletop/{room2}/assets/{asset}')).status_code==404
            passed('raster upload, same-room image authorization, cross-room isolation')
            s=(await view(gm,room))['scene']['pieces'][hidden]
            st=await command(gm,room,edits=[{'id':hidden,'expected':s['v'],'value':{**raw(s),'visibility':'selected','viewers':[mid]}}])
            assert hidden in (await receive(pws,st['revision']))['scene']['pieces']
            assert hidden not in (await receive(ows,st['revision']))['scene']['pieces']
            passed('selected-member visibility differs from shared player visibility')
            backup=(await gm.get(f'/api/tabletop/{room}/export')).json()
            assert backup['assets'] and 'invites' not in backup and 'members' not in backup
            response=await other.post(f'/api/tabletop/{room2}/import',json=backup,headers={'X-Operation-Id':uuid.uuid4().hex});assert response.status_code==200,response.text
            imported=next(s for s in response.json()['scenes'] if s['id']!=response.json()['scene']['id'])
            imported_view=await command(other,room2,'scene.switch',scene=imported['id'])
            assert len(imported_view['scene']['pieces'])==3
            assert all(x['visibility']=='gm' and x['owners']==[] for x in imported_view['scene']['pieces'].values())
            passed('portable source-and-image backup imports with fresh IDs and private defaults')
            final=await view(gm,room)
            await asyncio.gather(*(ws.close() for ws in (gws,pws,ows)))
            stop();start(folder)
            after=await view(gm,room)
            assert after['revision']==final['revision'] and after['scene']==final['scene']
            gws,pws,ows=await socket(gm,room),await socket(pl,room),await socket(observer,room)
            assert (await receive(pws,final['revision']))['scene']['pieces'][token]['x']==40
            passed('full two-worker service restart preserves confirmed data and session reconnect')
            await command(gm,room,'member',data={'id':mid})
            async with asyncio.timeout(12):
                try:
                    while True:await pws.recv()
                except ConnectionClosed as closed:assert closed.rcvd.code==4401
            assert (await pl.get(f'/api/tabletop/{room}/state')).status_code==401
            passed('revocation disconnects an existing WebSocket and invalidates its cookie')
        finally:await asyncio.gather(*(ws.close() for ws in (gws,pws,ows)),return_exceptions=True)

def main():
    with tempfile.TemporaryDirectory(prefix='tabletop-protocol-') as folder:
        try:start(folder);asyncio.run(exercise(folder))
        finally:stop()
    (OUT/'protocol-results.json').write_text(json.dumps({'status':'passed','checks':checks,'backend_workers':2,'independent_sessions':3,'mocked_api':False,'browser_tested':False},ensure_ascii=False,indent=2))
    print(f'{len(checks)} real HTTP/WebSocket checks passed')
if __name__=='__main__':main()
