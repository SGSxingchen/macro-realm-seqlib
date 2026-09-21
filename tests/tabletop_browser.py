"""Real HTTP + WebSocket integration across isolated Chromium contexts.

Starts two production app workers against one temporary SQLite database. No
Playwright route mocks. Only synthetic room/map fixtures are used. CI owns the
process group; restart tests stop that group, never a production process.
"""
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
from PIL import Image, ImageDraw
from playwright.sync_api import sync_playwright, expect

ROOT=Path(__file__).resolve().parents[1]
OUT=ROOT/'web/frontend/test-results/tabletop'
OUT.mkdir(parents=True,exist_ok=True)
BASE='http://127.0.0.1:8767'
server=None
log=None

def start(folder):
    global server,log
    log=open(OUT/'server.log','a')
    env={**os.environ,'TABLETOP_CREATE_KEY':'tabletop-ci-only','TABLETOP_DATA_DIR':folder}
    server=subprocess.Popen([sys.executable,'-m','uvicorn','tests.tabletop_server:app','--host','127.0.0.1','--port','8767','--workers','2','--ws-max-size','1024'],cwd=ROOT,env=env,stdout=log,stderr=log,start_new_session=True)
    for _ in range(100):
        try:
            if httpx.get(BASE+'/api/tabletop/health',trust_env=False,timeout=1).status_code==200:return
        except httpx.HTTPError:pass
        time.sleep(.15)
    raise AssertionError('Real test server did not become healthy')

def stop():
    global server,log
    if server and server.poll() is None:
        os.killpg(server.pid,signal.SIGTERM)
        try:server.wait(10)
        except subprocess.TimeoutExpired:os.killpg(server.pid,signal.SIGKILL);server.wait()
    if log:log.close()
    server=log=None

def get_state(page,room):
    response=page.request.get(BASE+f'/api/tabletop/{room}/state')
    assert response.status==200,response.text()
    return response.json()

def post(page,room,path,data,expected=200):
    response=page.request.post(BASE+f'/api/tabletop/{room}/{path}',data=data,headers={'Origin':BASE})
    assert response.status==expected,(response.status,response.text())
    return response.json()

def command(page,room,kind='pieces',**kw):
    st=get_state(page,room)
    return post(page,room,'commands',{'id':uuid.uuid4().hex,'kind':kind,'scene':st['scene']['id'],**kw})

def edit(page,room,p,**changes):
    value={k:v for k,v in p.items() if k!='editable'}
    return command(page,room,edits=[{'id':p['id'],'expected':p['v'],'value':{**value,**changes}}])

def online(page):expect(page.locator('[data-connection="online"]')).to_be_visible(timeout=20000)
def sync(page,rev):page.wait_for_function('(rev)=>Number(document.querySelector("[data-revision]")?.dataset.revision)>=rev',arg=rev,timeout=20000)
def screenshot(page,name):page.screenshot(path=str(OUT/(name+'.png')),full_page=True,animations='disabled')
def close_panel(page):
    button=page.get_by_role('button',name='关闭战术侧栏',exact=True)
    if button.count():button.click()
def check_layout(page,width):
    values=page.evaluate('''()=>{const board=document.querySelector('.tt-board').getBoundingClientRect();return {width:innerWidth,height:innerHeight,canvasTop:board.top,canvasHeight:board.height,documentWidth:document.documentElement.scrollWidth}}''')
    assert values['documentWidth']<=width+1,values
    assert values['canvasHeight']/values['height']>=.85,values
    return values

def main():
    errors=[];frames=[];active=None
    with tempfile.TemporaryDirectory(prefix='realm-tabletop-') as folder,sync_playwright() as p:
        start(folder)
        browser=p.chromium.launch(headless=True)
        try:
            gmctx=browser.new_context(viewport={'width':1440,'height':960},accept_downloads=True)
            plctx=browser.new_context(viewport={'width':1280,'height':900})
            obctx=browser.new_context(viewport={'width':390,'height':844},is_mobile=True,has_touch=True)
            for ctx in [gmctx,plctx,obctx]:ctx.on('page',lambda page:page.on('pageerror',lambda err:errors.append(str(err))))
            gm=gmctx.new_page();active=gm
            gm.goto(BASE+'/tabletop.html');expect(gm.get_by_role('heading',name='创建战术房间')).to_be_visible();screenshot(gm,'lobby')
            gm.get_by_label('房间名称',exact=True).fill('诸界远征 · 实测战场')
            gm.get_by_label('主持人昵称',exact=True).fill('KP · 青岚')
            gm.get_by_label('建房口令',exact=True).fill('tabletop-ci-only')
            gm.get_by_role('button',name='创建房间 ↗',exact=True).click();online(gm)
            from urllib.parse import urlparse,parse_qs
            room=parse_qs(urlparse(gm.url).query)['room'][0]
            invite=post(gm,room,'invite/player',{})['invite']
            spectator=post(gm,room,'invite/spectator',{})['invite']
            pl=plctx.new_page();active=pl
            pl.on('websocket',lambda ws:ws.on('framereceived',lambda frame:frames.append(frame if isinstance(frame,str) else frame.decode(errors='replace'))))
            pl.goto(BASE+f'/tabletop.html?room={room}#invite={invite}')
            expect(pl.get_by_role('heading',name='加入房间',exact=True)).to_be_visible()
            assert '#invite=' not in pl.url,'Invite must be removed from address bar after capture'
            pl.get_by_label('你的昵称',exact=True).fill('轮回者 · 旅人');pl.get_by_role('button',name='加入战术桌').click();online(pl)
            ob=obctx.new_page();ob.goto(BASE+f'/tabletop.html?room={room}#invite={spectator}')
            ob.get_by_label('你的昵称',exact=True).fill('记录者');ob.get_by_role('button',name='加入战术桌').click();online(ob)
            assert 'realm_tabletop' not in pl.evaluate('document.cookie'),'Identity must be HttpOnly'
            assert gmctx.cookies()!=plctx.cookies(),'Contexts must have independent sessions'
            pm=get_state(pl,room)['me']
            gm.get_by_role('button',name='放置第一个棋子',exact=True).click()
            name=gm.get_by_label('名称',exact=True);expect(name).to_be_visible()
            name.fill('跨界旅人');gm.get_by_role('button',name='保存属性',exact=True).click()
            gm.wait_for_timeout(400)
            st=get_state(gm,room);token=next(iter(st['scene']['pieces'].values()))
            st=edit(gm,room,token,owners=[pm],gm_note='NEVER_SEND_PRIVATE_NOTE',counters=[{'name':'灵力','value':'无限','maximum':''}],statuses=['受护持'])
            sync(pl,st['revision']);token=st['scene']['pieces'][token['id']]
            close_panel(gm)
            # Move by actual pointer dragging, then confirm both authoritative HTTP and peer WS.
            bounds=pl.locator('.tt-board').bounding_box();x=bounds['x']+80+token['x']*18;y=bounds['y']+80+token['y']*18
            pl.mouse.move(x,y);pl.mouse.down();pl.mouse.move(x+90,y+54,steps=12);pl.mouse.up()
            pl.wait_for_timeout(650)
            moved=get_state(pl,room)['scene']['pieces'][token['id']]
            assert abs(moved['x']-13)<.01 and abs(moved['y']-11)<.01,moved
            sync(gm,get_state(pl,room)['revision'])
            host_moved=get_state(gm,room)['scene']['pieces'][token['id']]
            assert abs(host_moved['x']-13)<.05,host_moved
            token=get_state(gm,room)['scene']['pieces'][token['id']]
            aura_id=uuid.uuid4().hex
            st=command(gm,room,edits=[{'id':aura_id,'expected':0,'value':{'id':aura_id,'kind':'circle','name':'半径 6m · 守护领域','follow':token['id'],'radius':6,'color':'#70cdb6'}}])
            secret_id=uuid.uuid4().hex
            st=command(gm,room,edits=[{'id':secret_id,'expected':0,'value':{'id':secret_id,'name':'NEVER_SEND_SECRET_TOKEN','visibility':'gm','x':32,'y':12,'color':'#c8899b'}}])
            hidden_aura=uuid.uuid4().hex
            st=command(gm,room,edits=[{'id':hidden_aura,'expected':0,'value':{'id':hidden_aura,'kind':'circle','name':'NEVER_SEND_HIDDEN_AURA','follow':secret_id,'radius':8}}])
            sync(pl,st['revision'])
            projected=get_state(pl,room)
            assert secret_id not in projected['scene']['pieces'] and hidden_aura not in projected['scene']['pieces']
            assert 'gm_note' not in projected['scene']['pieces'][token['id']]
            assert not any('NEVER_SEND' in frame for frame in frames),'Private values leaked in an actual WS frame'
            # A forged observer write must be rejected by the real server.
            post(ob,room,'commands',{'id':uuid.uuid4().hex,'kind':'pieces','scene':st['scene']['id'],'edits':[{'id':token['id'],'expected':token['v'],'value':{**{k:v for k,v in token.items() if k!='editable'},'x':40}}]},403)
            # Raster asset privacy, actual room upload and cross-session access.
            image=Image.new('RGB',(1200,800),'#182b36');draw=ImageDraw.Draw(image)
            for x0 in range(0,1200,120):draw.line((x0,0,x0,800),fill='#294753',width=2)
            for y0 in range(0,800,120):draw.line((0,y0,1200,y0),fill='#294753',width=2)
            for rect in [(130,100,350,290),(750,100,1080,300),(100,570,420,750),(800,560,1130,750)]:draw.rectangle(rect,fill='#27454c',outline='#486968',width=3)
            raster=io.BytesIO();image.save(raster,format='PNG')
            upload=gm.request.post(BASE+f'/api/tabletop/{room}/assets',data=raster.getvalue(),headers={'Origin':BASE,'Content-Type':'image/png'});assert upload.status==200,upload.text()
            asset=upload.json()['id']
            assert pl.request.get(BASE+f'/api/tabletop/{room}/assets/{asset}').status==404
            st=get_state(gm,room);st=command(gm,room,'scene',expected=st['scene']['v'],data={'background':asset,'map_width':60,'map_height':40,'grid':True,'grid_size':5})
            sync(pl,st['revision']);assert pl.request.get(BASE+f'/api/tabletop/{room}/assets/{asset}').status==200
            assert gm.request.post(BASE+f'/api/tabletop/{room}/assets',data=b'<svg/>',headers={'Origin':BASE}).status==400
            # Batch operations cannot race and overwrite a changed token.
            stale=token;token=get_state(gm,room)['scene']['pieces'][token['id']]
            st=edit(pl,room,get_state(pl,room)['scene']['pieces'][token['id']],x=14)
            post(gm,room,'commands',{'id':uuid.uuid4().hex,'kind':'pieces','scene':st['scene']['id'],'edits':[{'id':token['id'],'expected':token['v'],'value':{**{k:v for k,v in token.items() if k!='editable'},'x':70}}]},409)
            # Per-viewer visibility is not the same as player control.
            secret=get_state(gm,room)['scene']['pieces'][secret_id]
            st=edit(gm,room,secret,visibility='selected',viewers=[pm]);sync(pl,st['revision'])
            assert secret_id in get_state(pl,room)['scene']['pieces'] and secret_id not in get_state(ob,room)['scene']['pieces']
            # Offline UI suspends editing; reconnect pulls current authority without replay.
            plctx.set_offline(True);pl.wait_for_function('()=>document.querySelector("[data-connection]")?.dataset.connection!=="online"')
            expect(pl.get_by_role('button',name='棋子工具',exact=True)).to_be_disabled()
            st=edit(gm,room,get_state(gm,room)['scene']['pieces'][token['id']],x=16)
            plctx.set_offline(False);online(pl);sync(pl,st['revision']);assert get_state(pl,room)['scene']['pieces'][token['id']]['x']==16
            # Durable restart: both server worker processes are killed and recreated.
            revision=st['revision'];stop();pl.wait_for_timeout(1200);start(folder);online(pl);online(gm)
            assert get_state(gm,room)['revision']==revision
            assert get_state(pl,room)['scene']['pieces'][token['id']]['x']==16
            gm.reload();online(gm);assert get_state(gm,room)['revision']==revision
            # Use the drawing UI to create a rectangular zone.
            gm.get_by_role('button',name='矩形工具',exact=True).click();bounds=gm.locator('.tt-board').bounding_box()
            gm.mouse.click(bounds['x']+540,bounds['y']+470);gm.mouse.click(bounds['x']+720,bounds['y']+560)
            expect(gm.get_by_label('名称',exact=True)).to_be_visible();gm.get_by_label('名称',exact=True).fill('停滞结界 · 暂受压制');gm.get_by_label('生效 / 显示强调（取消表示暂停，不删除范围）',exact=True).uncheck();gm.get_by_role('button',name='保存属性',exact=True).click();gm.wait_for_timeout(400);close_panel(gm)
            assert any(obj['kind']=='rect' and not obj['enabled'] for obj in get_state(gm,room)['scene']['pieces'].values())
            # Measure tool does not mutate shared room state.
            before=get_state(gm,room)['revision'];gm.get_by_role('button',name='测距工具',exact=True).click()
            gm.mouse.click(bounds['x']+280,bounds['y']+390);gm.mouse.click(bounds['x']+400,bounds['y']+430)
            assert '平面距离' in gm.locator('.tt-board-footer').inner_text();assert get_state(gm,room)['revision']==before
            gm.get_by_role('button',name='清除测量',exact=True).click();gm.get_by_role('button',name='选择工具',exact=True).click()
            # Portable JSON includes real image data but no credentials; importing is append-only and private.
            backup=gm.request.get(BASE+f'/api/tabletop/{room}/export').json();assert backup['assets'] and 'invites' not in backup
            imported=gm.request.post(BASE+f'/api/tabletop/{room}/import',data=backup,headers={'Origin':BASE,'X-Operation-Id':uuid.uuid4().hex});assert imported.status==200,imported.text()
            assert len(imported.json()['scenes'])==2
            st=get_state(gm,room);new_scene=next(s['id'] for s in st['scenes'] if s['id']!=st['scene']['id'])
            old_scene=st['scene']['id'];st=command(gm,room,'scene.switch',scene=new_scene);sync(pl,st['revision']);assert not get_state(pl,room)['scene']['pieces']
            st=command(gm,room,'scene.switch',scene=old_scene);sync(pl,st['revision'])
            # Actual desktop and mobile geometry/screenshots. Narrow view does not stack chrome.
            metrics=[]
            for width,height in [(1440,960),(1024,768),(390,844),(320,720)]:
                gm.set_viewport_size({'width':width,'height':height});close_panel(gm);gm.wait_for_timeout(300)
                metrics.append(check_layout(gm,width));screenshot(gm,f'desk-{width}')
                gm.get_by_role('button',name='场景',exact=True).first.click();expect(gm.locator('.tt-sidepanel')).to_be_visible()
                assert gm.evaluate('document.documentElement.scrollWidth<=innerWidth+1')
                screenshot(gm,f'panel-{width}');close_panel(gm)
            # Revocation invalidates an already-open connection, not just future logins.
            gm.set_viewport_size({'width':1440,'height':960});active=pl
            command(gm,room,'member',data={'id':pm})
            expect(pl.get_by_role('heading',name='加入房间',exact=True)).to_be_visible(timeout=15000)
            assert pl.locator('canvas').count()==0
            assert pl.request.get(BASE+f'/api/tabletop/{room}/state').status==401
            assert errors==[],errors
            (OUT/'metrics.json').write_text(json.dumps(metrics,indent=2))
            (OUT/'summary.txt').write_text('PASS: three isolated browser contexts, real API/WebSocket, two backend workers, room create/join, pointer drag, shared updates, ownership and observer denial, private snapshot/asset filtering, selective visibility, connection revocation, offline/reconnect, full service restart persistence, real raster upload, safe text fields, CAS, portable backup/import, range and ruler UI, 1440/1024/390/320 layout. No mocked tabletop API.\n')
            print('Real tabletop browser integration: PASS')
        except BaseException:
            if active and not active.is_closed():screenshot(active,'failure')
            raise
        finally:browser.close();stop()
if __name__=='__main__':
    try:main()
    finally:stop()
