"""Independent /api/tabletop router; no coupling to resource content or admin identity."""
from __future__ import annotations
import asyncio
import hashlib
import hmac
import json
import os
import time
from functools import lru_cache
from pathlib import Path
from typing import Literal
from fastapi import APIRouter, HTTPException, Request, Response, WebSocket, WebSocketDisconnect
from pydantic import ValidationError
from starlette.concurrency import run_in_threadpool
from .models import Command, Create, Join
from .store import Store, TTL, uid
from .media import sanitize

router=APIRouter(prefix='/api/tabletop',tags=['tabletop'])
COOKIE='realm_tabletop'

@lru_cache(maxsize=1)
def store():
    path=os.getenv('TABLETOP_DATA_DIR')
    return Store(Path(path) if path else Path(__file__).resolve().parents[3]/'data'/'tabletop')

def origin_ok(connection):
    origin=connection.headers.get('origin','')
    # Explicit same-origin deployment boundary, including scheme; no reflected Origin.
    scheme='https' if connection.url.scheme in ('https','wss') else 'http'
    expected=f'{scheme}://{connection.headers.get("host", "")}'
    configured={o.strip() for o in os.getenv('TABLETOP_ORIGINS','').split(',') if o.strip()}
    return origin == expected or origin in configured

def mutation_guard(request):
    if not origin_ok(request): raise HTTPException(403,'不允许此来源，请检查 TABLETOP_ORIGINS')

def cookie(response,request,room,token):
    response.set_cookie(COOKIE,token,max_age=TTL,httponly=True,samesite='strict',secure=request.url.scheme=='https',path=f'/api/tabletop/{room}')
    response.headers['Cache-Control']='no-store'

def token(request): return request.cookies.get(COOKIE,'')

async def body(request,limit=1024*1024):
    data=bytearray()
    async for chunk in request.stream():
        data.extend(chunk)
        if len(data)>limit: raise HTTPException(413,'请求过大')
    return bytes(data)

async def parsed(request,model):
    mutation_guard(request)
    try: return model.model_validate_json(await body(request))
    except ValidationError: raise HTTPException(422,'字段或数据类型不正确')

@router.get('/health')
def health():
    return {'ok':True,'creation_enabled':bool(os.getenv('TABLETOP_CREATE_KEY'))}

@router.post('/rooms')
async def create(request: Request,response: Response):
    data=await parsed(request,Create)
    key=os.getenv('TABLETOP_CREATE_KEY','')
    if not key: raise HTTPException(503,'主持人建房未启用，请由站点维护者配置 TABLETOP_CREATE_KEY')
    if not hmac.compare_digest(hashlib.sha256(data.key.encode()).digest(),hashlib.sha256(key.encode()).digest()):
        raise HTTPException(403,'建房口令不正确')
    room,session=await run_in_threadpool(store().create,data.name,data.nick,request.client.host if request.client else 'local')
    cookie(response,request,room,session)
    return {'room':room}

@router.post('/{room}/join')
async def join(room: str,request: Request,response: Response):
    data=await parsed(request,Join)
    session=await run_in_threadpool(store().join,room,data.nick,data.invite,request.client.host if request.client else 'local')
    cookie(response,request,room,session)
    return {'room':room}

@router.get('/{room}/state')
def state(room: str,request: Request,response: Response):
    response.headers['Cache-Control']='no-store'
    return store().view(room,token(request))

@router.post('/{room}/invite/{role}')
def invite(room: str,role: Literal['player','spectator'],request: Request,response: Response):
    mutation_guard(request)
    response.headers['Cache-Control']='no-store'
    return {'invite':store().invite(room,token(request),role),'role':role}

@router.post('/{room}/commands')
async def command(room: str,request: Request,response: Response):
    data=await parsed(request,Command)
    response.headers['Cache-Control']='no-store'
    try: return await run_in_threadpool(store().command,room,token(request),data)
    except (ValidationError, TypeError, ValueError): raise HTTPException(422,'操作字段不正确，未保存改动')

@router.post('/{room}/assets')
async def upload(room: str,request: Request,response: Response):
    mutation_guard(request)
    view=await run_in_threadpool(store().view,room,token(request))
    if view['role']!='gm': raise HTTPException(403,'仅主持人可上传素材')
    raw,width,height=await run_in_threadpool(sanitize,await body(request,5*1024*1024))
    response.headers['Cache-Control']='no-store'
    return await run_in_threadpool(store().upload,room,token(request),raw,width,height)

@router.get('/{room}/assets/{asset}')
def asset(room: str,asset: str,request: Request):
    data=store().asset(room,token(request),asset)
    return Response(data,media_type='image/png',headers={'Cache-Control':'no-store','X-Content-Type-Options':'nosniff'})

@router.get('/{room}/export')
def export(room: str,request: Request):
    from .store import pack
    return Response(pack(store().export(room,token(request))),media_type='application/json',headers={'Cache-Control':'no-store','Content-Disposition':'attachment; filename="realm-tabletop-backup.json"'})

@router.post('/{room}/import')
async def restore(room: str,request: Request,response: Response):
    mutation_guard(request)
    view=await run_in_threadpool(store().view,room,token(request))
    if view['role']!='gm': raise HTTPException(403,'仅主持人可导入备份')
    try:
        data=json.loads(await body(request,120*1024*1024))
        operation=request.headers.get('X-Operation-Id','')
        cmd=Command(id=operation,kind='import',data=data)
        result=await run_in_threadpool(store().command,room,token(request),cmd)
        response.headers['Cache-Control']='no-store'
        return result
    except (ValueError, TypeError, ValidationError): raise HTTPException(422,'备份格式无效，未导入')

@router.websocket('/{room}/ws')
async def live(room: str,websocket: WebSocket):
    if not origin_ok(websocket):
        await websocket.close(code=4403); return
    session=token(websocket)
    lease=uid()
    try:
        await run_in_threadpool(store().socket_lease,room,session,lease)
    except HTTPException:
        await websocket.close(code=4401); return
    await websocket.accept()
    async def receive():
        # This channel is outbound state + heartbeat only. Writes use authenticated
        # HTTP transactions, never a client-supplied room snapshot.
        while True:
            msg=await websocket.receive_text()
            if msg!='ping':
                await websocket.close(code=4400)
                return
            await asyncio.sleep(.5)
    reader=asyncio.create_task(receive())
    last=''; renew=0.; beat=0.
    try:
        while not reader.done():
            now=time.monotonic()
            if now-renew>10:
                await run_in_threadpool(store().socket_lease,room,session,lease); renew=now
            snapshot=await run_in_threadpool(store().view,room,session)
            serialized=json.dumps(snapshot,ensure_ascii=False,separators=(',',':'))
            if serialized!=last:
                await asyncio.wait_for(websocket.send_text(json.dumps({'type':'snapshot','state':snapshot},ensure_ascii=False)),5)
                last=serialized
            if now-beat>5:
                await asyncio.wait_for(websocket.send_json({'type':'heartbeat'}),5); beat=now
            # SQLite is the cross-worker authority. Polling revision/filtered view
            # avoids a process-local broadcast list silently splitting a room.
            await asyncio.sleep(.35)
    except HTTPException:
        await websocket.close(code=4401)
    except (WebSocketDisconnect, RuntimeError, asyncio.TimeoutError, OSError): pass
    finally:
        reader.cancel()
        try: await reader
        except (BaseException,): pass
        await run_in_threadpool(store().release,lease)
