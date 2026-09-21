import { useCallback, useEffect, useRef, useState } from 'react';
import type { Action, Snapshot } from './model';
import { makeId } from './model';
export class ApiError extends Error {
  constructor(message: string, public status: number) { super(message); }
}
export async function request<T>(url: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch('/api/tabletop'+url, {credentials:'same-origin', ...init});
  if (!response.ok) {
    let message='服务暂不可用';
    try { const error=await response.json(); if(typeof error.detail==='string') message=error.detail; } catch { /* no server HTML rendered */ }
    throw new ApiError(message,response.status);
  }
  return response.json();
}
export const jsonPost = <T,>(url: string, value: unknown) => request<T>(url,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(value)});
export function useRoom(room: string) {
  const [state,setState]=useState<Snapshot|null>(null);
  const [status,setStatus]=useState<'connecting'|'online'|'offline'|'unauthorized'>('connecting');
  const [busy,setBusy]=useState(false);
  const [error,setError]=useState('');
  const pending=useRef(false);
  const current=useRef(state);
  const socket=useRef<WebSocket|null>(null);
  const online=useRef(false);
  const [retry,setRetry]=useState(0);
  const apply=useCallback((next: Snapshot) => {
    if(next.id!==room) return;
    if(!current.current || next.revision>=current.current.revision) {current.current=next;setState(next);}
  },[room]);
  useEffect(() => {
    let stopped=false, attempt=0, lastMessage=Date.now();
    let timer: ReturnType<typeof setTimeout>;
    const connect=() => {
      if(stopped) return;
      setStatus('connecting'); online.current=false; lastMessage=Date.now();
      const url=new URL(`/api/tabletop/${encodeURIComponent(room)}/ws`,location.href);
      url.protocol=location.protocol==='https:'?'wss:':'ws:';
      const ws=new WebSocket(url);socket.current=ws;
      ws.onmessage=event => {
        if(stopped||socket.current!==ws)return;
        lastMessage=Date.now();attempt=0;
        try {
          const msg=JSON.parse(event.data);
          if(msg.type==='snapshot') {apply(msg.state);online.current=true;setStatus('online');}
          if(msg.type==='heartbeat'&&current.current) {online.current=true;setStatus('online');}
        } catch {setError('同步消息无效，请重新连接');ws.close();}
      };
      ws.onclose=async event => {
        if(stopped||socket.current!==ws)return;
        online.current=false;setStatus('offline');
        if(event.code===4401) {setState(null);current.current=null;setStatus('unauthorized');return;}
        // A pre-accept 401 appears as 1006 in browsers; distinguish it by HTTP.
        try {await request<Snapshot>(`/${room}/state`);} catch(e) {
          if(stopped)return;
          if(e instanceof ApiError && [401,404].includes(e.status)) {setState(null);current.current=null;setStatus('unauthorized');return;}
        }
        if(!stopped) timer=setTimeout(connect,Math.min(1000*2**attempt++,10000));
      };
    };
    connect();
    const heartbeat=setInterval(() => {
      const ws=socket.current;
      if(ws?.readyState===WebSocket.OPEN) {
        if(Date.now()-lastMessage>16000) {online.current=false;setStatus('offline');ws.close();}
        else ws.send('ping');
      }
    },5000);
    const offline=()=>{online.current=false;setStatus('offline');socket.current?.close();};
    window.addEventListener('offline',offline);
    return()=>{stopped=true;clearTimeout(timer);clearInterval(heartbeat);window.removeEventListener('offline',offline);socket.current?.close();online.current=false;};
  },[room,retry,apply]);
  const act=async(action: Action) => {
    if(!online.current||pending.current) {setError('连接尚未就绪或上一步正在保存，请稍后操作');return false;}
    pending.current=true;setBusy(true);setError('');
    try {apply(await jsonPost<Snapshot>(`/${room}/commands`,{id:makeId(),scene:current.current?.scene.id,...action}));return true;}
    catch(e) {
      setError(e instanceof Error?e.message:'保存失败');
      if(!(e instanceof ApiError)) setError('网络中断，本次操作结果未确认。重连后核对服务器状态；没有自动重放旧操作。');
      try {apply(await request<Snapshot>(`/${room}/state`));}catch{/* reconnect handles authorization */}
      return false;
    } finally {pending.current=false;setBusy(false);}
  };
  const upload=async(file: File) => {
    if(!online.current||pending.current) throw new Error('请等待连接或保存完成');
    if(file.size>5*1024*1024) throw new Error('单张图片不得超过5MB');
    return request<{id:string;width:number;height:number}>(`/${room}/assets`,{method:'POST',body:file});
  };
  return {state,status,busy,error,setError,act,upload,apply,reconnect:()=>setRetry(n=>n+1)};
}
