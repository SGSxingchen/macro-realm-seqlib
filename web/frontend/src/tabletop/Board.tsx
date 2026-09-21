import { useEffect, useMemo, useRef, useState } from 'react';
import { Stage, Layer, Group, Circle, Rect, Line, Text, Wedge, Image as KImage } from 'react-konva';
import type Konva from 'konva';
import type { KonvaEventObject } from 'konva/lib/Node';
import { distance, newPiece, pathLength, positionOf, rawPiece, snapPoint } from './model';
import type { Action, Kind, Piece, Point, Snapshot } from './model';
export type Tool='select'|'pan'|'measure'|'calibrate'|'ping'|Kind;
type Props={state:Snapshot;disabled:boolean;tool:Tool;setTool:(tool:Tool)=>void;selected:string[];onSelect:(ids:string[])=>void;act:(action:Action)=>Promise<boolean>;onInspect:()=>void;onCalibrate:(length:number)=>void};
const NAMES: Record<Kind,string>={token:'新棋子',note:'文字标记',circle:'圆形范围',rect:'矩形范围',cone:'扇形范围',line:'线形范围',polygon:'多边形范围'};
function Raster({url,x=0,y=0,width,height}: {url:string;x?:number;y?:number;width:number;height:number}) {
  const [image,setImage]=useState<HTMLImageElement>();
  useEffect(()=>{
    let active=true;const controller=new AbortController();let objectUrl='';setImage(undefined);
    fetch(url,{credentials:'same-origin',signal:controller.signal}).then(r=>{if(!r.ok)throw new Error('图片不可见');return r.blob();}).then(blob=>{
      if(!active)return;
      objectUrl=URL.createObjectURL(blob);const img=new window.Image();img.onload=()=>{if(active)setImage(img);};img.src=objectUrl;
    }).catch(()=>{});
    return()=>{active=false;controller.abort();if(objectUrl)URL.revokeObjectURL(objectUrl);};
  },[url]);
  return image?<KImage image={image} x={x} y={y} width={width} height={height} listening={false}/>:null;
}
export function Board({state,disabled,tool,setTool,selected,onSelect,act,onInspect,onCalibrate}:Props) {
  const wrap=useRef<HTMLDivElement>(null);const stage=useRef<Konva.Stage>(null);
  const [size,setSize]=useState({width:600,height:400});
  const [camera,setCamera]=useState({x:80,y:80,scale:18});
  const cameras=useRef(new Map<string,typeof camera>());
  const cameraRef=useRef(camera);cameraRef.current=camera;
  const [drawing,setDrawing]=useState<Point[]>([]);const [cursor,setCursor]=useState<Point|null>(null);
  const [preview,setPreview]=useState<Record<string,Point>>({});
  const drag=useRef<{origin:Point;pieces:Piece[]}|null>(null);
  const box=useRef<Point|null>(null);const [selectionBox,setSelectionBox]=useState<{a:Point;b:Point}|null>(null);
  const pointerStart=useRef<{x:number;y:number;cx:number;cy:number}|null>(null);
  const pinch=useRef<{distance:number;point:Point;world:Point;scale:number}|null>(null);
  const [space,setSpace]=useState(false);
  const scene=state.scene;
  const pieces=useMemo(()=>Object.fromEntries(Object.entries(scene.pieces).map(([id,p])=>[id,{...p,...preview[id]}])),[scene.pieces,preview]);
  const isPan=tool==='pan'||space;
  useEffect(()=>{
    const el=wrap.current;if(!el)return;
    const observer=new ResizeObserver(()=>setSize({width:el.clientWidth,height:el.clientHeight}));observer.observe(el);return()=>observer.disconnect();
  },[]);
  useEffect(()=>{
    setCamera(cameras.current.get(scene.id)||{x:80,y:80,scale:18});setDrawing([]);setPreview({});onSelect([]);
    return()=>{cameras.current.set(scene.id,cameraRef.current);};
  },[scene.id]);
  useEffect(()=>{setDrawing([]);setSelectionBox(null);},[tool]);
  const point=():Point=>{const p=stage.current?.getPointerPosition()||{x:0,y:0};return{x:(p.x-camera.x)/camera.scale,y:(p.y-camera.y)/camera.scale};};
  const finish=async(points=drawing)=>{
    if(disabled||points.length<2)return;
    if(tool==='calibrate'){onCalibrate(distance(points[0],points[1]));setDrawing([]);setTool('select');return;}
    if(tool==='measure')return;
    if(tool==='polygon'&&points.length<3)return;
    if(!['circle','rect','cone','line','polygon'].includes(tool))return;
    const start=points[0],end=points[points.length-1];
    const piece=newPiece(tool as Kind,start,state.me);piece.name=NAMES[tool as Kind];
    piece.radius=Math.max(.01,distance(start,end));
    if(tool==='rect'){piece.x=Math.min(start.x,end.x);piece.y=Math.min(start.y,end.y);piece.width=Math.max(.01,Math.abs(end.x-start.x));piece.height=Math.max(.01,Math.abs(end.y-start.y));}
    if(tool==='cone')piece.rotation=Math.atan2(end.y-start.y,end.x-start.x)*180/Math.PI-30;
    if(tool==='line'||tool==='polygon')piece.points=points.map(p=>[p.x-start.x,p.y-start.y]);
    const ok=await act({kind:'pieces',edits:[{id:piece.id,expected:0,value:rawPiece(piece)}]});
    if(ok){setDrawing([]);onSelect([piece.id]);setTool('select');onInspect();}
  };
  const click=(event:KonvaEventObject<MouseEvent|TouchEvent>)=>{
    if(pinch.current||isPan||selectionBox)return;
    const p=['measure','calibrate'].includes(tool)?point():snapPoint(point(),scene);
    if(tool==='select'){if(event.target===stage.current)onSelect([]);return;}
    if(tool==='measure'){setDrawing(current=>[...current,p].slice(-100));return;}
    if(disabled)return;
    if(tool==='ping'){void act({kind:'ping',data:p});return;}
    if(tool==='token'||tool==='note'){
      const piece=newPiece(tool,p,state.me);
      if(tool==='note'){piece.width=8;piece.height=3;}
      void act({kind:'pieces',edits:[{id:piece.id,expected:0,value:rawPiece(piece)}]}).then(ok=>{if(ok){onSelect([piece.id]);setTool('select');onInspect();}});return;
    }
    const next=[...drawing,p];setDrawing(next);
    if(['circle','rect','cone','calibrate'].includes(tool)&&next.length===2)void finish(next);
  };
  const zoom=(factor:number,at={x:size.width/2,y:size.height/2})=>setCamera(c=>{
    const scale=Math.max(.03,Math.min(400,c.scale*factor));return {scale,x:at.x-(at.x-c.x)*scale/c.scale,y:at.y-(at.y-c.y)*scale/c.scale};
  });
  const pointerDown=(e:KonvaEventObject<MouseEvent>)=>{
    wrap.current?.focus({preventScroll:true});
    if(isPan||e.evt.button===1){pointerStart.current={x:e.evt.clientX,y:e.evt.clientY,cx:camera.x,cy:camera.y};e.evt.preventDefault();}
    else if(tool==='select'&&e.target===stage.current){box.current=point();}
  };
  const pointerMove=(e:KonvaEventObject<MouseEvent>)=>{
    setCursor(point());
    if(pointerStart.current){const a=pointerStart.current;setCamera(c=>({...c,x:a.cx+e.evt.clientX-a.x,y:a.cy+e.evt.clientY-a.y}));}
    if(box.current&&distance(box.current,point())*camera.scale>5)setSelectionBox({a:box.current,b:point()});
  };
  const pointerUp=()=>{
    pointerStart.current=null;box.current=null;
    if(selectionBox){const {a,b}=selectionBox;const ids=Object.values(pieces).filter(p=>{const q=positionOf(p,pieces);return p.editable&&q.x>=Math.min(a.x,b.x)&&q.x<=Math.max(a.x,b.x)&&q.y>=Math.min(a.y,b.y)&&q.y<=Math.max(a.y,b.y);}).map(p=>p.id);onSelect(ids);setTimeout(()=>setSelectionBox(null),0);}
  };
  const touchCenter=(touches:TouchList)=>{const r=wrap.current!.getBoundingClientRect();return {x:(touches[0].clientX+touches[1].clientX)/2-r.left,y:(touches[0].clientY+touches[1].clientY)/2-r.top};};
  const touchStart=(e:KonvaEventObject<TouchEvent>)=>{
    if(e.evt.touches.length===2){e.evt.preventDefault();stage.current?.find('.tt-draggable').forEach(n=>n.stopDrag());drag.current=null;setPreview({});
      const p=touchCenter(e.evt.touches);pinch.current={distance:Math.hypot(e.evt.touches[0].clientX-e.evt.touches[1].clientX,e.evt.touches[0].clientY-e.evt.touches[1].clientY),point:p,world:{x:(p.x-camera.x)/camera.scale,y:(p.y-camera.y)/camera.scale},scale:camera.scale};
    }else if(isPan){const t=e.evt.touches[0];pointerStart.current={x:t.clientX,y:t.clientY,cx:camera.x,cy:camera.y};}
  };
  const touchMove=(e:KonvaEventObject<TouchEvent>)=>{
    e.evt.preventDefault();
    if(e.evt.touches.length===2&&pinch.current){const p=touchCenter(e.evt.touches),a=pinch.current;const d=Math.hypot(e.evt.touches[0].clientX-e.evt.touches[1].clientX,e.evt.touches[0].clientY-e.evt.touches[1].clientY);const scale=Math.max(.03,Math.min(400,a.scale*d/Math.max(a.distance,1)));setCamera({scale,x:p.x-a.world.x*scale,y:p.y-a.world.y*scale});}
    else if(isPan&&pointerStart.current){const t=e.evt.touches[0],a=pointerStart.current;setCamera(c=>({...c,x:a.cx+t.clientX-a.x,y:a.cy+t.clientY-a.y}));}
  };
  const touchEnd=()=>{pointerStart.current=null;if(pinch.current)setTimeout(()=>{pinch.current=null;},100);};
  const startDrag=(p:Piece,e:KonvaEventObject<DragEvent>)=>{
    e.cancelBubble=true;if(pinch.current)return;
    const ids=selected.includes(p.id)?selected:[p.id];onSelect(ids);
    drag.current={origin:positionOf(p,scene.pieces),pieces:ids.map(id=>scene.pieces[id]).filter(p=>p?.editable&&!p.follow)};
  };
  const moveDrag=(p:Piece,e:KonvaEventObject<DragEvent>)=>{
    e.cancelBubble=true;const active=drag.current;if(!active)return;
    const at=snapPoint(e.target.position(),scene),dx=at.x-active.origin.x,dy=at.y-active.origin.y;
    setPreview(Object.fromEntries(active.pieces.map(p=>[p.id,{x:p.x+dx,y:p.y+dy}])));
  };
  const endDrag=async(p:Piece,e:KonvaEventObject<DragEvent>)=>{
    e.cancelBubble=true;const active=drag.current;drag.current=null;
    if(!active||pinch.current)return;
    const at=snapPoint(e.target.position(),scene),dx=at.x-active.origin.x,dy=at.y-active.origin.y;
    await act({kind:'pieces',edits:active.pieces.map(item=>({id:item.id,expected:item.v,value:{...rawPiece(item),x:item.x+dx,y:item.y+dy}}))});
    setPreview({});
  };
  const gridLines=useMemo(()=>{
    if(!scene.grid)return [];
    let step=scene.grid_size;while(step*camera.scale<12)step*=5;
    const left=-camera.x/camera.scale,top=-camera.y/camera.scale,lines: number[][]=[];
    for(let x=Math.floor(left/step)*step;x<left+size.width/camera.scale;x+=step)lines.push([x,top,x,top+size.height/camera.scale]);
    for(let y=Math.floor(top/step)*step;y<top+size.height/camera.scale;y+=step)lines.push([left,y,left+size.width/camera.scale,y]);
    return lines.slice(0,500);
  },[scene.grid,scene.grid_size,camera,size]);
  const visibleDrawing=cursor&&drawing.length?[...drawing,cursor]:drawing;
  const caption=drawing.length?`${pathLength(visibleDrawing).toFixed(2)} 米 · 平面距离`:'自由坐标 · 米';
  return <div className="tt-board" ref={wrap} tabIndex={0} aria-label="战术地图画布" onKeyDown={e=>{
    if((e.target as HTMLElement).matches('input,textarea,select'))return;
    if(e.key===' '){e.preventDefault();setSpace(true);}
    if(e.key==='Escape'){setDrawing([]);setTool('select');onSelect([]);}
    if(e.key==='Enter'){e.preventDefault();void finish();}
  }} onKeyUp={e=>{if(e.key===' ')setSpace(false);}} onBlur={()=>{setSpace(false);pointerStart.current=null;}}>
    <Stage ref={stage} width={Math.max(size.width,1)} height={Math.max(size.height,1)} onWheel={e=>{e.evt.preventDefault();zoom(e.evt.deltaY>0?1/1.12:1.12,stage.current?.getPointerPosition()||undefined);}}
      onMouseDown={pointerDown} onMouseMove={pointerMove} onMouseUp={pointerUp} onMouseLeave={()=>{pointerStart.current=null;box.current=null;setSelectionBox(null);}}
      onTouchStart={touchStart} onTouchMove={touchMove} onTouchEnd={touchEnd} onClick={click} onTap={click}>
      <Layer x={camera.x} y={camera.y} scaleX={camera.scale} scaleY={camera.scale}>
        {scene.background&&<Raster url={`/api/tabletop/${state.id}/assets/${scene.background}`} width={scene.map_width} height={scene.map_height}/>}
        {gridLines.map((points,i)=><Line key={i} points={points} stroke="#29414b" strokeWidth={.6/camera.scale} listening={false}/>)}
        {Object.values(pieces).sort((a,b)=>Number(a.kind==='token')-Number(b.kind==='token')).map(p=>{
          const at=positionOf(p,pieces),chosen=selected.includes(p.id),labelScale=1/camera.scale;
          const choose=(e:KonvaEventObject<MouseEvent|TouchEvent>)=>{
            if(tool!=='select'||space)return;e.cancelBubble=true;
            const shift='shiftKey'in e.evt&&e.evt.shiftKey;
            onSelect(shift?selected.includes(p.id)?selected.filter(id=>id!==p.id):[...selected,p.id]:[p.id]);
          };
          return <Group name="tt-draggable" key={p.id} x={at.x} y={at.y} draggable={tool==='select'&&!space&&!disabled&&p.editable&&!p.follow}
            onClick={choose} onTap={choose} onDblClick={onInspect} onDblTap={onInspect}
            onDragStart={e=>startDrag(p,e)} onDragMove={e=>moveDrag(p,e)} onDragEnd={e=>{void endDrag(p,e);}} opacity={p.enabled?1:.4}>
            <Group rotation={p.rotation}>
              {p.kind==='token'&&<><Circle radius={p.width/2} fill={p.color} stroke={chosen?'#f4d692':'#172d38'} strokeWidth={(chosen?3:1.5)/camera.scale}/>{p.asset?<Raster url={`/api/tabletop/${state.id}/assets/${p.asset}`} x={-p.width/2} y={-p.width/2} width={p.width} height={p.width}/>:<Text text={Array.from(p.name).slice(0,2).join('')} x={-p.width/2} y={-p.width/5} width={p.width} align="center" fill="#0b1b24" fontSize={p.width*.32} fontStyle="bold" listening={false}/>}<Line points={[0,-p.width/2,0,-p.width/2-.4]} stroke="#f4d692" strokeWidth={2/camera.scale} listening={false}/></>}
              {p.kind==='note'&&<Rect width={p.width} height={p.height} fill="#20313d" stroke={chosen?'#f4d692':p.color} strokeWidth={1.5/camera.scale} cornerRadius={.25}/>}
              {p.kind==='circle'&&<Circle radius={p.radius} fill={p.color+'24'} stroke={chosen?'#f4d692':p.color} strokeWidth={2/camera.scale}/>}
              {p.kind==='cone'&&<Wedge radius={p.radius} angle={p.angle} fill={p.color+'30'} stroke={chosen?'#f4d692':p.color} strokeWidth={2/camera.scale}/>}
              {p.kind==='rect'&&<Rect width={p.width} height={p.height} fill={p.color+'26'} stroke={chosen?'#f4d692':p.color} strokeWidth={2/camera.scale}/>}
              {(p.kind==='polygon'||p.kind==='line')&&<Line points={p.points.flat()} closed={p.kind==='polygon'} fill={p.kind==='polygon'?p.color+'26':undefined} stroke={chosen?'#f4d692':p.color} strokeWidth={p.kind==='line'?Math.max(p.width,2/camera.scale):2/camera.scale} opacity={p.kind==='line'?.65:1} hitStrokeWidth={Math.max(p.width,12/camera.scale)}/>}
            </Group>
            {chosen&&p.kind==='token'&&<Circle radius={p.width/2+.12} stroke="#f4d692" strokeWidth={2/camera.scale} listening={false}/>}
            <Group scaleX={labelScale} scaleY={labelScale} y={p.kind==='token'?p.width/2+.3:0} listening={false}>
              <Rect x={-55} y={0} width={110} height={p.statuses.length?40:24} fill="#0b1822d9" cornerRadius={4}/>
              <Text text={p.name+(p.locked?' 🔒':'')+(p.elevation?' ↑'+p.elevation:'')} x={-51} y={5} width={102} align="center" ellipsis wrap="none" fill="#eef6f5" fontSize={12}/>
              {p.statuses.length>0&&<Text text={p.statuses.join(' · ')} x={-51} y={23} width={102} ellipsis wrap="none" fontSize={10} fill="#edc78b"/>}
            </Group>
          </Group>;
        })}
        {visibleDrawing.length>1&&<Line points={visibleDrawing.flatMap(p=>[p.x,p.y])} stroke="#f1cb82" dash={[6/camera.scale,4/camera.scale]} strokeWidth={2/camera.scale} listening={false}/>}
        {drawing.map((p,i)=><Circle key={i} x={p.x} y={p.y} radius={4/camera.scale} fill="#f1cb82" listening={false}/>)}
        {selectionBox&&<Rect x={Math.min(selectionBox.a.x,selectionBox.b.x)} y={Math.min(selectionBox.a.y,selectionBox.b.y)} width={Math.abs(selectionBox.a.x-selectionBox.b.x)} height={Math.abs(selectionBox.a.y-selectionBox.b.y)} fill="#65dbce20" stroke="#65dbce" strokeWidth={1/camera.scale} listening={false}/>}
        {state.ping&&state.ping.scene===scene.id&&<Circle x={state.ping.x} y={state.ping.y} radius={24/camera.scale} stroke="#f1cb82" strokeWidth={3/camera.scale} listening={false}/>}
      </Layer>
    </Stage>
    <div className="tt-canvas-hint">{tool==='select'?'拖动棋子 · 空白框选 · 空格平移':tool==='pan'?'拖动画布 · 双指缩放':tool==='measure'?'依次点击测折线，Esc 清除':tool==='token'||tool==='note'?'点击地图放置；属性按需填写':tool==='ping'?'点击向所有人指示位置':tool==='calibrate'?'点底图两处，然后输入实际米数':'依次点击定义范围；线 / 多边形点击「完成」'}</div>
    <div className="tt-board-footer"><span>{caption}{drawing.length>1&&<small> · {drawing.length-1} 段</small>}</span>
      {drawing.length>0&&<><button onClick={()=>setDrawing([])}>清除测量</button>{['polygon','line'].includes(tool)&&<button onClick={()=>{void finish();}}>完成范围</button>}</>}
      <span className="tt-grow"/><button aria-label="缩小地图" onClick={()=>zoom(1/1.25)}>−</button><button onClick={()=>setCamera({x:80,y:80,scale:18})}>复位</button><button aria-label="放大地图" onClick={()=>zoom(1.25)}>＋</button>
    </div>
  </div>;
}
