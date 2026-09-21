import { useEffect, useState } from 'react';
import { makeId, positionOf, rawPiece, sceneEdit } from './model';
import type { Action, Piece, Snapshot } from './model';
type Base = { state: Snapshot; disabled: boolean; act: (action: Action) => Promise<boolean> };
export function PieceInspector({state,piece,disabled,act,upload,onClose}:Base & {piece:Piece;upload:(file:File)=>Promise<{id:string;width:number;height:number}>;onClose:()=>void}) {
  const [draft,setDraft]=useState(piece),[dirty,setDirty]=useState(false),[notice,setNotice]=useState('');
  const [uploading,setUploading]=useState(false);
  const gm=state.role==='gm', readonly=disabled||!piece.editable||uploading;
  useEffect(()=>{if(!dirty)setDraft(piece);},[piece.v,dirty]);
  const patch=(data:Partial<Piece>)=>{setDraft(p=>({...p,...data}));setDirty(true);setNotice('');};
  const numeric=(key:'x'|'y'|'width'|'height'|'radius'|'angle'|'rotation',label:string,min=-1e7,max=1e7)=><label>{label}<input type="number" step="any" min={min} max={max} required value={draft[key]} onChange={e=>patch({[key]:e.target.valueAsNumber})}/></label>;
  const listToggle=(key:'owners'|'viewers',id:string)=>patch({[key]:(draft[key]||[]).includes(id)?draft[key]!.filter(v=>v!==id):[...(draft[key]||[]),id]});
  return <form className="tt-properties" onSubmit={async e=>{e.preventDefault();if(readonly)return;const ok=await act({kind:'pieces',edits:[{id:draft.id,expected:draft.v,value:rawPiece(draft)}]});if(ok){setDirty(false);setNotice('属性已保存');}}}>
    <div className="tt-panel-title"><div><small>OBJECT / 自由对象</small><h2>{piece.name}</h2></div><button type="button" aria-label="关闭属性面板" onClick={onClose}>×</button></div>
    {!piece.editable&&<p className="tt-note">此对象只读。可见性与控制权由主持人分别设置。</p>}
    {dirty&&piece.v!==draft.v&&<p role="alert" className="tt-warning">此对象已被更新，未覆盖你的草稿。请重新载入后编辑。<button type="button" onClick={()=>{setDraft(piece);setDirty(false);}}>重新载入属性</button></p>}
    <fieldset disabled={readonly}>
      <label>名称<input maxLength={100} required value={draft.name} onChange={e=>patch({name:e.target.value})}/></label>
      <div className="tt-form-pair"><label>标记颜色<input type="color" value={draft.color} onChange={e=>patch({color:e.target.value})}/></label>{numeric('rotation','朝向（度）',-3600,3600)}</div>
      <div className="tt-form-pair">{numeric('x',draft.follow?'相对 X（米）':'X（米）')}{numeric('y',draft.follow?'相对 Y（米）':'Y（米）')}</div>
      {['circle','cone'].includes(draft.kind)?<div className="tt-form-pair">{numeric('radius','半径（米）',.001)}{draft.kind==='cone'&&numeric('angle','扇形角度',.01,360)}</div>:<div className="tt-form-pair">{numeric('width',draft.kind==='token'?'图标直径（米）':draft.kind==='line'?'线宽（米）':'宽度（米）',.001)}{['rect','note'].includes(draft.kind)&&numeric('height','高度（米）',.001)}</div>}
      {draft.kind==='token'&&<><div className="tt-form-pair"><label>高度标注<input maxLength={80} placeholder="例如 12 米 / 楼顶" value={draft.elevation} onChange={e=>patch({elevation:e.target.value})}/></label><label>体型 / 实际占地<input maxLength={80} placeholder="自由描述，不随头像缩放" value={draft.footprint} onChange={e=>patch({footprint:e.target.value})}/></label></div>
        {gm&&<label className="tt-upload">上传棋子图片<input type="file" accept="image/png,image/jpeg,image/webp" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;setUploading(true);try{const a=await upload(f);patch({asset:a.id});}catch(error){setNotice(error instanceof Error?error.message:'上传失败');}finally{setUploading(false);e.target.value='';}}}/></label>}
        {gm&&draft.asset&&<button type="button" onClick={()=>patch({asset:''})}>移除头像</button>}
      </>}
      {draft.kind!=='token'&&<label>范围位置<select value={draft.follow} onChange={e=>{const next=e.target.value,absolute=positionOf(draft,state.scene.pieces),target=state.scene.pieces[next];patch({follow:next,x:absolute.x-(target?.x||0),y:absolute.y-(target?.y||0)});}}><option value="">固定在地图</option>{Object.values(state.scene.pieces).filter(p=>p.kind==='token').map(p=><option key={p.id} value={p.id}>跟随：{p.name}</option>)}</select></label>}
      <label className="tt-check"><input type="checkbox" checked={draft.enabled} onChange={e=>patch({enabled:e.target.checked})}/>生效 / 显示强调（取消表示暂停，不删除范围）</label>
      <label>状态标签（每行一条）<textarea rows={3} value={draft.statuses.join('\n')} maxLength={2400} onChange={e=>patch({statuses:e.target.value.split('\n').slice(0,24)})} placeholder="灼烧 · 3层\n持续至我的下个小回合"/></label>
      <div className="tt-section-heading"><h3>自定义计数器</h3><button type="button" disabled={draft.counters.length>=16} onClick={()=>patch({counters:[...draft.counters,{name:'',value:'',maximum:''}]})}>＋ 添加</button></div>
      <p className="tt-help">文字、负值、“无限”均可；不自动扣减或判死。</p>
      {draft.counters.map((counter,i)=><div className="tt-counter-row" key={i}>{(['name','value','maximum'] as const).map((key,index)=><input key={key} aria-label={`计数器${i+1}${['名称','当前值','上限'][index]}`} placeholder={['名称','当前','上限'][index]} maxLength={key==='name'?40:80} value={counter[key]} onChange={e=>patch({counters:draft.counters.map((c,n)=>i===n?{...c,[key]:e.target.value}:c)})}/>)}<button type="button" aria-label={`删除计数器${i+1}`} onClick={()=>patch({counters:draft.counters.filter((_,n)=>n!==i)})}>×</button></div>)}
      <label>公开说明 / 规则原文<textarea rows={5} maxLength={12000} value={draft.note} onChange={e=>patch({note:e.target.value})} placeholder="保留自定义技能、条件和备注，不要求固定模板。"/></label>
      <label>资料路径或来源备注<input maxLength={2000} value={draft.source} onChange={e=>patch({source:e.target.value})} placeholder="可粘贴序列库的文件路径"/></label>
      {gm&&<details className="tt-permissions"><summary>控制权与可见性</summary>
        <label className="tt-check"><input type="checkbox" checked={draft.locked} onChange={e=>patch({locked:e.target.checked})}/>锁定玩家编辑</label>
        <h3>谁可以操作</h3>{state.members.filter(m=>m.role==='player').map(m=><label className="tt-check" key={m.id}><input type="checkbox" checked={draft.owners.includes(m.id)} onChange={()=>listToggle('owners',m.id)}/>{m.name} · {m.id.slice(-4)}</label>)}
        <label>谁可以看见<select value={draft.visibility} onChange={e=>patch({visibility:e.target.value as Piece['visibility']})}><option value="all">全体成员</option><option value="gm">仅主持人</option><option value="selected">指定成员＋主持人</option></select></label>
        {draft.visibility==='selected'&&state.members.filter(m=>m.role!=='gm').map(m=><label className="tt-check" key={m.id}><input type="checkbox" checked={(draft.viewers||[]).includes(m.id)} onChange={()=>listToggle('viewers',m.id)}/>{m.name} · {m.id.slice(-4)}</label>)}
        <label>主持人私密备注<textarea rows={3} maxLength={12000} value={draft.gm_note||''} onChange={e=>patch({gm_note:e.target.value})}/></label>
        <p className="tt-help">未授权内容不会发给玩家；已经公开过的信息无法从他人的记忆或截图中撤回。</p>
      </details>}
      <button className="tt-primary tt-save" type="submit" disabled={!dirty}>保存属性</button>
    </fieldset>
    {notice&&<p role="status" className="tt-note">{notice}</p>}
    {piece.source.startsWith('序列库/')&&<a className="tt-text-link" href={'/?open='+encodeURIComponent(piece.source)} target="_blank" rel="noopener noreferrer">在序列库查阅来源 ↗</a>}
  </form>;
}
export function SceneSettings({state,disabled,act,upload,onCalibrate}:Base & {upload:(file:File)=>Promise<{id:string;width:number;height:number}>;onCalibrate:()=>void}) {
  const scene=state.scene,[name,setName]=useState(scene.name),[newName,setNewName]=useState(''),[notice,setNotice]=useState(''),[uploading,setUploading]=useState(false);
  useEffect(()=>setName(scene.name),[scene.id,scene.name]);
  return <div className="tt-properties"><h2>场景与底图</h2><p className="tt-help">二维自由坐标，单位为米。底图完整共享，没有雾战或自动碰撞。</p>
    <fieldset disabled={disabled||uploading||state.role!=='gm'}>
      <form onSubmit={e=>{e.preventDefault();void act(sceneEdit(scene,{name}));}}><label>当前场景名称<input value={name} required maxLength={100} onChange={e=>setName(e.target.value)}/></label><button>保存名称</button></form>
      <div className="tt-form-pair"><label className="tt-check"><input type="checkbox" checked={scene.grid} onChange={e=>void act(sceneEdit(scene,{grid:e.target.checked}))}/>显示网格</label><label className="tt-check"><input type="checkbox" checked={scene.snap} onChange={e=>void act(sceneEdit(scene,{snap:e.target.checked}))}/>吸附网格</label></div>
      <form key={scene.v} onSubmit={e=>{e.preventDefault();const data=new FormData(e.currentTarget);void act(sceneEdit(scene,{grid_size:Number(data.get('size')),map_width:Number(data.get('width')),map_height:Number(data.get('height'))}));}}>
        <label>每格（米）<input name="size" type="number" min="0.001" max="10000000" step="any" defaultValue={scene.grid_size} required/></label>
        <div className="tt-form-pair"><label>底图宽（米）<input name="width" type="number" min="0.001" max="10000000" step="any" defaultValue={scene.map_width} required/></label><label>底图高（米）<input name="height" type="number" min="0.001" max="10000000" step="any" defaultValue={scene.map_height} required/></label></div><button>保存尺度</button>
      </form>
      <label className="tt-upload">上传 / 更换底图<input type="file" accept="image/png,image/jpeg,image/webp" onChange={async e=>{const f=e.target.files?.[0];if(!f)return;setUploading(true);try{const a=await upload(f);await act(sceneEdit(scene,{background:a.id,map_width:60,map_height:60*a.height/a.width}));setNotice('底图已上传。可用两点校准确定实际尺寸。');}catch(error){setNotice(error instanceof Error?error.message:'上传失败');}finally{setUploading(false);e.target.value='';}}}/></label>
      <p className="tt-help">PNG / JPG / WebP，单张 ≤5MB；每房间最多80份素材。只上传可向玩家展示的底图。</p>
      <button type="button" disabled={!scene.background} onClick={onCalibrate}>两点校准比例尺</button>
      {scene.background&&<button type="button" onClick={()=>void act(sceneEdit(scene,{background:''}))}>移除底图</button>}
      <form onSubmit={async e=>{e.preventDefault();if(await act({kind:'scene.add',data:{name:newName}}))setNewName('');}}><label>新场景名称<input value={newName} maxLength={100} required onChange={e=>setNewName(e.target.value)} placeholder="领域内部 / 城市街区"/></label><button>创建新场景</button></form>
      <button className="tt-danger" type="button" disabled={state.scenes.length<2} onClick={()=>{if(confirm('删除当前场景及其中所有对象？此操作不能撤销，建议先导出备份。'))void act({kind:'scene.delete',scene:scene.id,expected:state.revision});}}>删除当前场景</button>
    </fieldset>{notice&&<p role="status">{notice}</p>}
  </div>;
}
export function TurnPanel({state,disabled,act}:Base) {
  const scene=state.scene, order=scene.order.filter(id=>scene.pieces[id]), editable=state.role==='gm'&&!disabled;
  return <div className="tt-properties"><h2>行动记录</h2><p className="tt-help">手动记录；不锁定非当前单位，不扣 AP，不刷新冷却或状态。</p>
    <label>轮次<input type="number" min="0" max="1000000" value={scene.round} disabled={!editable} onChange={e=>{if(Number.isFinite(e.target.valueAsNumber))void act(sceneEdit(scene,{round:e.target.valueAsNumber}));}}/></label>
    <button disabled={!editable||!order.length} onClick={()=>{const i=order.indexOf(scene.turn);void act(sceneEdit(scene,{turn:order[(i+1)%order.length]}));}}>下一位（只推进记录）</button>
    {order.map((id,i)=><div className={'tt-turn-row'+(scene.turn===id?' active':'')} key={id}><button disabled={!editable} onClick={()=>void act(sceneEdit(scene,{turn:id}))}>{i+1} · {scene.pieces[id].name}</button><button disabled={!editable||i===0} aria-label={`上移${scene.pieces[id].name}`} onClick={()=>{const next=[...order];[next[i-1],next[i]]=[next[i],next[i-1]];void act(sceneEdit(scene,{order:next}));}}>↑</button><button disabled={!editable} aria-label={`移出行动列表${scene.pieces[id].name}`} onClick={()=>void act(sceneEdit(scene,{order:order.filter(k=>k!==id),turn:scene.turn===id?'':scene.turn}))}>×</button></div>)}
    {editable&&<label>添加单位<select value="" onChange={e=>{if(e.target.value)void act(sceneEdit(scene,{order:[...order,e.target.value]}));}}><option value="">选择棋子…</option>{Object.values(scene.pieces).filter(p=>p.kind==='token'&&!order.includes(p.id)).map(p=><option key={p.id} value={p.id}>{p.name}</option>)}</select></label>}
  </div>;
}
export function duplicatePieces(pieces:Piece[],me:string,gm:boolean):Action {
  const ids=Object.fromEntries(pieces.map(p=>[p.id,makeId()]));
  return {kind:'pieces',edits:pieces.map(p=>({id:ids[p.id],expected:0,value:{...rawPiece(p),id:ids[p.id],v:0,x:p.x+(p.follow?0:2),y:p.y+(p.follow?0:2),follow:ids[p.follow]||p.follow,...(!gm?{asset:'',owners:[me],viewers:[],visibility:'all' as const,locked:false,gm_note:''}:{})}}))};
}
