import { useEffect, useState } from 'react';
import type { CSSProperties } from 'react';
import type { Facets, TreeNode } from '../types';

type Group = 'kinds' | 'sides' | 'authors';
type Props = { tree: TreeNode[]; selectedCat: string; onPickCat: (path: string) => void; facets: Facets | null; selectedKinds: string[]; selectedSides: string[]; selectedAuthors: string[]; onToggle: (group: Group, name: string) => void; onClearFacets: () => void };
export function FilterRail(p: Props) {
  return <div className="realm-filters">
    <section><h3>资源分类</h3><button className={!p.selectedCat ? 'filter-category selected' : 'filter-category'} aria-pressed={!p.selectedCat} onClick={() => p.onPickCat('')}>全部档案</button>
      {p.tree.map(node => <Branch key={node.path} node={node} selected={p.selectedCat} onPick={p.onPickCat} depth={0} />)}
    </section>
    <Facet title="作品侧" group="sides" items={p.facets?.sides || []} selected={p.selectedSides} onToggle={p.onToggle} />
    <Facet title="资源类型" group="kinds" items={p.facets?.kinds || []} selected={p.selectedKinds} onToggle={p.onToggle} />
    <Facet title="创作者" group="authors" items={p.facets?.authors || []} selected={p.selectedAuthors} onToggle={p.onToggle} collapsed />
    {(p.selectedCat || p.selectedKinds.length + p.selectedSides.length + p.selectedAuthors.length > 0) && <button onClick={p.onClearFacets}>清除全部筛选</button>}
  </div>;
}
function Facet({ title, group, items, selected, onToggle, collapsed = false }: { title: string; group: Group; items: { name: string; count: number }[]; selected: string[]; onToggle: Props['onToggle']; collapsed?: boolean }) {
  const [open, setOpen] = useState(!collapsed);
  const visible = [...items, ...selected.filter(name => !items.some(item => item.name === name)).map(name => ({ name, count: 0 }))];
  if (!visible.length) return null;
  return <section><button className="filter-section-toggle" aria-expanded={open} onClick={() => setOpen(v => !v)}><span>{title}{selected.length ? ` · ${selected.length}` : ''}</span><span aria-hidden="true">{open ? '−' : '+'}</span></button>
    {open && <div className="filter-options">{visible.map(item => <button key={item.name} aria-pressed={selected.includes(item.name)} className={selected.includes(item.name) ? 'selected' : ''} onClick={() => onToggle(group, item.name)}><span>{item.name}</span><small>{item.count}</small></button>)}</div>}
  </section>;
}
function Branch({ node, selected, onPick, depth }: { node: TreeNode; selected: string; onPick: Props['onPickCat']; depth: number }) {
  const [open, setOpen] = useState(depth === 0 || selected.startsWith(node.path + '/'));
  useEffect(() => { if (selected.startsWith(node.path + '/')) setOpen(true); }, [selected, node.path]);
  return <div className="filter-branch" style={{ '--branch-depth': depth } as CSSProperties}>
    <div className="filter-branch-row">
      {node.children.length > 0 ? <button className="filter-expand" aria-expanded={open} aria-label={`${open ? '收起' : '展开'}${node.name}`} onClick={() => setOpen(v => !v)}>{open ? '▾' : '▸'}</button> : <span className="filter-expand-spacer" />}
      <button className={`filter-category${selected === node.path ? ' selected' : ''}`} aria-pressed={selected === node.path} onClick={() => onPick(node.path)}><span>{node.name}</span><small>{node.count}</small></button>
    </div>
    {open && node.children.map(child => <Branch key={child.path} node={child} selected={selected} onPick={onPick} depth={depth + 1} />)}
  </div>;
}
