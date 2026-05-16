import { useState } from 'react';
import { Facets, TreeNode } from '../types';

type Props = {
  tree: TreeNode[];
  selectedCat: string;
  onPickCat: (path: string) => void;
  facets: Facets | null;
  selectedKinds: string[];
  selectedSides: string[];
  selectedAuthors: string[];
  onToggle: (group: 'kinds' | 'sides' | 'authors', name: string) => void;
  onClearFacets: () => void;
};

export function FilterRail(p: Props) {
  return (
    <aside className="filter-rail terminal-scroll">
      <FacetSection title="作品侧" group="sides" items={p.facets?.sides || []} selected={p.selectedSides} onToggle={p.onToggle} />
      <FacetSection title="资源类型" group="kinds" items={p.facets?.kinds || []} selected={p.selectedKinds} onToggle={p.onToggle} />
      <FacetSection title="创作者" group="authors" items={p.facets?.authors || []} selected={p.selectedAuthors} onToggle={p.onToggle} collapsedDefault />
      {(p.selectedKinds.length + p.selectedSides.length + p.selectedAuthors.length > 0) && (
        <button className="rail-clear ghost-toggle" onClick={p.onClearFacets}>清除筛选</button>
      )}
      <div className="rail-divider" />
      <ClassificationTree nodes={p.tree} selected={p.selectedCat} onPick={p.onPickCat} />
    </aside>
  );
}

function FacetSection({
  title, group, items, selected, onToggle, collapsedDefault = false,
}: {
  title: string;
  group: 'kinds' | 'sides' | 'authors';
  items: { name: string; count: number }[];
  selected: string[];
  onToggle: (g: 'kinds' | 'sides' | 'authors', name: string) => void;
  collapsedDefault?: boolean;
}) {
  const [open, setOpen] = useState(!collapsedDefault);
  if (items.length === 0) return null;
  return (
    <section className={`facet ${open ? 'open' : 'closed'}`}>
      <button className="facet-head" onClick={() => setOpen(o => !o)}>
        <span className="facet-title">{title}</span>
        <em>{items.length}</em>
        <span className="facet-icon">{open ? '−' : '+'}</span>
      </button>
      {open && (
        <div className="facet-body">
          {items.map(it => {
            const on = selected.includes(it.name);
            return (
              <button key={it.name} className={`facet-pill ${on ? 'on' : ''}`} onClick={() => onToggle(group, it.name)}>
                <span>{it.name}</span><em>{it.count}</em>
              </button>
            );
          })}
        </div>
      )}
    </section>
  );
}

function ClassificationTree({ nodes, selected, onPick }: { nodes: TreeNode[]; selected: string; onPick: (p: string) => void }) {
  return (
    <section className="taxonomy-section">
      <div className="facet-head static">
        <span className="facet-title">分类</span>
        <em>{nodes.reduce((s, n) => s + n.count, 0)}</em>
      </div>
      <div className="tree">
        <button className={`tree-root ${!selected ? 'active' : ''}`} onClick={() => onPick('')}>
          <span>全部</span>
          <em>{nodes.reduce((s, n) => s + n.count, 0)}</em>
        </button>
        {nodes.map(n => <TreeBranch key={n.path} node={n} selected={selected} onPick={onPick} depth={0} />)}
      </div>
    </section>
  );
}

function TreeBranch({ node, selected, onPick, depth }: { node: TreeNode; selected: string; onPick: (p: string) => void; depth: number }) {
  const [open, setOpen] = useState(depth === 0);
  const has = node.children.length > 0;
  const isOn = selected === node.path;
  return (
    <div className="tree-item" style={{ '--depth': depth } as React.CSSProperties}>
      <div className="tree-row">
        {has ? (
          <button className="tree-twist" onClick={() => setOpen(o => !o)} aria-label={open ? '收起' : '展开'}>
            {open ? '▾' : '▸'}
          </button>
        ) : <span className="tree-twist-spacer" />}
        <button className={`tree-label ${isOn ? 'active' : ''}`} onClick={() => onPick(node.path)}>
          <span>{node.name}</span>
          <em>{node.count}</em>
        </button>
      </div>
      {has && open && (
        <div className="tree-children">
          {node.children.map(c => <TreeBranch key={c.path} node={c} selected={selected} onPick={onPick} depth={depth + 1} />)}
        </div>
      )}
    </div>
  );
}
