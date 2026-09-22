import { useEffect, useRef, useState } from 'react';
import { FixedSizeList as List } from 'react-window';
import type { Resource } from '../types';
import { ResourceCard } from './ResourceCard';

type Props = { items: Resource[]; activePath: string; onOpen: (path: string) => void; highlightTokens: string[]; loading: boolean; scrollKey?: string; error?: boolean };
const ROW_HEIGHT = 104;
const positions = new Map<string, number>();

export function ResourceList({ items, activePath, onOpen, highlightTokens, loading, scrollKey = '', error }: Props) {
  const wrap = useRef<HTMLDivElement>(null);
  const list = useRef<List>(null);
  const offset = useRef(positions.get(scrollKey) || 0);
  const [size, setSize] = useState({ width: 0, height: 0 });
  useEffect(() => {
    const node = wrap.current;
    if (!node) return;
    const observer = new ResizeObserver(() => setSize({ width: node.clientWidth, height: node.clientHeight }));
    observer.observe(node);
    return () => { observer.disconnect(); if (positions.size > 40) positions.delete(positions.keys().next().value!); positions.set(scrollKey, offset.current); };
  }, [scrollKey]);
  return <div className="realm-resource-list" ref={wrap} aria-busy={loading} onKeyDown={event => {
    const index = Number((event.target as HTMLElement).closest('[data-resource-index]')?.getAttribute('data-resource-index'));
    if (!['ArrowDown', 'ArrowUp', 'Home', 'End'].includes(event.key) || !items.length || !(event.target as HTMLElement).closest('[data-resource-index]')) return;
    event.preventDefault();
    const next = event.key === 'Home' ? 0 : event.key === 'End' ? items.length - 1 : Math.max(0, Math.min(items.length - 1, index + (event.key === 'ArrowDown' ? 1 : -1)));
    list.current?.scrollToItem(next);
    requestAnimationFrame(() => wrap.current?.querySelector<HTMLAnchorElement>(`[data-resource-index="${next}"]`)?.focus());
  }}>
    {!items.length ? loading ? <div className="realm-skeleton" aria-label="正在加载档案">{Array.from({ length: 5 }, (_, i) => <div key={i} />)}</div> : !error && <div className="realm-no-results"><span aria-hidden="true">∅</span><h3>未找到匹配档案</h3><p>试试更短的关键词，或移除上方筛选条件。</p></div> : size.width > 0 && size.height > 0 && <List ref={list} width={size.width} height={size.height} itemCount={items.length} itemSize={ROW_HEIGHT} itemKey={index => items[index].path} initialScrollOffset={offset.current} onScroll={({ scrollOffset }) => { offset.current = scrollOffset; }} overscanCount={4}>
      {({ index, style }) => <div style={style} className="realm-resource-row"><ResourceCard item={items[index]} index={index} active={activePath === items[index].path} onOpen={() => onOpen(items[index].path)} highlightTokens={highlightTokens} /></div>}
    </List>}
  </div>;
}
