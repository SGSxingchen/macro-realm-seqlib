import { useEffect, useRef, useState } from 'react';
import { FixedSizeList as List } from 'react-window';
import { Resource } from '../types';
import { ResourceCard } from './ResourceCard';

type Props = {
  items: Resource[];
  activePath: string;
  onOpen: (path: string) => void;
  highlightTokens: string[];
  loading: boolean;
};

const CARD_HEIGHT = 116;

export function ResourceList({ items, activePath, onOpen, highlightTokens, loading }: Props) {
  const wrapRef = useRef<HTMLDivElement>(null);
  const [size, setSize] = useState({ w: 0, h: 0 });
  useEffect(() => {
    const el = wrapRef.current;
    if (!el) return;
    const ro = new ResizeObserver(() => setSize({ w: el.clientWidth, h: el.clientHeight }));
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  if (loading && items.length === 0) {
    return (
      <div className="reslist reslist-skeleton" ref={wrapRef}>
        {Array.from({ length: 5 }, (_, i) => <div className="skeleton-card" key={i} />)}
      </div>
    );
  }

  if (!loading && items.length === 0) {
    return (
      <div className="reslist reslist-empty" ref={wrapRef}>
        <p>没有匹配的资源。</p>
        <small>可以尝试：缩短关键词、改用拼音首字母、用空格拆成多个词。</small>
      </div>
    );
  }

  return (
    <div className="reslist" ref={wrapRef}>
      {size.w > 0 && size.h > 0 && (
        <List
          height={size.h}
          width={size.w}
          itemCount={items.length}
          itemSize={CARD_HEIGHT}
          overscanCount={6}
        >
          {({ index, style }) => {
            const it = items[index];
            return (
              <div style={style} className="reslist-row">
                <ResourceCard
                  item={it}
                  active={activePath === it.path}
                  onOpen={() => onOpen(it.path)}
                  highlightTokens={highlightTokens}
                />
              </div>
            );
          }}
        </List>
      )}
    </div>
  );
}
