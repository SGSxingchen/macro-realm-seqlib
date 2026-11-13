import type { Resource } from '../types';
import { highlight } from '../utils';
import { resourceLink } from './Reader/adaptive';

type Props = {
  item: Resource;
  active: boolean;
  onOpen: () => void;
  highlightTokens: string[];
  index?: number;
};

export function sideKey(side?: string): string | undefined {
  if (!side) return undefined;
  if (side.includes('战技')) return 'martial';
  if (side.includes('神秘')) return 'mystic';
  if (side.includes('科技')) return 'tech';
  if (side.includes('特殊')) return 'special';
  return undefined;
}

export function ResourceCard({ item, active, onOpen, highlightTokens, index }: Props) {
  const marked = (text: string) =>
    highlight(text, highlightTokens).map((part, i) =>
      part.mark ? <mark key={i}>{part.text}</mark> : <span key={i}>{part.text}</span>
    );

  const seqMatch = item.title.match(/^(\d{1,4})】\s*(.+)$/);
  const rawSnippet = (item.snippet || '').trim();
  const redundantSnippet =
    !rawSnippet || rawSnippet === item.category || rawSnippet === item.top_kind || rawSnippet === item.side;
  const fallbackMeta = item.authors?.length
    ? `制作 / 审核：${item.authors.join(' · ')}`
    : item.category && item.category !== item.top_kind
      ? item.category
      : '点击展开完整规则档案';

  return (
    <a
      href={resourceLink(item.path)}
      className={`realm-resource-card${active ? ' selected' : ''}`}
      data-resource-index={index}
      aria-current={active ? 'true' : undefined}
      onClick={event => {
        if (event.button === 0 && !event.metaKey && !event.ctrlKey && !event.shiftKey && !event.altKey) {
          event.preventDefault();
          onOpen();
        }
      }}
    >
      <div className="realm-resource-tags">
        {item.top_kind && <span className="card-tag-kind">{item.top_kind}</span>}
        {item.side && item.side !== item.top_kind && (
          <span className="card-tag-side" data-side={sideKey(item.side)}>{item.side}</span>
        )}
        {seqMatch && <span className="card-seq-num">#{marked(seqMatch[1])}</span>}
      </div>
      <strong title={item.title}>
        {seqMatch ? marked(seqMatch[2]) : marked(item.title)}
      </strong>
      <div className="realm-resource-snippet">
        {redundantSnippet ? fallbackMeta : marked(rawSnippet)}
      </div>
    </a>
  );
}
