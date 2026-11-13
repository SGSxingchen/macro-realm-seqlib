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

function cleanCardSnippet(raw: string, title: string): string {
  let s = raw.replace(/^\uFEFF/, '').trim();
  if (!s) return '';
  if (s.startsWith(title)) {
    s = s.slice(title.length).trim();
  }
  const bareTitle = title.replace(/^\d{1,4}】\s*/, '').trim();
  if (bareTitle && s.startsWith(bareTitle)) {
    s = s.slice(bareTitle.length).trim();
  }
  s = s
    .replace(/^(?:[（(【\[]\s*(?:制作人|制作者|投稿人|调整人|修改人|重置人|审核人|审核者|复查人|文本优化|作者|协力)\s*[:：][^）)】\]]+[）)】\]]\s*)+/g, '')
    .replace(/^(?:[【\[]\s*\d{2,4}\s*[./年-]\s*\d{1,2}[^】\]]*[】\]]\s*)+/g, '')
    .replace(/\s+/g, ' ')
    .trim();
  return s;
}

export function ResourceCard({ item, active, onOpen, highlightTokens, index }: Props) {
  const marked = (text: string) =>
    highlight(text, highlightTokens).map((part, i) =>
      part.mark ? <mark key={i}>{part.text}</mark> : <span key={i}>{part.text}</span>
    );

  const seqMatch = item.title.match(/^(\d{1,4})】\s*(.+)$/);
  const rawSnippet = (item.snippet || '').trim();
  const displaySnippet = highlightTokens.length > 0 ? rawSnippet : cleanCardSnippet(rawSnippet, item.title);
  const redundantSnippet =
    !displaySnippet ||
    displaySnippet === item.category ||
    displaySnippet === item.top_kind ||
    displaySnippet === item.side;
  const fallbackMeta = item.authors?.length
    ? `署名：${item.authors.join(' · ')}`
    : item.category && item.category !== item.top_kind
      ? item.category
      : '点击查阅完整序列原文';

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
        {seqMatch && (
          <span className="card-seq-num" style={{ marginLeft: 0 }}>
            {marked(seqMatch[1])}
          </span>
        )}
        {item.side && item.side !== item.top_kind && (
          <span className="card-tag-side" data-side={sideKey(item.side)}>{item.side}</span>
        )}
        {item.top_kind && (
          <span className="card-tag-kind" style={{ marginLeft: 'auto' }}>
            {item.top_kind}
          </span>
        )}
      </div>
      <strong title={item.title}>
        {seqMatch ? marked(seqMatch[2]) : marked(item.title)}
      </strong>
      <div className="realm-resource-snippet">
        {redundantSnippet ? fallbackMeta : marked(displaySnippet)}
      </div>
    </a>
  );
}
