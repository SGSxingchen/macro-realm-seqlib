import type { Resource } from '../types';
import { highlight } from '../utils';
import { resourceLink } from './Reader/adaptive';

type Props = { item: Resource; active: boolean; onOpen: () => void; highlightTokens: string[]; index?: number };
export function ResourceCard({ item, active, onOpen, highlightTokens, index }: Props) {
  const marked = (text: string) => highlight(text, highlightTokens).map((part, i) => part.mark ? <mark key={i}>{part.text}</mark> : <span key={i}>{part.text}</span>);
  return <a href={resourceLink(item.path)} className={`realm-resource-card${active ? ' selected' : ''}`} data-resource-index={index} aria-current={active ? 'true' : undefined} onClick={event => {
    if (event.button === 0 && !event.metaKey && !event.ctrlKey && !event.shiftKey && !event.altKey) { event.preventDefault(); onOpen(); }
  }}>
    <div className="realm-resource-tags">{item.top_kind && <span>{item.top_kind}</span>}{item.side && item.side !== item.top_kind && <span>{item.side}</span>}<span className="resource-open-mark" aria-hidden="true">↗</span></div>
    <strong title={item.title}>{marked(item.title)}</strong>
    <div className="realm-resource-snippet">{marked(item.snippet || item.category || '打开查看完整档案')}</div>
  </a>;
}
