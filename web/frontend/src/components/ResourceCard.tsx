import { Resource } from '../types';
import { kb, highlight, stamp } from '../utils';

type Props = {
  item: Resource;
  active: boolean;
  onOpen: () => void;
  highlightTokens: string[];
};

export function ResourceCard({ item, active, onOpen, highlightTokens }: Props) {
  const titleParts = highlight(item.title, highlightTokens);
  return (
    <button className={active ? 'res-card active' : 'res-card'} onClick={onOpen}>
      <div className="res-card-row">
        <span className="file-id">SEQ-{stamp(item.path)}</span>
        {item.side && <span className="res-tag tag-side">{item.side}</span>}
        {item.top_kind && item.top_kind !== item.side && <span className="res-tag tag-kind">{item.top_kind}</span>}
      </div>
      <b>{titleParts.map((p, i) => p.mark ? <mark key={i}>{p.text}</mark> : <span key={i}>{p.text}</span>)}</b>
      {item.snippet && (
        <small className="res-snippet">
          {highlight(item.snippet, highlightTokens).map((p, i) => p.mark ? <mark key={i}>{p.text}</mark> : <span key={i}>{p.text}</span>)}
        </small>
      )}
      <small className="res-path">{item.path}</small>
      <span className="card-meta">
        <em>{item.category || '根目录'}</em>
        <em>{kb(item.size)}</em>
      </span>
    </button>
  );
}
