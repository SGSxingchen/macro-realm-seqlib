import { useState } from 'react';
import type { Section } from './adaptive';

export function MarkedText({ text, query }: { text: string; query: string }) {
  if (!query.trim()) return <>{text}</>;
  const needle = query.toLocaleLowerCase();
  const haystack = text.toLocaleLowerCase();
  const nodes = [];
  let start = 0;
  let match = haystack.indexOf(needle);
  while (match !== -1) {
    nodes.push(text.slice(start, match));
    nodes.push(<mark data-reader-match key={match}>{text.slice(match, match + query.length)}</mark>);
    start = match + query.length;
    match = haystack.indexOf(needle, start);
  }
  nodes.push(text.slice(start));
  return <>{nodes}</>;
}

export function AdaptiveSection({ section, query, rawMode, onCopy, onShare, idPrefix = '' }: {
  section: Section; query: string; rawMode: boolean; idPrefix?: string;
  onCopy: (raw: string) => void; onShare: (id: string) => void;
}) {
  const [showSource, setShowSource] = useState(false);
  return (
    <section id={idPrefix + section.id} data-section-id={section.id} className={`archive-section archive-${section.kind}${rawMode ? ' archive-raw' : ''}`} aria-label={section.title}>
      {!rawMode && section.kind !== 'intro' && (
        <header className="archive-section-bar">
          <span>{section.kind === 'entry' ? '能力 / 条目' : '章节 / 记录'}</span>
          <div>
            <button type="button" onClick={() => onCopy(section.raw)}>复制原文</button>
            <button type="button" onClick={() => onShare(section.id)} aria-label={`分享定位：${section.title}`}>定位链接</button>
            <button type="button" aria-expanded={showSource} onClick={() => setShowSource(v => !v)}>{showSource ? '关闭对照' : '原文对照'}</button>
          </div>
        </header>
      )}
      {rawMode ? <pre className="archive-source"><MarkedText text={section.raw} query={query} /></pre> : (
        <div className="archive-parts">
          {section.parts.map((part, i) => part.field ? (
            <div className={`archive-field${/名称$/.test(part.field.key) ? ' is-name' : ''}${/效果|限制|要求|条件|副作用/.test(part.field.key) ? ' is-rule' : ''}`} key={i}>
              <div className="archive-field-label"><MarkedText text={part.field.key} query={query} /></div>
              <div className="archive-field-value"><MarkedText text={part.field.value} query={query} /></div>
            </div>
          ) : <div className="archive-prose" key={i}><MarkedText text={part.raw} query={query} /></div>)}
        </div>
      )}
      {!rawMode && showSource && <pre className="archive-source source-compare">{section.raw}</pre>}
    </section>
  );
}
