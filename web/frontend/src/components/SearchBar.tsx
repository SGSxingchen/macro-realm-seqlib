import { useEffect, useRef, useState } from 'react';

type Props = { value: string; onChange: (value: string) => void; count: number; searching?: boolean; pinyinReady?: boolean; onClear: () => void };
export function SearchBar({ value, onChange, count, searching, pinyinReady, onClear }: Props) {
  const ref = useRef<HTMLInputElement>(null);
  const composing = useRef(false);
  const [draft, setDraft] = useState(value);
  useEffect(() => { if (!composing.current) setDraft(value); }, [value]);
  useEffect(() => {
    const onKey = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') { event.preventDefault(); ref.current?.focus(); ref.current?.select(); }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);
  return <div className="realm-search">
    <label className="realm-sr-only" htmlFor="realm-query">搜索档案标题或正文</label>
    <span aria-hidden="true">⌕</span>
    <input id="realm-query" ref={ref} type="search" autoComplete="off" spellCheck={false} placeholder={pinyinReady ? '名称、规则、拼音首字母…' : '搜索名称或规则正文…'} value={draft}
      onCompositionStart={() => { composing.current = true; }} onCompositionEnd={e => { composing.current = false; onChange(e.currentTarget.value); }}
      onChange={e => { setDraft(e.target.value); if (!composing.current) onChange(e.target.value); }}
      onKeyDown={e => {
        if (composing.current || e.nativeEvent.isComposing) return;
        if (e.key === 'Escape') ref.current?.blur();
        if (e.key === 'ArrowDown' || e.key === 'Enter') {
          const first = document.querySelector<HTMLAnchorElement>('[data-resource-index]');
          if (first) { e.preventDefault(); first.focus(); }
        }
      }} />
    {draft ? <button aria-label="清空关键词" onClick={() => { setDraft(''); onClear(); ref.current?.focus(); }}>×</button> : <kbd title="Ctrl / ⌘ + K">⌘ K</kbd>}
    <span className="realm-sr-only" role="status">{searching ? '检索中' : `${count} 个结果`}</span>
  </div>;
}
