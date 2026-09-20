import { useEffect, useRef } from 'react';

/** A local disclosure, not another permanently reserved toolbar row. */
export function ReferencePicker({ path, title, options, onChange }: {
  path: string;
  title: string;
  options: { path: string; title: string }[];
  onChange: (path: string) => void;
}) {
  const details = useRef<HTMLDetailsElement>(null);
  const summary = useRef<HTMLElement>(null);

  useEffect(() => {
    const outside = (event: PointerEvent) => {
      if (details.current?.open && !details.current.contains(event.target as Node)) details.current.open = false;
    };
    document.addEventListener('pointerdown', outside);
    return () => document.removeEventListener('pointerdown', outside);
  }, []);

  const close = () => {
    if (details.current) details.current.open = false;
    summary.current?.focus({ preventScroll: true });
  };

  return <details ref={details} className="reader-reference-picker"
    onKeyDown={event => {
      if (event.key === 'Escape' && details.current?.open) {
        event.preventDefault(); event.stopPropagation(); close();
      }
    }}
    onBlur={event => {
      if (event.relatedTarget && !event.currentTarget.contains(event.relatedTarget as Node)) event.currentTarget.open = false;
    }}>
    <summary ref={summary} aria-label="更换参照档案" title={`固定参照：${title}`}>
      <span className="reader-reference-badge">参照</span>
      <span className="reader-reference-title">{title}</span>
      <span aria-hidden="true">▾</span>
    </summary>
    <div className="reader-reference-popover">
      <label>固定参照
        <select aria-label="选择参照档案" value={path} onChange={event => { onChange(event.target.value); close(); }}>
          {options.map(option => <option key={option.path} value={option.path}>{option.title}</option>)}
        </select>
      </label>
      <p>右侧保持固定，左侧通过标签切换。两份档案独立滚动，不强行对齐不同模板。</p>
      <button type="button" onClick={close}>完成</button>
    </div>
  </details>;
}
