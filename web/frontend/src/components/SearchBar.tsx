import { useEffect, useRef } from 'react';

type Props = {
  value: string;
  onChange: (v: string) => void;
  count: number;
  searching?: boolean;
  pinyinReady?: boolean;
  onClear: () => void;
};

export function SearchBar({ value, onChange, count, searching, pinyinReady, onClear }: Props) {
  const ref = useRef<HTMLInputElement>(null);
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        ref.current?.focus();
        ref.current?.select();
      }
      if (e.key === 'Escape' && document.activeElement === ref.current) {
        ref.current?.blur();
      }
    };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, []);

  return (
    <div className="searchbar">
      <div className="searchbar-icon">⌕</div>
      <input
        ref={ref}
        type="search"
        autoComplete="off"
        spellCheck={false}
        placeholder={pinyinReady ? '搜索：标题 / 拼音 / 路径 / 正文 / 多词 AND…' : '搜索：标题 / 路径 / 正文…'}
        value={value}
        onChange={e => onChange(e.target.value)}
      />
      {searching && <span className="searchbar-pulse" />}
      {value && (
        <button className="searchbar-clear" onClick={onClear} title="清空">×</button>
      )}
      <span className="searchbar-hint">
        <kbd>⌘</kbd>/<kbd>Ctrl</kbd> + <kbd>K</kbd>
      </span>
      <span className="searchbar-count">{value ? `${count} 命中` : `${count} 条`}</span>
    </div>
  );
}
