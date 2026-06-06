// web/frontend/src/components/ui/Pagination.tsx
/** 传统页码条:‹ 1 2 3 … N ›。pageCount <= 1 时不渲染。 */
export function Pagination({ page, pageCount, onChange }: {
  page: number;
  pageCount: number;
  onChange: (page: number) => void;
}) {
  if (pageCount <= 1) return null;

  // 始终含首末页,当前页 ±1,间隙折叠为省略号
  const wanted = new Set<number>([1, pageCount, page - 1, page, page + 1]);
  const pages: (number | '…')[] = [];
  let prev = 0;
  for (let i = 1; i <= pageCount; i++) {
    if (!wanted.has(i)) continue;
    if (i - prev > 1) pages.push('…');
    pages.push(i);
    prev = i;
  }

  return (
    <nav className="pager" aria-label="分页">
      <button type="button" disabled={page <= 1} onClick={() => onChange(page - 1)}>‹</button>
      {pages.map((p, i) => p === '…'
        ? <span key={`gap-${i}`} className="pager-gap">…</span>
        : (
          <button
            key={p}
            type="button"
            className={p === page ? 'active' : ''}
            onClick={() => onChange(p)}
          >{p}</button>
        ))}
      <button type="button" disabled={page >= pageCount} onClick={() => onChange(page + 1)}>›</button>
    </nav>
  );
}
