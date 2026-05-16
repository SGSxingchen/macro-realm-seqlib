import { FormEvent, useEffect, useMemo, useState } from 'react';
import { api } from '../api';
import { NormalizationReview, NormalizationReviewSummary } from '../types';
import { words } from '../utils';
import { ThemeToggle } from './ThemeToggle';

const reviewIdFromUrl = () => new URL(window.location.href).searchParams.get('id') || '';

function normalizeLines(text: string) {
  return text.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
}

function lineStats(before: string, after: string) {
  const a = normalizeLines(before);
  const b = normalizeLines(after);
  const max = Math.max(a.length, b.length);
  let changed = 0;
  for (let i = 0; i < max; i += 1) {
    if ((a[i] || '') !== (b[i] || '')) changed += 1;
  }
  return { before: a.length, after: b.length, changed };
}

export function NormalizationReviewPage() {
  const [items, setItems] = useState<NormalizationReviewSummary[]>([]);
  const [selectedId, setSelectedId] = useState(reviewIdFromUrl);
  const [review, setReview] = useState<NormalizationReview | null>(null);
  const [signer, setSigner] = useState('');
  const [note, setNote] = useState('');
  const [error, setError] = useState('');
  const [listLoading, setListLoading] = useState(true);
  const [detailLoading, setDetailLoading] = useState(false);
  const stats = useMemo(() => review ? lineStats(review.original_content, review.normalized_content) : null, [review]);

  const refreshList = () => {
    setListLoading(true);
    api<{ items: NormalizationReviewSummary[] }>('/api/normalization/reviews')
      .then(r => {
        setItems(r.items);
        setSelectedId(id => id || r.items[0]?.id || '');
      })
      .catch(e => setError(e instanceof Error ? e.message : String(e)))
      .finally(() => setListLoading(false));
  };

  useEffect(() => {
    refreshList();
  }, []);

  useEffect(() => {
    if (!selectedId) {
      setReview(null);
      return;
    }
    setDetailLoading(true);
    setError('');
    api<NormalizationReview>('/api/normalization/reviews/' + encodeURIComponent(selectedId))
      .then(setReview)
      .catch(e => setError(e instanceof Error ? e.message : String(e)))
      .finally(() => setDetailLoading(false));
  }, [selectedId]);

  const sign = async (e: FormEvent) => {
    e.preventDefault();
    if (!review) return;
    setError('');
    try {
      const next = await api<NormalizationReview>(
        `/api/normalization/reviews/${encodeURIComponent(review.id)}/sign`,
        { method: 'POST', body: JSON.stringify({ signer, note }) },
      );
      setReview(next);
      setNote('');
      refreshList();
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  };

  return (
    <main className="review-page">
      <header className="review-topbar">
        <div>
          <p className="eyebrow">规范化审核</p>
          <h1>人工审核总览</h1>
        </div>
        <div className="review-top-actions">
          <ThemeToggle />
          <div className="review-status">
            {items.filter(item => item.status === 'approved').length}/{items.length} 已通过
          </div>
        </div>
      </header>

      {error ? <section className="review-card error-box">{error}</section> : null}

      <section className="review-workbench">
        <aside className="review-card review-task-list">
          <div className="review-list-head">
            <h2>全部任务</h2>
            <button type="button" onClick={refreshList}>刷新</button>
          </div>
          {listLoading ? <p className="hint">正在读取审核任务。</p> : null}
          {!listLoading && items.length === 0 ? <p className="hint">还没有审核任务。批量规范脚本创建任务后，会出现在这里。</p> : null}
          <div className="review-list">
            {items.map(item => (
              <button
                key={item.id}
                type="button"
                className={`review-list-item ${selectedId === item.id ? 'active' : ''}`}
                onClick={() => setSelectedId(item.id)}
              >
                <span>
                  <strong>{item.title}</strong>
                  <small>{item.resource_path}</small>
                </span>
                <i className={item.status === 'approved' ? 'ok' : ''}>
                  {item.signature_count}/{item.required_signatures}
                </i>
              </button>
            ))}
          </div>
        </aside>

        <section className="review-detail">
          {detailLoading ? <section className="review-card"><p className="hint">正在读取审核内容。</p></section> : null}
          {!detailLoading && !review ? <section className="review-card"><p className="hint">请选择一个审核任务。</p></section> : null}

          {review ? (
            <>
              <section className="review-card review-meta">
                <div>
                  <span>资源路径</span>
                  <strong>{review.resource_path}</strong>
                </div>
                <div>
                  <span>行数变化</span>
                  <strong>{stats?.before} → {stats?.after}，差异行 {stats?.changed}</strong>
                </div>
                <div>
                  <span>字数变化</span>
                  <strong>{words(review.original_content)} → {words(review.normalized_content)}</strong>
                </div>
                <div>
                  <span>审核状态</span>
                  <strong>{review.status === 'approved' ? '已通过' : '待签名'}</strong>
                </div>
                {review.note ? <p>{review.note}</p> : null}
              </section>

              <section className="review-diff">
                <article className="review-pane">
                  <h2>规范化前</h2>
                  <pre>{review.original_content}</pre>
                </article>
                <article className="review-pane">
                  <h2>规范化后</h2>
                  <pre>{review.normalized_content}</pre>
                </article>
              </section>

              <section className="review-card review-sign">
                <div>
                  <h2>签名确认</h2>
                  <p className="hint">凑齐 1 个签名后，该任务即视为通过。</p>
                </div>
                <form onSubmit={sign}>
                  <input value={signer} onChange={e => setSigner(e.target.value)} placeholder="审核人署名" />
                  <input value={note} onChange={e => setNote(e.target.value)} placeholder="备注，可留空" />
                  <button type="submit" disabled={!signer.trim()}>签名通过</button>
                </form>
                {review.signatures.length > 0 ? (
                  <ul className="signature-list">
                    {review.signatures.map(s => (
                      <li key={`${s.signer}-${s.signed_at}`}>
                        <strong>{s.signer}</strong>
                        <span>{s.signed_at}</span>
                        {s.note ? <em>{s.note}</em> : null}
                      </li>
                    ))}
                  </ul>
                ) : null}
              </section>
            </>
          ) : null}
        </section>
      </section>
    </main>
  );
}
