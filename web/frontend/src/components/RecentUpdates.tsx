import { useEffect, useMemo, useState } from 'react';
import { api, buildQuery, routePath } from '../api';
import { ChangeDetail, ChangeKind, ChangeItem, GitChanges } from '../types';
import { kb } from '../utils';

const labels: Record<ChangeKind, string> = {
  added: '新增资源',
  modified: '更新资源',
  deleted: '下架/删除',
  renamed: '移动/改名',
};

const kindOrder: ChangeKind[] = ['added', 'modified', 'renamed', 'deleted'];

function detailKey(kind: ChangeKind, item: ChangeItem) {
  return `${kind}:${item.old_path || ''}:${item.path}`;
}

function isPublicResource(item: ChangeItem) {
  return item.root === '序列库' && item.path.startsWith('序列库/');
}

function normalizePublicChanges(data: GitChanges): GitChanges {
  const readable = {
    added: data.readable.added.filter(isPublicResource),
    modified: data.readable.modified.filter(isPublicResource),
    deleted: data.readable.deleted.filter(isPublicResource),
    renamed: data.readable.renamed.filter(isPublicResource),
  };
  const stats = {
    added: readable.added.length,
    modified: readable.modified.length,
    deleted: readable.deleted.length,
    renamed: readable.renamed.length,
    total: readable.added.length + readable.modified.length + readable.deleted.length + readable.renamed.length,
  };
  return { ...data, readable, stats };
}

type SplitRow = {
  kind: 'context' | 'change' | 'gap';
  oldNo?: number | null;
  newNo?: number | null;
  oldText?: string;
  newText?: string;
  oldType?: 'removed' | 'context' | 'empty';
  newType?: 'added' | 'context' | 'empty';
};

/** 把 unified diff 流配对成左右两栏。
 * - context  → 左右同文本
 * - 小段（同向 ≤ BLOCK_THRESHOLD 行）的 removed/added 连续段 → i-i 配对，看小修小补
 * - 大段（同向超过阈值，意味着大面积重写）→ 不配对，左侧整块红 + 右侧空白，再左侧空白 + 右侧整块绿
 *   这样大重写时不会强行让 25 行红和 25 行绿假装一一对应，避免误导
 * - gap → 两侧省略号
 */
const BLOCK_THRESHOLD = 6;

function toSplit(rows: ChangeDetail['rows']): SplitRow[] {
  const out: SplitRow[] = [];
  let i = 0;
  while (i < rows.length) {
    const r = rows[i];
    if (r.type === 'gap') {
      out.push({ kind: 'gap' });
      i++;
      continue;
    }
    if (r.type === 'context') {
      out.push({
        kind: 'context',
        oldNo: r.old_no, newNo: r.new_no,
        oldText: r.text || '', newText: r.text || '',
        oldType: 'context', newType: 'context',
      });
      i++;
      continue;
    }
    const removed: typeof rows = [];
    const added: typeof rows = [];
    while (i < rows.length && rows[i].type === 'removed') { removed.push(rows[i]); i++; }
    while (i < rows.length && rows[i].type === 'added') { added.push(rows[i]); i++; }

    const isLargeRewrite = removed.length >= BLOCK_THRESHOLD || added.length >= BLOCK_THRESHOLD;
    if (isLargeRewrite) {
      // 大段重写：先把 removed 整块靠左展示，右侧空白同高；再把 added 整块靠右展示
      for (const left of removed) {
        out.push({
          kind: 'change',
          oldNo: left.old_no ?? null, newNo: null,
          oldText: left.text ?? '', newText: '',
          oldType: 'removed', newType: 'empty',
        });
      }
      for (const right of added) {
        out.push({
          kind: 'change',
          oldNo: null, newNo: right.new_no ?? null,
          oldText: '', newText: right.text ?? '',
          oldType: 'empty', newType: 'added',
        });
      }
    } else {
      const max = Math.max(removed.length, added.length);
      for (let j = 0; j < max; j++) {
        const left = removed[j];
        const right = added[j];
        out.push({
          kind: 'change',
          oldNo: left?.old_no ?? null,
          newNo: right?.new_no ?? null,
          oldText: left?.text ?? '',
          newText: right?.text ?? '',
          oldType: left ? 'removed' : 'empty',
          newType: right ? 'added' : 'empty',
        });
      }
    }
  }
  return out;
}

function isHeavyRewrite(detail: ChangeDetail): boolean {
  // 改动占新版总行数的比例 >70% 视为重写
  const total = Math.max(detail.new_line_count, detail.old_line_count, 1);
  return (detail.additions + detail.deletions) / total > 0.7;
}

function DiffRows({ detail }: { detail: ChangeDetail }) {
  const rows = useMemo(() => toSplit(detail.rows), [detail.rows]);
  const heavy = isHeavyRewrite(detail);
  return (
    <div className="diff-box">
      <div className="diff-summary">
        <span className="add-count">+{detail.additions}</span>
        <span className="del-count">−{detail.deletions}</span>
        <span className="line-count">{detail.old_line_count} → {detail.new_line_count} 行</span>
        {heavy && <span className="rewrite-flag">大面积重写</span>}
        {detail.truncated && <b>内容较长，已截断显示</b>}
      </div>
      <div className="diff-split" role="table" aria-label="对比">
        <div className="diff-split-head">
          <span>原版本</span>
          <span>新版本</span>
        </div>
        {rows.length ? rows.map((row, i) => {
          if (row.kind === 'gap') {
            return (
              <div className="diff-split-row gap" key={i}>
                <div className="diff-side">…</div>
                <div className="diff-side">…</div>
              </div>
            );
          }
          return (
            <div className={`diff-split-row ${row.kind}`} key={i}>
              <div className={`diff-side side-${row.oldType}`}>
                <span className="line-no">{row.oldNo ?? ''}</span>
                <code>{row.oldText || (row.oldType === 'empty' ? '' : ' ')}</code>
              </div>
              <div className={`diff-side side-${row.newType}`}>
                <span className="line-no">{row.newNo ?? ''}</span>
                <code>{row.newText || (row.newType === 'empty' ? '' : ' ')}</code>
              </div>
            </div>
          );
        }) : <p className="no-change">没有文本差异。</p>}
      </div>
    </div>
  );
}

export function RecentUpdates({ onOpen }: { onOpen: (path: string) => void }) {
  const [data, setData] = useState<GitChanges | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [activeKey, setActiveKey] = useState('');
  const [activeKind, setActiveKind] = useState<ChangeKind | null>(null);
  const [activeItem, setActiveItem] = useState<ChangeItem | null>(null);
  const [details, setDetails] = useState<Record<string, ChangeDetail>>({});
  const [detailLoading, setDetailLoading] = useState(false);

  const load = async () => {
    setLoading(true);
    setError('');
    try {
      setData(normalizePublicChanges(await api<GitChanges>('/api/git/changes?public_only=true')));
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => { load(); }, []);

  const closeModal = () => {
    setActiveKey('');
    setActiveKind(null);
    setActiveItem(null);
  };

  useEffect(() => {
    if (!activeKey) return;
    const onKey = (e: KeyboardEvent) => { if (e.key === 'Escape') closeModal(); };
    window.addEventListener('keydown', onKey);
    return () => window.removeEventListener('keydown', onKey);
  }, [activeKey]);

  const title = useMemo(() => {
    if (!data) return '最近更新';
    return `${data.from_ref} 之后的公开资源更新`;
  }, [data]);

  const copy = () => {
    if (!data) return;
    navigator.clipboard?.writeText(data.markdown || data.text).catch(() => {});
  };

  const openModal = async (kind: ChangeKind, item: ChangeItem) => {
    if (!data) return;
    const key = detailKey(kind, item);
    setActiveKey(key);
    setActiveKind(kind);
    setActiveItem(item);
    if (details[key]) return;
    setDetailLoading(true);
    try {
      const qs = buildQuery({
        kind,
        from_ref: data.from_ref,
        old_path: item.old_path || undefined,
        public_only: true,
      });
      const detail = await api<ChangeDetail>(`/api/git/change-detail/${routePath(item.path)}${qs}`);
      setDetails(prev => ({ ...prev, [key]: detail }));
    } catch (e: unknown) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setDetailLoading(false);
    }
  };

  const activeDetail = activeKey ? details[activeKey] : null;
  const canOpen = !!(activeItem && activeKind && activeItem.exists && activeItem.root === '序列库' && activeKind !== 'deleted');

  return (
    <section className="updates-view terminal-scroll">
      <div className="updates-hero">
        <div>
          <p className="eyebrow">RELEASE DIFF</p>
          <h2>{title}</h2>
          <p>{data ? `共 ${data.stats.total} 项变化。点开任意条目查看左右对比。` : '正在读取版本变更。'}</p>
        </div>
        <div className="updates-actions">
          <button type="button" onClick={load} disabled={loading}>{loading ? '刷新中' : '刷新'}</button>
          <button type="button" onClick={copy} disabled={!data}>复制摘要</button>
        </div>
      </div>

      {error && <div className="notice-line">{error}</div>}
      {loading && !data && <div className="updates-loading">正在生成变更清单...</div>}

      {data && (
        <>
          <div className="updates-stats">
            <span><b>{data.stats.total}</b>总计</span>
            <span><b>{data.stats.added}</b>新增</span>
            <span><b>{data.stats.modified}</b>修改</span>
            <span><b>{data.stats.renamed}</b>移动</span>
            <span><b>{data.stats.deleted}</b>删除</span>
          </div>
          <div className="updates-groups">
            {kindOrder.filter(k => (data.readable[k]?.length || 0) > 0).map(kind => (
              <section className={`updates-group ${kind}`} key={kind}>
                <h3>{labels[kind]}<em>{data.readable[kind].length}</em></h3>
                <div className="updates-list">
                  {data.readable[kind].map(item => {
                    const key = detailKey(kind, item);
                    return (
                      <button
                        type="button"
                        className="update-item"
                        key={key}
                        onClick={() => openModal(kind, item)}
                      >
                        <b>{item.title}</b>
                        <span>{item.category || '根目录'}{typeof item.size === 'number' ? ` · ${kb(item.size)}` : ''}</span>
                        <small>{kind === 'renamed' ? `${item.old_path} → ${item.path}` : item.path}</small>
                      </button>
                    );
                  })}
                </div>
              </section>
            ))}
            {data.stats.total === 0 && <div className="updates-loading">本次无公开资源变更。</div>}
          </div>
        </>
      )}

      {activeItem && activeKind && (
        <div className="diff-modal-backdrop" onClick={closeModal}>
          <div className="diff-modal" onClick={e => e.stopPropagation()}>
            <header className="diff-modal-head">
              <div>
                <p className="eyebrow">{labels[activeKind]}</p>
                <h3>{activeItem.title}</h3>
                <small>{activeKind === 'renamed' ? `${activeItem.old_path} → ${activeItem.path}` : activeItem.path}</small>
              </div>
              <div className="diff-modal-actions">
                {canOpen && <button type="button" onClick={() => { onOpen(activeItem.path); closeModal(); }}>打开当前资源</button>}
                <button type="button" className="diff-modal-close" onClick={closeModal} aria-label="关闭">×</button>
              </div>
            </header>
            <div className="diff-modal-body">
              {detailLoading && !activeDetail && <div className="updates-loading">正在读取差异...</div>}
              {activeDetail && <DiffRows detail={activeDetail} />}
            </div>
          </div>
        </div>
      )}
    </section>
  );
}
