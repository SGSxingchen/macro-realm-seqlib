import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { api, buildQuery } from '../../api';
import type { CommunityPost, CommunityThreadsResponse, CommunityViewerState } from '../../types';
import { formatPostTime } from '../CommunityBoard';
import type { Section } from './adaptive';

export type AnchorTarget = {
  id: string;
  label: string;
};

type Props = {
  resourcePath: string;
  sections: Section[];
  viewer?: CommunityViewerState | null;
  onOpenAuth?: () => void;
  onJumpToAnchor: (anchorId: string) => void;
  pendingAnchor?: AnchorTarget | null;
  onClearPendingAnchor?: () => void;
  onAnchorCountsChange?: (counts: Record<string, number>) => void;
};

function cleanCardLabel(title: string): string {
  return title.replace(/[（(\[【]\s*(?:EX|[SABCDEF][+-]?)\s*级?\s*[）)\]】]/gi, '').trim() || title;
}

export function ResourceComments({
  resourcePath,
  sections,
  viewer,
  onOpenAuth,
  onJumpToAnchor,
  pendingAnchor,
  onClearPendingAnchor,
  onAnchorCountsChange,
}: Props) {
  const [data, setData] = useState<CommunityThreadsResponse | null>(null);
  const [loading, setLoading] = useState(false);
  const [sort, setSort] = useState<'latest' | 'hot' | 'oldest'>('latest');
  const [filterAnchor, setFilterAnchor] = useState<{ id: string; label: string } | null>(null);
  const [selectedAnchorId, setSelectedAnchorId] = useState('');
  const [replyTarget, setReplyTarget] = useState<CommunityPost | null>(null);
  const [content, setContent] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [notice, setNotice] = useState('');
  const textareaRef = useRef<HTMLTextAreaElement>(null);

  const anchorOptions = useMemo<AnchorTarget[]>(() => {
    return sections
      .filter(s => s.kind !== 'intro')
      .map(s => ({
        id: s.id,
        label: cleanCardLabel(s.title),
      }));
  }, [sections]);

  const anchorMap = useMemo(() => new Map(anchorOptions.map(a => [a.id, a.label])), [anchorOptions]);
  const labelToIdMap = useMemo(() => new Map(anchorOptions.map(a => [a.label, a.id])), [anchorOptions]);

  const resolveLiveAnchorId = useCallback((anchorId: string, anchorLabel: string) => {
    if (anchorId && anchorMap.has(anchorId)) return anchorId;
    const clean = cleanCardLabel(anchorLabel || '');
    if (clean && labelToIdMap.has(clean)) return labelToIdMap.get(clean)!;
    return anchorId;
  }, [anchorMap, labelToIdMap]);

  const loadComments = useCallback(async () => {
    setLoading(true);
    try {
      const qs = buildQuery({
        path: resourcePath,
        anchor_id: filterAnchor?.id || undefined,
        anchor_label: filterAnchor?.label || undefined,
        sort,
        limit: 60,
      });
      const res = await api<CommunityThreadsResponse>('/api/community/comments' + qs);
      setData(res);
      if (onAnchorCountsChange) {
        const counts: Record<string, number> = {};
        const labelCounts = res.label_counts || {};
        for (const opt of anchorOptions) {
          const byLabel = labelCounts[opt.label] || 0;
          const byId = res.anchor_counts?.[opt.id]?.count || 0;
          const finalCnt = Math.max(byLabel, byId);
          if (finalCnt > 0) counts[opt.id] = finalCnt;
        }
        onAnchorCountsChange(counts);
      }
    } catch {
      /* 在静态预览或无后端 mock 场景下安静降级为空评论列表，不干扰正文阅读 */
      setData(prev => prev ?? { items: [], count: 0, total_threads: 0, total_posts: 0, anchor_counts: {}, limit: 60, offset: 0 });
    } finally {
      setLoading(false);
    }
  }, [resourcePath, filterAnchor, sort, anchorOptions, onAnchorCountsChange]);

  useEffect(() => {
    setFilterAnchor(null);
    setSelectedAnchorId('');
    setReplyTarget(null);
    setContent('');
    setData(null);
  }, [resourcePath]);

  useEffect(() => {
    loadComments();
  }, [loadComments, viewer?.authenticated, viewer?.is_moderator]);

  // 当用户点击某张能力卡上的「评论此卡」时，自动绑定锚点并聚焦输入框
  useEffect(() => {
    if (!pendingAnchor) return;
    setSelectedAnchorId(pendingAnchor.id);
    onClearPendingAnchor?.();
    requestAnimationFrame(() => {
      textareaRef.current?.focus({ preventScroll: false });
    });
  }, [pendingAnchor, onClearPendingAnchor]);

  const flash = (msg: string) => {
    setNotice(msg);
    window.setTimeout(() => setNotice(prev => (prev === msg ? '' : prev)), 3500);
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!viewer?.authenticated) {
      onOpenAuth?.();
      return;
    }
    const text = content.trim();
    if (!text) return;
    setSubmitting(true);
    try {
      const anchorLabel = selectedAnchorId ? (anchorMap.get(selectedAnchorId) || selectedAnchorId) : '';
      await api('/api/community/comments', {
        method: 'POST',
        body: JSON.stringify({
          resource_path: resourcePath,
          content: text,
          anchor_id: selectedAnchorId,
          anchor_label: anchorLabel,
          reply_to_id: replyTarget ? replyTarget.id : null,
        }),
      });
      setContent('');
      setReplyTarget(null);
      flash('评论已送达');
      await loadComments();
    } catch (err) {
      flash(err instanceof Error ? err.message : '评论发送失败');
    } finally {
      setSubmitting(false);
    }
  };

  const handleToggleLike = async (post: CommunityPost) => {
    if (!viewer?.authenticated) {
      onOpenAuth?.();
      return;
    }
    try {
      await api(`/api/community/posts/${post.id}/like`, { method: 'POST' });
      await loadComments();
    } catch (err) {
      flash(err instanceof Error ? err.message : '点赞失败');
    }
  };

  const handleTogglePin = async (post: CommunityPost) => {
    try {
      await api(`/api/community/posts/${post.id}/pin`, {
        method: 'POST',
        body: JSON.stringify({ pinned: !post.is_pinned }),
      });
      flash(post.is_pinned ? '已取消置顶' : '已置顶该评论');
      await loadComments();
    } catch (err) {
      flash(err instanceof Error ? err.message : '置顶操作失败');
    }
  };

  const handleDelete = async (post: CommunityPost) => {
    try {
      await api(`/api/community/posts/${post.id}`, { method: 'DELETE' });
      flash('评论已删除');
      await loadComments();
    } catch (err) {
      flash(err instanceof Error ? err.message : '删除失败');
    }
  };

  const handleQuickBan = async (post: CommunityPost) => {
    const targetQq = post.author.qq;
    if (!targetQq) return;
    const reason = window.prompt(`禁言 ${post.author.nickname}（QQ: ${targetQq}）72 小时，输入原因：`, '不当评论');
    if (reason === null) return;
    try {
      await api('/api/community/moderation/ban', {
        method: 'POST',
        body: JSON.stringify({
          qq: targetQq,
          duration_hours: 72,
          reason: reason.trim() || '不当评论',
          nickname_hint: post.author.nickname,
        }),
      });
      flash(`已对 ${post.author.nickname} 执行 72 小时禁言`);
    } catch (err) {
      flash(err instanceof Error ? err.message : '禁言失败');
    }
  };

  const threads = data?.items || [];
  const anchorCounts = data?.anchor_counts || {};
  const activeAnchorTags = Object.entries(anchorCounts);

  return (
    <section id="resource-comments-panel" className="comm-resource-section" aria-label="档案讨论与能力卡锐评">
      <header className="comm-res-head">
        <div>
          <span className="comm-kicker">实战反馈与规则研讨</span>
          <h2>序列评论区 <span>{data?.total_posts ?? 0}</span></h2>
        </div>
        <div className="comm-sort-switch" role="group" aria-label="评论排序">
          <button type="button" aria-pressed={sort === 'latest'} onClick={() => setSort('latest')}>最新</button>
          <button type="button" aria-pressed={sort === 'hot'} onClick={() => setSort('hot')}>高赞</button>
          <button type="button" aria-pressed={sort === 'oldest'} onClick={() => setSort('oldest')}>最早</button>
        </div>
      </header>

      {activeAnchorTags.length > 0 && (
        <div className="comm-anchor-filter-bar" aria-label="按能力卡锚点筛选评论">
          <button
            type="button"
            className={!filterAnchor ? 'active' : ''}
            onClick={() => setFilterAnchor(null)}
          >
            全部评论 ({data?.total_posts ?? 0})
          </button>
          {activeAnchorTags.map(([key, info]) => {
            const label = info.label || anchorMap.get(key) || key;
            const liveId = resolveLiveAnchorId(info.anchor_id || key, label);
            const isActive = Boolean(filterAnchor && (filterAnchor.label === label || filterAnchor.id === liveId));
            return (
              <button
                key={key}
                type="button"
                className={isActive ? 'active' : ''}
                onClick={() => setFilterAnchor(prev => (prev && prev.label === label ? null : { id: liveId, label }))}
              >
                #{label} ({info.count})
              </button>
            );
          })}
        </div>
      )}

      {notice && <div className="comm-inline-toast" role="status">{notice}</div>}

      <form className="comm-composer is-compact" onSubmit={handleSubmit}>
        <div className="comm-composer-head">
          {viewer?.authenticated && viewer.user ? (
            <div className="comm-composer-user">
              <img src={viewer.user.avatar_url} alt="" className="comm-avatar-sm" referrerPolicy="no-referrer" />
              <span><b>{viewer.user.nickname}</b></span>
              {viewer.user.badge && <em className="comm-role-badge">{viewer.user.badge}</em>}
            </div>
          ) : (
            <div className="comm-composer-guest">
              <span>登录后可针对整份序列或具体能力卡发表评论与点赞</span>
              {onOpenAuth && (
                <button type="button" className="comm-primary-btn" onClick={onOpenAuth}>
                  QQ 验证登录
                </button>
              )}
            </div>
          )}

          {anchorOptions.length > 0 && (
            <label className="comm-anchor-picker">
              <span>关联能力卡：</span>
              <select
                aria-label="选择要关联的能力卡或章节锚点"
                value={selectedAnchorId}
                onChange={e => setSelectedAnchorId(e.target.value)}
              >
                <option value="">整份序列（不指定单卡）</option>
                {anchorOptions.map(opt => (
                  <option key={opt.id} value={opt.id}>
                    {opt.label}
                  </option>
                ))}
              </select>
            </label>
          )}
        </div>

        {replyTarget && (
          <div className="comm-reply-banner">
            <span>
              回复 <b>@{replyTarget.author.nickname}</b>：{replyTarget.content.slice(0, 38)}
              {replyTarget.content.length > 38 ? '…' : ''}
            </span>
            <button type="button" onClick={() => setReplyTarget(null)}>取消回复 ×</button>
          </div>
        )}

        <textarea
          ref={textareaRef}
          rows={3}
          maxLength={1600}
          disabled={!viewer?.authenticated || Boolean(viewer.user?.is_banned)}
          placeholder={
            !viewer?.authenticated
              ? '通过交易群 QQ 私聊验证码登录后即可参与本序列研讨…'
              : viewer.user?.is_banned
                ? `发言受限：${viewer.user.ban_reason || '请联系管理员'}`
                : selectedAnchorId
                  ? `针对「${anchorMap.get(selectedAnchorId) || '所选能力卡'}」写下实战评价、构筑联动或疑问…`
                  : '写下对本序列的实战体验、轮回构筑搭配或规则讨论…'
          }
          value={content}
          onChange={e => setContent(e.target.value)}
        />

        <div className="comm-composer-foot">
          <span className="comm-char-count">{content.length} / 1600</span>
          {viewer?.authenticated ? (
            <button
              type="submit"
              className="comm-primary-btn"
              disabled={submitting || !content.trim() || Boolean(viewer.user?.is_banned)}
            >
              {submitting ? '发送中…' : replyTarget ? '发送回复' : '发表评论'}
            </button>
          ) : (
            onOpenAuth && <button type="button" onClick={onOpenAuth}>登录后评论</button>
          )}
        </div>
      </form>

      {loading && !data ? (
        <div className="comm-empty-state">正在读取评论…</div>
      ) : threads.length === 0 ? (
        <div className="comm-empty-state">
          {filterAnchor ? '该能力卡锚点下暂无评论。' : '本序列暂未收到评论，欢迎分享第一条实战笔记。'}
        </div>
      ) : (
        <div className="comm-thread-list">
          {threads.map(post => (
            <article key={post.id} className={`comm-post-card${post.is_pinned ? ' is-pinned' : ''}`}>
              <header className="comm-post-head">
                <img src={post.author.avatar_url} alt="" className="comm-avatar" referrerPolicy="no-referrer" />
                <div className="comm-post-meta">
                  <div className="comm-author-line">
                    {post.is_pinned && <span className="comm-pin-tag">置顶</span>}
                    <strong>{post.author.nickname}</strong>
                    {post.author.badge && <em className="comm-role-badge">{post.author.badge}</em>}
                    <span className="comm-qq-tag">
                      {viewer?.is_moderator && post.author.qq ? `QQ ${post.author.qq}` : post.author.qq_masked}
                    </span>
                    {(post.anchor_id || post.anchor_label) && (
                      <button
                        type="button"
                        className="comm-anchor-pill"
                        onClick={() => onJumpToAnchor(resolveLiveAnchorId(post.anchor_id, post.anchor_label))}
                        title="点击定位到正文对应能力卡"
                      >
                        #{post.anchor_label || anchorMap.get(post.anchor_id) || '定位能力卡'} ↵
                      </button>
                    )}
                  </div>
                  <time dateTime={post.created_at}>{formatPostTime(post.created_ts)}</time>
                </div>
                <div className="comm-post-actions">
                  <button
                    type="button"
                    className={`comm-like-btn${post.liked_by_me ? ' is-liked' : ''}`}
                    aria-pressed={post.liked_by_me}
                    onClick={() => handleToggleLike(post)}
                  >
                    ▲ 共鸣 {post.like_count > 0 ? post.like_count : ''}
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      if (!viewer?.authenticated) {
                        onOpenAuth?.();
                        return;
                      }
                      setReplyTarget(post);
                      if (post.anchor_id) setSelectedAnchorId(post.anchor_id);
                      textareaRef.current?.focus();
                    }}
                  >
                    回复
                  </button>
                  {post.can_pin && (
                    <button type="button" onClick={() => handleTogglePin(post)}>
                      {post.is_pinned ? '取消置顶' : '置顶'}
                    </button>
                  )}
                  {post.can_delete && (
                    <button type="button" className="is-danger" onClick={() => handleDelete(post)}>
                      删除
                    </button>
                  )}
                  {viewer?.is_moderator && post.author.qq && post.author.qq !== viewer.user?.qq && (
                    <button type="button" className="is-danger" onClick={() => handleQuickBan(post)}>
                      禁言
                    </button>
                  )}
                </div>
              </header>
              <div className="comm-post-body">{post.content}</div>

              {post.replies && post.replies.length > 0 && (
                <div className="comm-replies">
                  {post.replies.map(rep => (
                    <div key={rep.id} className="comm-reply-item">
                      <div className="comm-reply-head">
                        <strong>{rep.author.nickname}</strong>
                        {rep.author.badge && <em className="comm-role-badge">{rep.author.badge}</em>}
                        {rep.reply_to_nickname && (
                          <span className="comm-reply-target">回复 @{rep.reply_to_nickname}</span>
                        )}
                        <time dateTime={rep.created_at}>{formatPostTime(rep.created_ts)}</time>
                        <div className="comm-reply-actions">
                          <button
                            type="button"
                            className={`comm-like-btn${rep.liked_by_me ? ' is-liked' : ''}`}
                            onClick={() => handleToggleLike(rep)}
                          >
                            ▲ {rep.like_count > 0 ? rep.like_count : '共鸣'}
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              if (!viewer?.authenticated) {
                                onOpenAuth?.();
                                return;
                              }
                              setReplyTarget(rep);
                              textareaRef.current?.focus();
                            }}
                          >
                            回复
                          </button>
                          {rep.can_delete && (
                            <button type="button" className="is-danger" onClick={() => handleDelete(rep)}>
                              删除
                            </button>
                          )}
                        </div>
                      </div>
                      <div className="comm-reply-body">{rep.content}</div>
                    </div>
                  ))}
                </div>
              )}
            </article>
          ))}
        </div>
      )}
    </section>
  );
}
