import { useCallback, useEffect, useState } from 'react';
import { api, buildQuery } from '../api';
import type { CommunityHomeResponse, CommunityPost, CommunityViewerState } from '../types';

export function formatPostTime(isoOrTs: string | number): string {
  const ms = typeof isoOrTs === 'number' ? isoOrTs * 1000 : Date.parse(isoOrTs);
  if (Number.isNaN(ms)) return '';
  const diffSec = Math.floor((Date.now() - ms) / 1000);
  if (diffSec < 60) return '刚刚';
  if (diffSec < 3600) return `${Math.floor(diffSec / 60)} 分钟前`;
  if (diffSec < 86400) return `${Math.floor(diffSec / 3600)} 小时前`;
  const d = new Date(ms);
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hh = String(d.getHours()).padStart(2, '0');
  const mm = String(d.getMinutes()).padStart(2, '0');
  return y === new Date().getFullYear() ? `${m}-${day} ${hh}:${mm}` : `${y}-${m}-${day}`;
}

export function cleanSeqTitle(title: string): string {
  return title.replace(/^\d{1,4}】\s*/, '').trim() || title;
}

type Props = {
  viewer: CommunityViewerState | null;
  recordCount: number;
  categories: Array<{ name: string; path: string; count: number }>;
  onOpenAuth: () => void;
  onEnterLibrary: (query?: string, categoryPath?: string) => void;
  onOpenResource: (path: string, anchor?: string) => void;
};

export function CommunityBoard({
  viewer,
  recordCount,
  categories,
  onOpenAuth,
  onEnterLibrary,
  onOpenResource,
}: Props) {
  const [data, setData] = useState<CommunityHomeResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [sort, setSort] = useState<'latest' | 'hot'>('latest');
  const [quickQuery, setQuickQuery] = useState('');
  const [content, setContent] = useState('');
  const [replyTarget, setReplyTarget] = useState<CommunityPost | null>(null);
  const [submitting, setSubmitting] = useState(false);
  const [notice, setNotice] = useState('');
  const [mobileStream, setMobileStream] = useState<'board' | 'recent'>('board');

  const flash = (msg: string) => {
    setNotice(msg);
    window.setTimeout(() => setNotice(prev => (prev === msg ? '' : prev)), 4000);
  };

  const loadHome = useCallback(async (activeSort: 'latest' | 'hot' = sort) => {
    setLoading(true);
    setError('');
    try {
      const res = await api<CommunityHomeResponse>('/api/community/home' + buildQuery({ sort: activeSort, limit: 35 }));
      setData(res);
    } catch (err) {
      setError(err instanceof Error ? err.message : '留言板读取失败');
    } finally {
      setLoading(false);
    }
  }, [sort]);

  useEffect(() => {
    loadHome(sort);
  }, [sort, viewer?.authenticated, viewer?.is_moderator, loadHome]);

  const submitPost = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!viewer?.authenticated) {
      onOpenAuth();
      return;
    }
    const text = content.trim();
    if (!text) return;
    setSubmitting(true);
    try {
      await api('/api/community/board', {
        method: 'POST',
        body: JSON.stringify({
          content: text,
          reply_to_id: replyTarget ? replyTarget.id : null,
        }),
      });
      setContent('');
      setReplyTarget(null);
      flash('留言已发布');
      await loadHome(sort);
    } catch (err) {
      flash(err instanceof Error ? err.message : '发送失败');
    } finally {
      setSubmitting(false);
    }
  };

  const handleToggleLike = async (post: CommunityPost) => {
    if (!viewer?.authenticated) {
      onOpenAuth();
      return;
    }
    try {
      await api(`/api/community/posts/${post.id}/like`, { method: 'POST' });
      await loadHome(sort);
    } catch (err) {
      flash(err instanceof Error ? err.message : '操作失败');
    }
  };

  const handleTogglePin = async (post: CommunityPost) => {
    try {
      await api(`/api/community/posts/${post.id}/pin`, {
        method: 'POST',
        body: JSON.stringify({ pinned: !post.is_pinned }),
      });
      flash(post.is_pinned ? '已取消置顶' : '已设为置顶');
      await loadHome(sort);
    } catch (err) {
      flash(err instanceof Error ? err.message : '置顶操作失败');
    }
  };

  const handleDelete = async (post: CommunityPost) => {
    try {
      await api(`/api/community/posts/${post.id}`, { method: 'DELETE' });
      flash('已删除该条发言');
      await loadHome(sort);
    } catch (err) {
      flash(err instanceof Error ? err.message : '删除失败');
    }
  };

  const handleQuickBan = async (post: CommunityPost) => {
    const targetQq = post.author.qq;
    if (!targetQq) return;
    const reason = window.prompt(`限制 ${post.author.nickname}（QQ: ${targetQq}）发言，输入原因（留空使用默认说明）：`, '违反界域留言规范');
    if (reason === null) return;
    try {
      await api('/api/community/moderation/ban', {
        method: 'POST',
        body: JSON.stringify({
          qq: targetQq,
          duration_hours: 72,
          reason: reason.trim() || '违反界域留言规范',
          nickname_hint: post.author.nickname,
        }),
      });
      flash(`已对 ${post.author.nickname} 执行 72 小时禁言（可在后台解除或改为永久）`);
    } catch (err) {
      flash(err instanceof Error ? err.message : '禁言操作失败');
    }
  };

  const boardThreads = data?.board.items || [];
  const recentComments = data?.recent_resource_comments || [];
  const stats = data?.stats;

  return (
    <section className="comm-board-page" aria-label="界域主页留言板">
      <div className="comm-board-container">
        {/* 顶部主横幅：快速检索序列 + 终端状态带 */}
        <header className="comm-hero-bar">
          <div className="comm-hero-main">
            <div className="comm-kicker">
              宏观界域 · 调查员联络终端 {recordCount > 0 ? `· 现存 ${recordCount} 项强化序列` : ''}
            </div>
            <h1>界域公共留言板</h1>
            <p>交流跑团战报、序列构筑心得、模组情报与规则答疑；点击右侧动态可直接跳转至对应序列与能力卡锚点。</p>
          </div>
          <form
            className="comm-hero-search"
            onSubmit={e => {
              e.preventDefault();
              onEnterLibrary(quickQuery.trim());
            }}
          >
            <input
              type="search"
              aria-label="快速检索强化序列"
              placeholder="直接检索序列名称、词条、侧别或作者…"
              value={quickQuery}
              onChange={e => setQuickQuery(e.target.value)}
            />
            <button type="submit" className="comm-primary-btn">检索序列</button>
            <button type="button" onClick={() => onEnterLibrary('')}>进入序列库 →</button>
          </form>
          {categories.length > 0 && (
            <div className="comm-cat-strip" aria-label="序列分区直达">
              {categories.slice(0, 6).map(cat => (
                <button
                  key={cat.path}
                  type="button"
                  className="comm-cat-chip"
                  onClick={() => onEnterLibrary('', cat.path)}
                >
                  <span>{cat.name}</span>
                  <b>{cat.count}</b>
                </button>
              ))}
              {stats && (
                <span className="comm-stat-summary">
                  公共留言 <b>{stats.board_post_count}</b> · 档案评论 <b>{stats.resource_comment_count}</b> · 已认证调查员 <b>{stats.member_count}</b>
                </span>
              )}
            </div>
          )}
        </header>

        {notice && (
          <div className="comm-inline-toast" role="status" aria-live="polite">
            {notice}
          </div>
        )}

        <div className="comm-mobile-tabs" role="tablist" aria-label="主页栏目切换">
          <button
            type="button"
            role="tab"
            aria-selected={mobileStream === 'board'}
            className={mobileStream === 'board' ? 'active' : ''}
            onClick={() => setMobileStream('board')}
          >
            公共留言板 ({data?.board.total_posts ?? 0})
          </button>
          <button
            type="button"
            role="tab"
            aria-selected={mobileStream === 'recent'}
            className={mobileStream === 'recent' ? 'active' : ''}
            onClick={() => setMobileStream('recent')}
          >
            最新序列评论 ({recentComments.length})
          </button>
        </div>

        <div className={`comm-board-grid mobile-${mobileStream}`}>
          {/* 左主栏：发帖框 + 公共留言流 */}
          <div className="comm-board-main">
            <form className="comm-composer" onSubmit={submitPost}>
              <div className="comm-composer-head">
                {viewer?.authenticated && viewer.user ? (
                  <div className="comm-composer-user">
                    <img src={viewer.user.avatar_url} alt="" className="comm-avatar-sm" referrerPolicy="no-referrer" />
                    <span>
                      以 <b>{viewer.user.nickname}</b>（{viewer.user.qq_masked}）身份发言
                    </span>
                    {viewer.user.badge && <em className="comm-role-badge">{viewer.user.badge}</em>}
                  </div>
                ) : (
                  <div className="comm-composer-guest">
                    <span>当前为访客只读模式 · 群成员通过 QQ 私聊验证码登录后可留言与评论</span>
                    <button type="button" className="comm-primary-btn" onClick={onOpenAuth}>
                      QQ 验证登录
                    </button>
                  </div>
                )}
              </div>

              {replyTarget && (
                <div className="comm-reply-banner">
                  <span>
                    正在回复 <b>@{replyTarget.author.nickname}</b>：{replyTarget.content.slice(0, 42)}
                    {replyTarget.content.length > 42 ? '…' : ''}
                  </span>
                  <button type="button" onClick={() => setReplyTarget(null)}>取消回复 ×</button>
                </div>
              )}

              <textarea
                rows={3}
                maxLength={1200}
                disabled={!viewer?.authenticated || viewer.user?.is_banned}
                placeholder={
                  !viewer?.authenticated
                    ? '请先点击右上角「QQ 验证登录」验证主群成员与角色卡身份…'
                    : viewer.user?.is_banned
                      ? `当前账号已被限制发言：${viewer.user.ban_reason || '请联系管理员'}`
                      : replyTarget
                        ? `回复 @${replyTarget.author.nickname}…`
                        : '在界域留言板写下你的想法、构筑讨论、招募或跑团感想…'
                }
                value={content}
                onChange={e => setContent(e.target.value)}
              />

              <div className="comm-composer-foot">
                <span className="comm-char-count">{content.length} / 1200</span>
                {viewer?.authenticated ? (
                  <button
                    type="submit"
                    className="comm-primary-btn"
                    disabled={submitting || !content.trim() || Boolean(viewer.user?.is_banned)}
                  >
                    {submitting ? '发送中…' : replyTarget ? '发送回复' : '发布留言'}
                  </button>
                ) : (
                  <button type="button" onClick={onOpenAuth}>登录后发言</button>
                )}
              </div>
            </form>

            <div className="comm-stream-toolbar">
              <h2>公共留言 ({data?.board.total_posts ?? 0})</h2>
              <div className="comm-sort-switch" role="group" aria-label="留言排序">
                <button
                  type="button"
                  aria-pressed={sort === 'latest'}
                  onClick={() => setSort('latest')}
                >
                  最新
                </button>
                <button
                  type="button"
                  aria-pressed={sort === 'hot'}
                  onClick={() => setSort('hot')}
                >
                  高赞共鸣
                </button>
              </div>
            </div>

            {error && (
              <div className="comm-error-box" role="alert">
                <span>{error}</span>
                <button type="button" onClick={() => loadHome(sort)}>重试</button>
              </div>
            )}

            {loading && !data ? (
              <div className="comm-empty-state">正在同步界域留言板…</div>
            ) : boardThreads.length === 0 ? (
              <div className="comm-empty-state">
                暂无公共留言。验证群内 QQ 后即可留下第一条调查员通讯。
              </div>
            ) : (
              <div className="comm-thread-list">
                {boardThreads.map(post => (
                  <article
                    key={post.id}
                    className={`comm-post-card${post.is_pinned ? ' is-pinned' : ''}`}
                    data-post-id={post.id}
                  >
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
                              onOpenAuth();
                              return;
                            }
                            setReplyTarget(post);
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
                        {viewer?.is_moderator && post.author.qq && post.author.qq !== '10000' && post.author.qq !== viewer.user?.qq && (
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
                                      onOpenAuth();
                                      return;
                                    }
                                    setReplyTarget(rep);
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
          </div>

          {/* 右侧栏：全站最新序列档案评论动态流 */}
          <aside className="comm-board-side" aria-label="最新序列评论动态">
            <div className="comm-side-panel">
              <div className="comm-side-head">
                <h2>最新序列档案评论</h2>
                <button type="button" onClick={() => onEnterLibrary('')}>全部序列 →</button>
              </div>
              {recentComments.length === 0 ? (
                <p className="comm-side-empty">
                  各序列档案下暂未收到评论。打开任意职业、技能表或特质即可针对整篇或单张能力卡发表锐评。
                </p>
              ) : (
                <div className="comm-recent-feed">
                  {recentComments.map(c => (
                    <div key={c.id} className="comm-feed-item">
                      <div className="comm-feed-meta">
                        <strong>{c.author.nickname}</strong>
                        <time dateTime={c.created_at}>{formatPostTime(c.created_ts)}</time>
                      </div>
                      <button
                        type="button"
                        className="comm-feed-target"
                        onClick={() => onOpenResource(c.resource_path, c.anchor_id || '')}
                      >
                        <span className="comm-feed-doc">{cleanSeqTitle(c.resource_title || c.resource_path)}</span>
                        {c.anchor_label && <span className="comm-feed-anchor">#{c.anchor_label}</span>}
                      </button>
                      <p className="comm-feed-excerpt">{c.content}</p>
                    </div>
                  ))}
                </div>
              )}
            </div>
          </aside>
        </div>
      </div>
    </section>
  );
}
