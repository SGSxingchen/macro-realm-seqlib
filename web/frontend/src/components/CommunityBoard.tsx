import { useCallback, useEffect, useRef, useState } from 'react';
import { api, buildQuery } from '../api';
import { copyText } from '../library-state';
import type {
  CommunityHomeResponse,
  CommunityPost,
  CommunityThreadDetailResponse,
  CommunityViewerState,
} from '../types';

export function formatPostTime(isoOrTs: string | number | undefined): string {
  if (isoOrTs === undefined || isoOrTs === null || isoOrTs === '') return '';
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
  activeThreadId?: number;
  onSelectThread: (threadId: number | null) => void;
  onOpenAuth: () => void;
  onEnterLibrary: (query?: string, categoryPath?: string) => void;
  onOpenResource: (path: string, anchor?: string) => void;
};

export function CommunityBoard({
  viewer,
  recordCount,
  categories,
  activeThreadId,
  onSelectThread,
  onOpenAuth,
  onEnterLibrary,
  onOpenResource,
}: Props) {
  const [data, setData] = useState<CommunityHomeResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState('');
  const [sort, setSort] = useState<'latest' | 'created' | 'hot'>('latest');
  const [boardSearch, setBoardSearch] = useState('');
  const [appliedBoardSearch, setAppliedBoardSearch] = useState('');
  const [quickSeqQuery, setQuickSeqQuery] = useState('');
  const [mobileStream, setMobileStream] = useState<'board' | 'recent'>('board');
  const [notice, setNotice] = useState('');

  // 新建主题弹窗状态
  const newThreadDialogRef = useRef<HTMLDialogElement>(null);
  const [composerOpen, setComposerOpen] = useState(false);
  const [newTitle, setNewTitle] = useState('');
  const [newContent, setNewContent] = useState('');
  const [submittingThread, setSubmittingThread] = useState(false);

  // 主题详情（1F / 2F / 3F 楼层）状态
  const [threadDetail, setThreadDetail] = useState<CommunityPost | null>(null);
  const [threadLoading, setThreadLoading] = useState(false);
  const [threadError, setThreadError] = useState('');
  const [replyContent, setReplyContent] = useState('');
  const [replyQuote, setReplyQuote] = useState<CommunityPost | null>(null);
  const [submittingReply, setSubmittingReply] = useState(false);
  const replyInputRef = useRef<HTMLTextAreaElement>(null);

  const flash = (msg: string) => {
    setNotice(msg);
    window.setTimeout(() => setNotice(prev => (prev === msg ? '' : prev)), 3800);
  };

  useEffect(() => {
    const dialog = newThreadDialogRef.current;
    if (!dialog) return;
    if (composerOpen && !dialog.open) dialog.showModal();
    else if (!composerOpen && dialog.open) dialog.close();
  }, [composerOpen]);

  const loadHome = useCallback(async () => {
    setLoading(true);
    setError('');
    try {
      const qs = buildQuery({
        sort,
        q: appliedBoardSearch || undefined,
        limit: 45,
      });
      const res = await api<CommunityHomeResponse>('/api/community/home' + qs);
      setData(res);
    } catch (err) {
      setError(err instanceof Error ? err.message : '读取论坛列表失败');
    } finally {
      setLoading(false);
    }
  }, [sort, appliedBoardSearch]);

  const loadThreadDetail = useCallback(async (tid: number) => {
    setThreadLoading(true);
    setThreadError('');
    try {
      const res = await api<CommunityThreadDetailResponse>(`/api/community/threads/${tid}`);
      setThreadDetail(res.thread);
    } catch (err) {
      setThreadError(err instanceof Error ? err.message : '无法加载该主题帖');
      setThreadDetail(null);
    } finally {
      setThreadLoading(false);
    }
  }, []);

  useEffect(() => {
    loadHome();
  }, [loadHome, viewer?.authenticated, viewer?.is_moderator]);

  useEffect(() => {
    if (activeThreadId && activeThreadId > 0) {
      loadThreadDetail(activeThreadId);
    } else {
      setThreadDetail(null);
      setThreadError('');
      setReplyQuote(null);
      setReplyContent('');
    }
  }, [activeThreadId, loadThreadDetail, viewer?.authenticated, viewer?.is_moderator]);

  const openNewThreadModal = () => {
    if (!viewer?.authenticated) {
      onOpenAuth();
      return;
    }
    setComposerOpen(true);
  };

  const handleCreateThread = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!viewer?.authenticated) {
      onOpenAuth();
      return;
    }
    const titleTrim = newTitle.trim();
    const contentTrim = newContent.trim();
    if (!titleTrim || !contentTrim) {
      flash('请填写主题标题和正文内容');
      return;
    }
    setSubmittingThread(true);
    try {
      const res = await api<{ ok: boolean; post: CommunityPost }>('/api/community/board', {
        method: 'POST',
        body: JSON.stringify({
          title: titleTrim,
          content: contentTrim,
          reply_to_id: null,
        }),
      });
      setNewTitle('');
      setNewContent('');
      setComposerOpen(false);
      flash('新主题已发布');
      await loadHome();
      onSelectThread(res.post.id);
    } catch (err) {
      flash(err instanceof Error ? err.message : '发布失败');
    } finally {
      setSubmittingThread(false);
    }
  };

  const handleCreateReply = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!viewer?.authenticated) {
      onOpenAuth();
      return;
    }
    if (!threadDetail) return;
    const contentTrim = replyContent.trim();
    if (!contentTrim) return;
    setSubmittingReply(true);
    try {
      await api('/api/community/board', {
        method: 'POST',
        body: JSON.stringify({
          title: '',
          content: contentTrim,
          reply_to_id: replyQuote ? replyQuote.id : threadDetail.id,
        }),
      });
      setReplyContent('');
      setReplyQuote(null);
      flash('回帖已发送');
      await Promise.all([loadThreadDetail(threadDetail.id), loadHome()]);
    } catch (err) {
      flash(err instanceof Error ? err.message : '回帖失败');
    } finally {
      setSubmittingReply(false);
    }
  };

  const handleToggleLike = async (post: CommunityPost) => {
    if (!viewer?.authenticated) {
      onOpenAuth();
      return;
    }
    try {
      await api(`/api/community/posts/${post.id}/like`, { method: 'POST' });
      if (activeThreadId) await loadThreadDetail(activeThreadId);
      await loadHome();
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
      flash(post.is_pinned ? '已取消置顶' : '已置顶该主题');
      if (activeThreadId) await loadThreadDetail(activeThreadId);
      await loadHome();
    } catch (err) {
      flash(err instanceof Error ? err.message : '置顶操作失败');
    }
  };

  const handleDelete = async (post: CommunityPost, isRootThread = false) => {
    try {
      await api(`/api/community/posts/${post.id}`, { method: 'DELETE' });
      flash(isRootThread ? '主题帖已删除' : '该楼层已删除');
      if (isRootThread) {
        onSelectThread(null);
      } else if (activeThreadId) {
        await loadThreadDetail(activeThreadId);
      }
      await loadHome();
    } catch (err) {
      flash(err instanceof Error ? err.message : '删除失败');
    }
  };

  const handleQuickBan = async (post: CommunityPost) => {
    const targetQq = post.author.qq;
    if (!targetQq) return;
    const reason = window.prompt(`限制轮回者 ${post.author.nickname}（QQ: ${targetQq}）发言 72 小时，输入原因：`, '违反轮回者论坛规范');
    if (reason === null) return;
    try {
      await api('/api/community/moderation/ban', {
        method: 'POST',
        body: JSON.stringify({
          qq: targetQq,
          duration_hours: 72,
          reason: reason.trim() || '违反轮回者论坛规范',
          nickname_hint: post.author.nickname,
        }),
      });
      flash(`已禁言 ${post.author.nickname} 72 小时`);
    } catch (err) {
      flash(err instanceof Error ? err.message : '禁言失败');
    }
  };

  const shareCurrentThread = async (tid: number) => {
    const url = new URL(window.location.href);
    url.search = `?thread=${tid}`;
    url.hash = '';
    try {
      await copyText(url.href);
      flash('已复制本帖链接');
    } catch {
      flash('浏览器限制自动复制，请直接复制地址栏链接');
    }
  };

  const boardItems = data?.board.items || [];
  const recentComments = data?.recent_resource_comments || [];
  const stats = data?.stats;

  return (
    <section className="forum-page" aria-label="宏观界域轮回者论坛">
      <div className="forum-shell">
        {notice && (
          <div className="comm-inline-toast" role="status" aria-live="polite">
            {notice}
          </div>
        )}

        {!activeThreadId && (
          <div className="comm-mobile-tabs" role="tablist" aria-label="主页视图切换">
            <button
              type="button"
              role="tab"
              aria-selected={mobileStream === 'board'}
              className={mobileStream === 'board' ? 'active' : ''}
              onClick={() => setMobileStream('board')}
            >
              主题列表 ({data?.board.total_threads ?? 0})
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={mobileStream === 'recent'}
              className={mobileStream === 'recent' ? 'active' : ''}
              onClick={() => setMobileStream('recent')}
            >
              序列锐评与检索 ({recentComments.length})
            </button>
          </div>
        )}

        <div className={`forum-grid${activeThreadId ? ' is-reading-thread' : ` mobile-${mobileStream}`}`}>
          {/* ================= 左侧主区：主题列表 或 单帖楼层详情 ================= */}
          <div className="forum-main">
            {activeThreadId ? (
              /* ---------- 单帖楼层阅读视图 (1F / 2F / 3F ...) ---------- */
              <div className="forum-thread-view">
                <div className="forum-thread-topbar">
                  <button type="button" className="forum-back-btn" onClick={() => onSelectThread(null)}>
                    ← 返回主题列表
                  </button>
                  <div className="forum-thread-top-actions">
                    <button type="button" onClick={() => shareCurrentThread(activeThreadId)}>
                      分享本帖
                    </button>
                    {threadDetail?.can_pin && (
                      <button type="button" onClick={() => handleTogglePin(threadDetail)}>
                        {threadDetail.is_pinned ? '取消置顶' : '置顶主题'}
                      </button>
                    )}
                    {threadDetail?.can_delete && (
                      <button type="button" className="is-danger" onClick={() => handleDelete(threadDetail, true)}>
                        删除整帖
                      </button>
                    )}
                  </div>
                </div>

                {threadLoading && !threadDetail ? (
                  <div className="forum-state-box">正在展开主题帖…</div>
                ) : threadError || !threadDetail ? (
                  <div className="forum-state-box" role="alert">
                    <p>{threadError || '该主题不存在或已被删除'}</p>
                    <button type="button" onClick={() => onSelectThread(null)}>返回主题列表</button>
                  </div>
                ) : (
                  <>
                    <article className="forum-detail-box">
                      <header className="forum-detail-head">
                        <div className="forum-detail-title-row">
                          {threadDetail.is_pinned && <span className="comm-pin-tag">置顶</span>}
                          <h1>{threadDetail.title || '无题通讯'}</h1>
                        </div>
                        <div className="forum-detail-sub">
                          <span>共 {(threadDetail.reply_count ?? 0) + 1} 个楼层</span>
                          <span>发布于 {formatPostTime(threadDetail.created_ts)}</span>
                          {threadDetail.last_reply_ts && threadDetail.last_reply_ts > threadDetail.created_ts && (
                            <span>最后回复 {formatPostTime(threadDetail.last_reply_ts)}</span>
                          )}
                        </div>
                      </header>

                      {/* 1F 楼主 */}
                      <div className="forum-floor is-op" id={`floor-1`}>
                        <div className="forum-floor-head">
                          <img src={threadDetail.author.avatar_url} alt="" className="comm-avatar" referrerPolicy="no-referrer" />
                          <div className="forum-floor-meta">
                            <div className="comm-author-line">
                              <strong>{threadDetail.author.nickname}</strong>
                              <span className="forum-op-tag">楼主</span>
                              {threadDetail.author.badge && <em className="comm-role-badge">{threadDetail.author.badge}</em>}
                              <span className="comm-qq-tag">
                                {viewer?.is_moderator && threadDetail.author.qq ? `QQ ${threadDetail.author.qq}` : threadDetail.author.qq_masked}
                              </span>
                            </div>
                            <time dateTime={threadDetail.created_at}>{formatPostTime(threadDetail.created_ts)}</time>
                          </div>
                          <span className="forum-floor-num">1F</span>
                        </div>
                        <div className="forum-floor-body">{threadDetail.content}</div>
                        <div className="forum-floor-foot">
                          <button
                            type="button"
                            className={`comm-like-btn${threadDetail.liked_by_me ? ' is-liked' : ''}`}
                            onClick={() => handleToggleLike(threadDetail)}
                          >
                            ▲ 共鸣 {threadDetail.like_count > 0 ? threadDetail.like_count : ''}
                          </button>
                          <button
                            type="button"
                            onClick={() => {
                              if (!viewer?.authenticated) {
                                onOpenAuth();
                                return;
                              }
                              setReplyQuote(null);
                              replyInputRef.current?.focus();
                            }}
                          >
                            回复楼主
                          </button>
                          {viewer?.is_moderator && threadDetail.author.qq && threadDetail.author.qq !== '10000' && threadDetail.author.qq !== viewer.user?.qq && (
                            <button type="button" className="is-danger" onClick={() => handleQuickBan(threadDetail)}>
                              禁言
                            </button>
                          )}
                        </div>
                      </div>

                      {/* 2F..NF 跟帖楼层 */}
                      {(threadDetail.replies || []).map((rep, idx) => {
                        const floorNo = rep.floor || idx + 2;
                        const isOp = rep.author.qq_masked === threadDetail.author.qq_masked && rep.author.nickname === threadDetail.author.nickname;
                        return (
                          <div key={rep.id} className="forum-floor" id={`floor-${floorNo}`}>
                            <div className="forum-floor-head">
                              <img src={rep.author.avatar_url} alt="" className="comm-avatar" referrerPolicy="no-referrer" />
                              <div className="forum-floor-meta">
                                <div className="comm-author-line">
                                  <strong>{rep.author.nickname}</strong>
                                  {isOp && <span className="forum-op-tag">楼主</span>}
                                  {rep.author.badge && <em className="comm-role-badge">{rep.author.badge}</em>}
                                  <span className="comm-qq-tag">
                                    {viewer?.is_moderator && rep.author.qq ? `QQ ${rep.author.qq}` : rep.author.qq_masked}
                                  </span>
                                  {rep.reply_to_id && rep.reply_to_id !== threadDetail.id && rep.reply_to_nickname && (
                                    <span className="comm-reply-target">回复 @{rep.reply_to_nickname}</span>
                                  )}
                                </div>
                                <time dateTime={rep.created_at}>{formatPostTime(rep.created_ts)}</time>
                              </div>
                              <span className="forum-floor-num">{floorNo}F</span>
                            </div>
                            <div className="forum-floor-body">{rep.content}</div>
                            <div className="forum-floor-foot">
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
                                  setReplyQuote(rep);
                                  replyInputRef.current?.focus();
                                }}
                              >
                                回复此层
                              </button>
                              {rep.can_delete && (
                                <button type="button" className="is-danger" onClick={() => handleDelete(rep, false)}>
                                  删除
                                </button>
                              )}
                              {viewer?.is_moderator && rep.author.qq && rep.author.qq !== viewer.user?.qq && (
                                <button type="button" className="is-danger" onClick={() => handleQuickBan(rep)}>
                                  禁言
                                </button>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </article>

                    {/* 底部回帖框 */}
                    <form className="forum-reply-box" onSubmit={handleCreateReply}>
                      <div className="forum-reply-head">
                        <strong>参与回帖</strong>
                        {!viewer?.authenticated && (
                          <button type="button" className="comm-primary-btn" onClick={onOpenAuth}>
                            QQ 验证登录
                          </button>
                        )}
                      </div>
                      {replyQuote && (
                        <div className="comm-reply-banner">
                          <span>
                            正在回复 <b>{replyQuote.floor ? `${replyQuote.floor}F ` : ''}@{replyQuote.author.nickname}</b>：
                            {replyQuote.content.slice(0, 45)}
                            {replyQuote.content.length > 45 ? '…' : ''}
                          </span>
                          <button type="button" onClick={() => setReplyQuote(null)}>取消指定楼层 ×</button>
                        </div>
                      )}
                      <textarea
                        ref={replyInputRef}
                        rows={4}
                        maxLength={4000}
                        disabled={!viewer?.authenticated || Boolean(viewer.user?.is_banned)}
                        placeholder={
                          !viewer?.authenticated
                            ? '使用交易群内 QQ 私聊验证码登录后即可跟帖回复…'
                            : viewer.user?.is_banned
                              ? `发言受限：${viewer.user.ban_reason || '请联系管理员'}`
                              : replyQuote
                                ? `回复 @${replyQuote.author.nickname}…`
                                : '撰写你的回帖内容…'
                        }
                        value={replyContent}
                        onChange={e => setReplyContent(e.target.value)}
                      />
                      <div className="comm-composer-foot">
                        <span className="comm-char-count">{replyContent.length} / 4000</span>
                        {viewer?.authenticated ? (
                          <button
                            type="submit"
                            className="comm-primary-btn"
                            disabled={submittingReply || !replyContent.trim() || Boolean(viewer.user?.is_banned)}
                          >
                            {submittingReply ? '发送中…' : '发送回帖'}
                          </button>
                        ) : (
                          <button type="button" onClick={onOpenAuth}>登录后回帖</button>
                        )}
                      </div>
                    </form>
                  </>
                )}
              </div>
            ) : (
              /* ---------- 经典论坛紧凑主题列表视图 ---------- */
              <div className="forum-list-panel">
                <header className="forum-toolbar">
                  <div className="forum-title-group">
                    <h1>轮回者终端论坛</h1>
                    <div className="forum-sort-tabs" role="group" aria-label="主题排序">
                      <button type="button" aria-pressed={sort === 'latest'} onClick={() => setSort('latest')}>
                        最新回复
                      </button>
                      <button type="button" aria-pressed={sort === 'created'} onClick={() => setSort('created')}>
                        最新发布
                      </button>
                      <button type="button" aria-pressed={sort === 'hot'} onClick={() => setSort('hot')}>
                        高赞共鸣
                      </button>
                    </div>
                  </div>

                  <div className="forum-toolbar-right">
                    <form
                      className="forum-search-inline"
                      onSubmit={e => {
                        e.preventDefault();
                        setAppliedBoardSearch(boardSearch.trim());
                      }}
                    >
                      <input
                        type="search"
                        aria-label="搜索论坛帖子"
                        placeholder="搜帖子标题或作者…"
                        value={boardSearch}
                        onChange={e => {
                          setBoardSearch(e.target.value);
                          if (!e.target.value.trim() && appliedBoardSearch) setAppliedBoardSearch('');
                        }}
                      />
                    </form>
                    <button type="button" className="comm-primary-btn forum-new-btn" onClick={openNewThreadModal}>
                      + 发布新主题
                    </button>
                  </div>
                </header>

                {error && (
                  <div className="comm-error-box" role="alert">
                    <span>{error}</span>
                    <button type="button" onClick={loadHome}>重试</button>
                  </div>
                )}

                {loading && !data ? (
                  <div className="forum-state-box">正在读取主题列表…</div>
                ) : boardItems.length === 0 ? (
                  <div className="forum-state-box">
                    <p>{appliedBoardSearch ? `没有找到包含「${appliedBoardSearch}」的主题帖。` : '当前还没有主题帖，点击右上角「+ 发布新主题」开启第一帖吧。'}</p>
                  </div>
                ) : (
                  <div className="forum-topic-table" role="list">
                    {boardItems.map(item => {
                      const replyCnt = item.reply_count ?? 0;
                      const isSameBody = !item.excerpt || item.excerpt === item.title;
                      return (
                        <div
                          key={item.id}
                          role="listitem"
                          className={`forum-topic-row${item.is_pinned ? ' is-pinned' : ''}`}
                          data-post-id={item.id}
                        >
                          <img
                            src={item.author.avatar_url}
                            alt=""
                            className="forum-row-avatar"
                            referrerPolicy="no-referrer"
                          />
                          <div className="forum-row-main">
                            <div className="forum-row-titleline">
                              {item.is_pinned && <span className="comm-pin-tag">置顶</span>}
                              {item.author.badge === '系统公告' && <span className="forum-sys-tag">公告</span>}
                              <button
                                type="button"
                                className="forum-topic-link"
                                onClick={() => onSelectThread(item.id)}
                              >
                                {item.title || '无题通讯'}
                              </button>
                            </div>
                            {!isSameBody && (
                              <p className="forum-row-excerpt" onClick={() => onSelectThread(item.id)}>
                                {item.excerpt}
                              </p>
                            )}
                            <div className="forum-row-meta">
                              <span className="forum-meta-author">{item.author.nickname}</span>
                              {item.author.badge && item.author.badge !== '系统公告' && (
                                <em className="comm-role-badge">{item.author.badge}</em>
                              )}
                              <span>发布于 {formatPostTime(item.created_ts)}</span>
                              {replyCnt > 0 && item.last_reply_nickname && (
                                <span className="forum-meta-last">
                                  · 最后由 <b>{item.last_reply_nickname}</b> 回复于 {formatPostTime(item.last_reply_ts || item.created_ts)}
                                </span>
                              )}
                              {item.like_count > 0 && (
                                <span className="forum-meta-likes">▲ {item.like_count}</span>
                              )}
                            </div>
                          </div>

                          <div className="forum-row-right">
                            {viewer?.is_moderator && (
                              <div className="forum-row-mod">
                                <button type="button" onClick={() => handleTogglePin(item)} title={item.is_pinned ? '取消置顶' : '置顶'}>
                                  {item.is_pinned ? '取消置顶' : '置顶'}
                                </button>
                                <button type="button" className="is-danger" onClick={() => handleDelete(item, false)} title="删除主题">
                                  删
                                </button>
                              </div>
                            )}
                            <button
                              type="button"
                              className={`forum-reply-pill${replyCnt > 0 ? ' has-replies' : ''}`}
                              onClick={() => onSelectThread(item.id)}
                              aria-label={`${replyCnt} 条回复`}
                            >
                              {replyCnt}
                            </button>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                )}
              </div>
            )}
          </div>

          {/* ================= 右侧边栏：身份卡 + 序列终端速查 + 最新档案评论 ================= */}
          <aside className="forum-sidebar" aria-label="轮回者状态与序列库入口">
            {/* 1. 轮回者状态与发帖卡 */}
            <div className="forum-side-box">
              {viewer?.authenticated && viewer.user ? (
                <div className="forum-user-box">
                  <div className="forum-user-row">
                    <img src={viewer.user.avatar_url} alt="" className="comm-avatar" referrerPolicy="no-referrer" />
                    <div className="forum-user-info">
                      <div className="comm-author-line">
                        <strong>{viewer.user.nickname}</strong>
                        {viewer.user.badge && <em className="comm-role-badge">{viewer.user.badge}</em>}
                      </div>
                      <span>{viewer.user.qq_masked} · 角色卡 {viewer.user.card_count} 份</span>
                    </div>
                  </div>
                </div>
              ) : (
                <div className="forum-guest-box">
                  <strong>宏观界域 · 轮回者交流区</strong>
                  <p>验证交易群（{viewer?.group_id || 672204828}）QQ 与角色卡资格后，可发帖招募、研讨构筑与评价强化序列。</p>
                  <button type="button" className="comm-primary-btn forum-full-btn" onClick={onOpenAuth}>
                    QQ 验证登录
                  </button>
                </div>
              )}

              {stats && (
                <div className="forum-stat-grid">
                  <div>
                    <b>{stats.board_thread_count ?? data?.board.total_threads ?? 0}</b>
                    <span>主题帖</span>
                  </div>
                  <div>
                    <b>{stats.resource_comment_count}</b>
                    <span>序列评论</span>
                  </div>
                  <div>
                    <b>{stats.member_count}</b>
                    <span>轮回者</span>
                  </div>
                </div>
              )}
            </div>

            {/* 2. 强化序列库速查入口 */}
            <div className="forum-side-box">
              <div className="comm-side-head">
                <h2>强化序列库速查</h2>
                <button type="button" onClick={() => onEnterLibrary('')}>
                  进入序列库 ({recordCount}) →
                </button>
              </div>
              <form
                className="forum-seq-search"
                onSubmit={e => {
                  e.preventDefault();
                  onEnterLibrary(quickSeqQuery.trim());
                }}
              >
                <input
                  type="search"
                  aria-label="快速检索强化序列"
                  placeholder="检索职业、技能表、特质…"
                  value={quickSeqQuery}
                  onChange={e => setQuickSeqQuery(e.target.value)}
                />
                <button type="submit">检索</button>
              </form>
              {categories.length > 0 && (
                <div className="forum-cat-grid">
                  {categories.slice(0, 6).map(cat => (
                    <button
                      key={cat.path}
                      type="button"
                      className="forum-cat-btn"
                      onClick={() => onEnterLibrary('', cat.path)}
                    >
                      <span>{cat.name}</span>
                      <b>{cat.count}</b>
                    </button>
                  ))}
                </div>
              )}
            </div>

            {/* 3. 全站最新序列档案评论动态 */}
            <div className="forum-side-box">
              <div className="comm-side-head">
                <h2>最新序列评论</h2>
                <button type="button" onClick={() => onEnterLibrary('')}>查阅全部 →</button>
              </div>
              {recentComments.length === 0 ? (
                <p className="comm-side-empty">
                  暂无序列评论。在「序列检索」打开任意强化档案即可针对单张能力卡发表实战评价。
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
                        onClick={() =>
                          onOpenResource(
                            c.resource_path,
                            c.anchor_id || (c.anchor_label ? `label:${c.anchor_label}` : '')
                          )
                        }
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

      {/* ================= 发布新主题弹窗 ================= */}
      <dialog
        ref={newThreadDialogRef}
        className="comm-auth-dialog forum-composer-dialog"
        aria-labelledby="forum-new-thread-title"
        onCancel={e => {
          e.preventDefault();
          setComposerOpen(false);
        }}
        onClick={e => {
          if (e.target === e.currentTarget) setComposerOpen(false);
        }}
      >
        <form className="comm-auth-card" onSubmit={handleCreateThread}>
          <header className="comm-auth-head">
            <div>
              <span className="comm-kicker">轮回者终端论坛</span>
              <h2 id="forum-new-thread-title">发布新主题</h2>
            </div>
            <button type="button" className="comm-dialog-close" onClick={() => setComposerOpen(false)} aria-label="关闭发帖窗口">
              ×
            </button>
          </header>

          <label className="comm-field-label" htmlFor="forum-input-title">
            主题标题
          </label>
          <input
            id="forum-input-title"
            className="forum-title-input"
            type="text"
            maxLength={100}
            placeholder="简明扼要写出你的主题（如：求评神秘侧天师后期构筑 / 周六晚跑团招募）…"
            value={newTitle}
            onChange={e => setNewTitle(e.target.value)}
          />

          <label className="comm-field-label" htmlFor="forum-input-body">
            正文内容
          </label>
          <textarea
            id="forum-input-body"
            className="forum-body-textarea"
            rows={7}
            maxLength={4000}
            placeholder="输入完整讨论内容、构筑细节、战报记录或招募要求…"
            value={newContent}
            onChange={e => setNewContent(e.target.value)}
          />

          <div className="comm-composer-foot">
            <span className="comm-char-count">标题 {newTitle.length}/100 · 正文 {newContent.length}/4000</span>
            <div className="comm-auth-actions" style={{ marginTop: 0 }}>
              <button type="button" onClick={() => setComposerOpen(false)}>取消</button>
              <button
                type="submit"
                className="comm-primary-btn"
                disabled={submittingThread || !newTitle.trim() || !newContent.trim()}
              >
                {submittingThread ? '发布中…' : '发布主题'}
              </button>
            </div>
          </div>
        </form>
      </dialog>
    </section>
  );
}
