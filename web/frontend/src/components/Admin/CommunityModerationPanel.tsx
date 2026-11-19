import { useCallback, useEffect, useState } from 'react';
import { api, buildQuery } from '../../api';
import type {
  AdminModerationPostsResponse,
  AdminModerationUser,
  AdminModerationUsersResponse,
  CommunityPost,
} from '../../types';
import { formatPostTime } from '../CommunityBoard';
import { StatusBadge } from '../ui/StatusBadge';

type Props = {
  onOpenResource?: (path: string, anchor?: string) => void;
};

export function CommunityModerationPanel({ onOpenResource }: Props) {
  const [subTab, setSubTab] = useState<'posts' | 'users'>('posts');
  const [postsData, setPostsData] = useState<AdminModerationPostsResponse | null>(null);
  const [usersData, setUsersData] = useState<AdminModerationUsersResponse | null>(null);
  const [scope, setScope] = useState<'all' | 'board' | 'resource'>('all');
  const [statusFilter, setStatusFilter] = useState<'active' | 'deleted' | 'all'>('active');
  const [userFilter, setUserFilter] = useState<'all' | 'banned' | 'moderator'>('all');
  const [query, setQuery] = useState('');
  const [loading, setLoading] = useState(false);
  const [notice, setNotice] = useState('');

  // 置顶公告表单
  const [annTitle, setAnnTitle] = useState('');
  const [announcement, setAnnouncement] = useState('');
  const [annNickname, setAnnNickname] = useState('界域终端公告');

  // 手动封禁指定 QQ 表单
  const [banQq, setBanQq] = useState('');
  const [banHours, setBanHours] = useState('72');
  const [banReason, setBanReason] = useState('');

  const flash = (msg: string) => {
    setNotice(msg);
    window.setTimeout(() => setNotice(prev => (prev === msg ? '' : prev)), 4000);
  };

  const loadPosts = useCallback(async () => {
    setLoading(true);
    try {
      const qs = buildQuery({ scope, status: statusFilter, q: query.trim() || undefined, limit: 80 });
      const res = await api<AdminModerationPostsResponse>('/api/community/moderation/posts' + qs);
      setPostsData(res);
    } catch (err) {
      flash(err instanceof Error ? err.message : '读取发言列表失败');
    } finally {
      setLoading(false);
    }
  }, [scope, statusFilter, query]);

  const loadUsers = useCallback(async () => {
    setLoading(true);
    try {
      const qs = buildQuery({ filter_mode: userFilter, q: query.trim() || undefined, limit: 100 });
      const res = await api<AdminModerationUsersResponse>('/api/community/moderation/users' + qs);
      setUsersData(res);
    } catch (err) {
      flash(err instanceof Error ? err.message : '读取成员列表失败');
    } finally {
      setLoading(false);
    }
  }, [userFilter, query]);

  useEffect(() => {
    if (subTab === 'posts') loadPosts();
    else loadUsers();
  }, [subTab, loadPosts, loadUsers]);

  const handlePublishAnnouncement = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!announcement.trim()) return;
    try {
      await api('/api/community/moderation/announcement', {
        method: 'POST',
        body: JSON.stringify({
          title: annTitle.trim(),
          content: announcement.trim(),
          nickname: annNickname.trim() || '界域终端公告',
          pinned: true,
        }),
      });
      setAnnTitle('');
      setAnnouncement('');
      flash('置顶公告已发布到轮回者论坛顶部');
      await loadPosts();
    } catch (err) {
      flash(err instanceof Error ? err.message : '发布公告失败');
    }
  };

  const handleManualBan = async (e: React.FormEvent) => {
    e.preventDefault();
    const cleaned = banQq.trim();
    if (!cleaned || cleaned.length < 5) {
      flash('请输入有效的账号标识（数字 QQ 号或 oid:... / OpenID）');
      return;
    }
    const hoursNum = Number(banHours);
    try {
      await api('/api/community/moderation/ban', {
        method: 'POST',
        body: JSON.stringify({
          qq: cleaned,
          duration_hours: hoursNum <= 0 ? 0 : hoursNum,
          unban: false,
          reason: banReason.trim() || '后台手动限制发言',
        }),
      });
      setBanQq('');
      setBanReason('');
      flash(`已对账号 ${cleaned} 执行禁言`);
      await loadUsers();
    } catch (err) {
      flash(err instanceof Error ? err.message : '封禁操作失败');
    }
  };

  const handleUnban = async (qq: string) => {
    try {
      await api('/api/community/moderation/ban', {
        method: 'POST',
        body: JSON.stringify({ qq, unban: true }),
      });
      flash(`已解除 QQ ${qq} 的禁言`);
      await loadUsers();
    } catch (err) {
      flash(err instanceof Error ? err.message : '解除禁言失败');
    }
  };

  const handleToggleRole = async (u: AdminModerationUser) => {
    const nextRole = u.role === 'moderator' ? 'member' : 'moderator';
    try {
      await api(`/api/community/moderation/users/${encodeURIComponent(u.qq)}/role`, {
        method: 'POST',
        body: JSON.stringify({ role: nextRole }),
      });
      flash(`已将 ${u.nickname}（${u.qq}）设为${nextRole === 'moderator' ? '社区协管' : '普通成员'}`);
      await loadUsers();
    } catch (err) {
      flash(err instanceof Error ? err.message : '角色修改失败');
    }
  };

  const handlePinPost = async (p: CommunityPost) => {
    try {
      await api(`/api/community/posts/${p.id}/pin`, {
        method: 'POST',
        body: JSON.stringify({ pinned: !p.is_pinned }),
      });
      flash(p.is_pinned ? '已取消置顶' : '已设为置顶');
      await loadPosts();
    } catch (err) {
      flash(err instanceof Error ? err.message : '置顶操作失败');
    }
  };

  const handleDeletePost = async (p: CommunityPost) => {
    try {
      await api(`/api/community/posts/${p.id}`, { method: 'DELETE' });
      flash(`已删除 #${p.id}`);
      await loadPosts();
    } catch (err) {
      flash(err instanceof Error ? err.message : '删除失败');
    }
  };

  const stats = postsData?.stats;

  return (
    <div className="console-grid">
      {notice && <div className="notice-line" style={{ gridColumn: '1 / -1' }}>{notice}</div>}

      {/* 社区数据概览与发置顶公告 */}
      <section className="console-card">
        <h3>社区状态与统计</h3>
        {stats ? (
          <div className="kv-grid">
            <span>主页公共留言</span><b>{stats.board_post_count} 条</b>
            <span>序列档案评论</span><b>{stats.resource_comment_count} 条</b>
            <span>已登录群成员</span><b>{stats.member_count} 人</b>
            <span>当前禁言账号</span><b>{stats.banned_count} 人</b>
          </div>
        ) : (
          <p className="hint">正在读取社区统计…</p>
        )}
        <hr style={{ border: 0, borderTop: '1px solid var(--line)', margin: '14px 0' }} />
        <h4>手动封禁 / 禁言指定账号</h4>
        <form onSubmit={handleManualBan} style={{ display: 'grid', gap: '8px' }}>
          <input
            placeholder="目标账号标识（如 oid:XXXX... 或旧数字 QQ 号）"
            value={banQq}
            onChange={e => setBanQq(e.target.value.slice(0, 80))}
          />
          <div className="row">
            <select value={banHours} onChange={e => setBanHours(e.target.value)}>
              <option value="24">禁言 24 小时</option>
              <option value="72">禁言 3 天（72h）</option>
              <option value="168">禁言 7 天</option>
              <option value="0">永久限制发言</option>
            </select>
            <input
              placeholder="原因备注（选填）"
              value={banReason}
              onChange={e => setBanReason(e.target.value)}
            />
          </div>
          <button type="submit" className="danger">执行禁言</button>
        </form>
      </section>

      <section className="console-card">
        <h3>发布论坛置顶公告</h3>
        <p className="hint">无需绑定个人 QQ，直接以系统身份在轮回者论坛顶部置顶展示。</p>
        <form onSubmit={handlePublishAnnouncement} style={{ display: 'grid', gap: '8px' }}>
          <div className="row">
            <input
              placeholder="公告标题（如：轮回者论坛版规与交易群说明）"
              value={annTitle}
              onChange={e => setAnnTitle(e.target.value)}
            />
            <input
              placeholder="署名（默认：界域终端公告）"
              value={annNickname}
              onChange={e => setAnnNickname(e.target.value)}
            />
          </div>
          <textarea
            rows={4}
            maxLength={4000}
            placeholder="输入要置顶的公告正文（如版本更新提示、跑团招募导航、版规说明）…"
            value={announcement}
            onChange={e => setAnnouncement(e.target.value)}
          />
          <button type="submit" disabled={!announcement.trim()}>发布并置顶</button>
        </form>
      </section>

      {/* 留言/评论审计 与 成员/黑名单列表 */}
      <section className="console-card wide">
        <div className="section-head" style={{ flexWrap: 'wrap', gap: '10px' }}>
          <div className="row">
            <button
              type="button"
              className={subTab === 'posts' ? 'active' : ''}
              onClick={() => setSubTab('posts')}
            >
              发言与评论管理 {postsData ? `(${postsData.total})` : ''}
            </button>
            <button
              type="button"
              className={subTab === 'users' ? 'active' : ''}
              onClick={() => setSubTab('users')}
            >
              登录成员与封禁名单 {usersData ? `(${usersData.total})` : ''}
            </button>
          </div>
          <div className="row">
            <input
              type="search"
              placeholder="搜正文 / 昵称 / QQ / 序列名…"
              value={query}
              onChange={e => setQuery(e.target.value)}
            />
            <button type="button" onClick={() => (subTab === 'posts' ? loadPosts() : loadUsers())}>
              {loading ? '查询中…' : '刷新'}
            </button>
          </div>
        </div>

        {subTab === 'posts' ? (
          <>
            <div className="row" style={{ marginBottom: '12px', flexWrap: 'wrap' }}>
              <select value={scope} onChange={e => setScope(e.target.value as 'all' | 'board' | 'resource')}>
                <option value="all">全部版块（留言板 + 序列评论）</option>
                <option value="board">仅主页留言板</option>
                <option value="resource">仅序列档案评论</option>
              </select>
              <select value={statusFilter} onChange={e => setStatusFilter(e.target.value as 'active' | 'deleted' | 'all')}>
                <option value="active">正常显示中</option>
                <option value="deleted">已删除归档</option>
                <option value="all">全部状态</option>
              </select>
            </div>

            <div className="comm-admin-table">
              {(postsData?.items || []).length === 0 ? (
                <p className="hint">当前筛选条件下没有记录。</p>
              ) : (
                postsData!.items.map(item => (
                  <div key={item.id} className="comm-admin-row">
                    <div className="comm-admin-row-head">
                      <span>
                        <b>#{item.id}</b> · [{item.scope === 'board' ? '主页留言' : '序列评论'}] ·{' '}
                        <strong>{item.author.nickname}</strong>（QQ: <code>{item.author.qq}</code>）
                      </span>
                      <span>
                        {item.is_pinned && <StatusBadge ok>置顶</StatusBadge>}
                        {item.status === 'deleted' && <StatusBadge ok={false} warn>已删 ({item.deleted_by})</StatusBadge>}
                        <small style={{ marginLeft: '8px' }}>{formatPostTime(item.created_ts)}</small>
                      </span>
                    </div>
                    {item.scope === 'resource' && (
                      <div className="comm-admin-target">
                        所属档案：
                        {onOpenResource ? (
                          <button
                            type="button"
                            className="comm-link-inline"
                            onClick={() => onOpenResource(item.resource_path, item.anchor_id || '')}
                          >
                            {item.resource_title || item.resource_path}
                            {item.anchor_label ? ` #${item.anchor_label}` : ''} ↗
                          </button>
                        ) : (
                          <span>{item.resource_title || item.resource_path}</span>
                        )}
                      </div>
                    )}
                    {item.title && item.scope === 'board' && !item.parent_id && (
                      <div style={{ fontWeight: 600, fontSize: '13.5px' }}>标题：{item.title}</div>
                    )}
                    <pre className="comm-admin-content">{item.raw_content || item.content}</pre>
                    <div className="row">
                      {item.status === 'active' && !item.parent_id && (
                        <button type="button" onClick={() => handlePinPost(item)}>
                          {item.is_pinned ? '取消置顶' : '设为置顶'}
                        </button>
                      )}
                      {item.status === 'active' && (
                        <button type="button" className="danger" onClick={() => handleDeletePost(item)}>
                          删除发言
                        </button>
                      )}
                      {item.author.qq && item.author.qq !== '10000' && (
                        <button
                          type="button"
                          onClick={() => {
                            setBanQq(item.author.qq || '');
                            setSubTab('users');
                          }}
                        >
                          填入封禁框
                        </button>
                      )}
                    </div>
                  </div>
                ))
              )}
            </div>
          </>
        ) : (
          <>
            <div className="row" style={{ marginBottom: '12px' }}>
              <select value={userFilter} onChange={e => setUserFilter(e.target.value as 'all' | 'banned' | 'moderator')}>
                <option value="all">全部已登记 QQ</option>
                <option value="banned">当前被禁言名单</option>
                <option value="moderator">社区协管名单</option>
              </select>
            </div>
            <div className="comm-admin-table">
              {(usersData?.items || []).length === 0 ? (
                <p className="hint">暂无匹配的成员记录。</p>
              ) : (
                usersData!.items.map(u => (
                  <div key={u.qq} className="comm-admin-row">
                    <div className="comm-admin-row-head">
                      <span>
                        <strong>{u.nickname}</strong> · 账号: <code>{u.qq}</code> · 有效发言 {u.post_count} 条
                      </span>
                      <div className="row">
                        {u.role === 'env_admin' && <StatusBadge ok>环境变量维护者</StatusBadge>}
                        {u.role === 'moderator' && <StatusBadge ok>社区协管</StatusBadge>}
                        {u.is_banned && (
                          <StatusBadge ok={false} warn>
                            已禁言{u.banned_until === -1 ? '（永久）' : ''}
                          </StatusBadge>
                        )}
                      </div>
                    </div>
                    {u.is_banned && u.ban_reason && (
                      <p className="hint" style={{ margin: '4px 0' }}>原因：{u.ban_reason}</p>
                    )}
                    <div className="row">
                      {u.is_banned ? (
                        <button type="button" onClick={() => handleUnban(u.qq)}>解除禁言</button>
                      ) : (
                        u.role !== 'env_admin' && (
                          <button
                            type="button"
                            className="danger"
                            onClick={async () => {
                              await api('/api/community/moderation/ban', {
                                method: 'POST',
                                body: JSON.stringify({ qq: u.qq, duration_hours: 72, reason: '后台快捷禁言 72h' }),
                              });
                              flash(`已禁言 ${u.nickname} 72 小时`);
                              loadUsers();
                            }}
                          >
                            禁言 72h
                          </button>
                        )
                      )}
                      {u.role !== 'env_admin' && (
                        <button type="button" onClick={() => handleToggleRole(u)}>
                          {u.role === 'moderator' ? '撤销协管身份' : '设为前台协管'}
                        </button>
                      )}
                    </div>
                  </div>
                ))
              )}
            </div>
          </>
        )}
      </section>
    </div>
  );
}
