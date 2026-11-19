import { useEffect, useRef, useState } from 'react';
import { api } from '../api';
import { copyText } from '../library-state';
import type { CommunityViewerState } from '../types';

type Props = {
  open: boolean;
  viewer: CommunityViewerState | null;
  onClose: () => void;
  onLoggedIn: (viewer: CommunityViewerState) => void;
  onLoggedOut: () => void;
};

type Challenge =
  | { status: 'expired' }
  | { status: 'pending'; command: string; bot_name: string; expires_at: string }
  | {
      status: 'verified';
      is_new_user: boolean;
      user_key_masked: string;
      suggested_nickname: string;
      sealdice_nickname: string;
      avatar_url: string;
      confirmation: string;
    };

const challengeUrl = '/api/auth/web-challenges';

export function QqLoginModal({ open, viewer, onClose, onLoggedIn, onLoggedOut }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [nickname, setNickname] = useState('');
  const [challenge, setChallenge] = useState<Challenge | null>(null);
  const [attempt, setAttempt] = useState(0);
  const starting = useRef<Promise<Challenge> | null>(null);
  const renew = useRef(false);
  const [editingNick, setEditingNick] = useState(false);
  const [editNickValue, setEditNickValue] = useState('');
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [notice, setNotice] = useState('');

  const botName = viewer?.bot_name || '阿瓦娜';

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (open && !dialog.open) {
      setError('');
      setNotice('');
      setEditingNick(false);
      if (viewer?.user) {
        setEditNickValue(viewer.user.nickname);
      }
      dialog.showModal();
    } else if (!open && dialog.open) {
      dialog.close();
    }
  }, [open, viewer?.user]);

  useEffect(() => {
    if (!open || viewer?.authenticated) return;
    let cancelled = false;
    let polling = false;
    let waiting = false;
    let timer: number | undefined;
    setChallenge(null);
    setNickname('');
    setError('');
    setNotice('');

    const show = (next: Challenge) => {
      if (cancelled) return;
      setChallenge(next);
      setError('');
      waiting = next.status === 'pending';
      if (next.status === 'verified') setNickname(next.suggested_nickname);
    };
    const poll = async () => {
      if (!waiting || polling || document.hidden) return;
      polling = true;
      try {
        show(await api<Challenge>(`${challengeUrl}/current`, { signal: AbortSignal.timeout(5000) }));
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : '暂时无法检查核验状态');
      } finally {
        polling = false;
      }
    };
    const start = async () => {
      try {
        if (!starting.current) {
          const url = renew.current ? `${challengeUrl}?renew=true` : challengeUrl;
          renew.current = false;
          starting.current = api<Challenge>(url, { method: 'POST', signal: AbortSignal.timeout(8000) })
            .finally(() => { starting.current = null; });
        }
        show(await starting.current);
        if (!cancelled) timer = window.setInterval(poll, 2000);
      } catch (err) {
        if (!cancelled) {
          setChallenge({ status: 'expired' });
          setError(err instanceof Error ? err.message : '暂时无法获取登录码');
        }
      }
    };
    void start();
    document.addEventListener('visibilitychange', poll);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
      document.removeEventListener('visibilitychange', poll);
    };
  }, [open, viewer?.authenticated, attempt]);

  const requestCode = (replace = false) => {
    renew.current = replace;
    setAttempt(value => value + 1);
  };

  const handleCopyCommand = async () => {
    if (challenge?.status !== 'pending') return;
    try {
      await copyText(challenge.command);
      setNotice('已复制指令，请到主群 @机器人 后粘贴发送。');
    } catch {
      setNotice('请手动复制上方登录指令。');
    }
  };

  const handleCompleteRegistration = async (event: React.FormEvent) => {
    event.preventDefault();
    if (challenge?.status !== 'verified' || submitting) return;
    setSubmitting(true);
    setError('');
    try {
      const res = await api<CommunityViewerState & { ok: boolean }>(`${challengeUrl}/consume`, {
        method: 'POST',
        headers: { 'X-Login-Confirmation': challenge.confirmation },
        body: JSON.stringify({ nickname: challenge.is_new_user ? nickname.trim() : '' }),
        signal: AbortSignal.timeout(8000),
      });
      if (!res.authenticated || !res.user) throw new Error('未能建立登录会话');
      onLoggedIn(res);
      onClose();
    } catch (err) {
      // 请求可能已成功提交，先查询当前会话再决定是否需要重新取码。
      const viewerNow = await api<CommunityViewerState>('/api/auth/me', { signal: AbortSignal.timeout(5000) })
        .catch(() => null);
      if (viewerNow?.authenticated && viewerNow.user) {
        onLoggedIn(viewerNow);
        onClose();
      } else {
        setError(err instanceof Error ? err.message : '完成登录失败');
      }
    } finally {
      setSubmitting(false);
    }
  };

  const handleSaveNickname = async (event: React.FormEvent) => {
    event.preventDefault();
    const trimmed = editNickValue.trim();
    if (!trimmed) {
      setError('昵称不能为空');
      return;
    }
    setSubmitting(true);
    setError('');
    setNotice('');
    try {
      const res = await api<CommunityViewerState & { ok: boolean }>('/api/auth/profile', {
        method: 'PATCH',
        body: JSON.stringify({ nickname: trimmed }),
      });
      if (res.authenticated && res.user) {
        onLoggedIn(res);
        setEditingNick(false);
        setNotice('昵称已更新');
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '修改昵称失败');
    } finally {
      setSubmitting(false);
    }
  };

  const handleLogout = async () => {
    setSubmitting(true);
    try {
      await api('/api/auth/logout', { method: 'POST' });
      onLoggedOut();
      onClose();
    } catch (err) {
      setError(err instanceof Error ? err.message : '退出失败');
    } finally {
      setSubmitting(false);
    }
  };

  return (
    <dialog
      ref={dialogRef}
      className="comm-auth-dialog"
      aria-labelledby="comm-auth-title"
      onCancel={e => {
        e.preventDefault();
        if (!submitting) onClose();
      }}
      onClick={e => {
        if (e.target === e.currentTarget && !submitting) onClose();
      }}
    >
      <div className="comm-auth-card">
        <header className="comm-auth-head">
          <div>
            <span className="comm-kicker">宏观界域 · 轮回者终端认证</span>
            <h2 id="comm-auth-title">
              {viewer?.authenticated
                ? '当前已认证轮回者身份'
                : challenge?.status === 'verified'
                  ? '确认轮回者身份'
                  : '主群机器人登录码认证'}
            </h2>
          </div>
          <button type="button" className="comm-dialog-close" disabled={submitting} onClick={onClose} aria-label="关闭登录窗口">
            ×
          </button>
        </header>

        {viewer?.authenticated && viewer.user ? (
          <div className="comm-auth-profile">
            <div className="comm-profile-row">
              <img
                src={viewer.user.avatar_url}
                alt=""
                className="comm-avatar-lg"
                referrerPolicy="no-referrer"
              />
              <div className="comm-profile-meta" style={{ flex: 1, minWidth: 0 }}>
                <div className="comm-profile-name">
                  <strong>{viewer.user.nickname}</strong>
                  {viewer.user.badge && <span className="comm-role-badge">{viewer.user.badge}</span>}
                </div>
                <p>身份标识：{viewer.user.qq_masked} · 已加入宏观界域主群</p>
                {viewer.user.is_banned && (
                  <p className="comm-ban-alert" role="alert">
                    发言受限：{viewer.user.ban_reason || '已被管理员禁言'}
                  </p>
                )}
              </div>
            </div>

            {editingNick ? (
              <form onSubmit={handleSaveNickname} style={{ marginTop: '12px', display: 'grid', gap: '8px' }}>
                <label className="comm-field-label" htmlFor="comm-edit-nick-input">
                  修改论坛展示昵称（最长 32 字）
                </label>
                <div className="comm-input-row">
                  <input
                    id="comm-edit-nick-input"
                    type="text"
                    maxLength={32}
                    value={editNickValue}
                    onChange={e => setEditNickValue(e.target.value)}
                    placeholder="输入新的论坛展示昵称…"
                  />
                  <button type="submit" className="comm-primary-btn" disabled={submitting || !editNickValue.trim()}>
                    {submitting ? '保存中…' : '保存昵称'}
                  </button>
                  <button
                    type="button"
                    disabled={submitting}
                    onClick={() => {
                      setEditingNick(false);
                      setEditNickValue(viewer.user?.nickname || '');
                      setError('');
                    }}
                  >
                    取消
                  </button>
                </div>
              </form>
            ) : null}

            {notice && <p className="comm-copied-hint" role="status">{notice}</p>}
            {error && <p className="comm-auth-error" role="alert">{error}</p>}

            <div className="comm-auth-actions">
              {!editingNick && !viewer.user.is_banned && (
                <button
                  type="button"
                  onClick={() => {
                    setEditNickValue(viewer.user?.nickname || '');
                    setEditingNick(true);
                    setError('');
                    setNotice('');
                  }}
                >
                  修改昵称
                </button>
              )}
              <button type="button" onClick={onClose}>返回浏览</button>
              <button type="button" className="comm-danger-btn" disabled={submitting} onClick={handleLogout}>
                退出当前账号
              </button>
            </div>
          </div>
        ) : challenge?.status !== 'verified' ? (
          <div className="comm-auth-form">
            <p className="comm-auth-desc">
              在<strong>宏观界域主群</strong> @ 官方机器人 <strong>{challenge?.status === 'pending' ? challenge.bot_name : botName}</strong>，
              发送下方指令。收到回复后回到此页面。
            </p>
            {challenge?.status === 'pending' ? (
              <>
                <div className="comm-bot-row">
                  <code data-testid="web-login-command">{challenge.command}</code>
                  <button type="button" onClick={handleCopyCommand}>复制指令</button>
                </div>
                <p className="comm-auth-desc">登录码 5 分钟内有效。等待 QQ 核验…</p>
              </>
            ) : challenge?.status === 'expired' ? (
              <div className="comm-auth-actions">
                <button type="button" className="comm-primary-btn" onClick={() => requestCode()}>重新获取登录码</button>
              </div>
            ) : <p className="comm-auth-desc" role="status">正在获取登录码…</p>}
            {notice && <p className="comm-copied-hint" role="status">{notice}</p>}
            {error && <p className="comm-auth-error" role="alert">{error}</p>}
          </div>
        ) : (
          <form className="comm-auth-form" onSubmit={handleCompleteRegistration}>
            <div className="comm-profile-row" style={{ marginBottom: '8px' }}>
              {challenge.avatar_url && (
                <img
                  src={challenge.avatar_url}
                  alt=""
                  className="comm-avatar-lg"
                  referrerPolicy="no-referrer"
                />
              )}
              <div className="comm-profile-meta">
                <div className="comm-profile-name">
                  <strong>已核验主群成员身份</strong>
                </div>
                <p>
                  标识：{challenge.user_key_masked}
                  {challenge.sealdice_nickname ? ` · 海豹原名：${challenge.sealdice_nickname}` : ''}
                </p>
              </div>
            </div>

            <p className="comm-auth-desc">请确认这是刚才发指令的 QQ 账号。若身份不符，请换一个登录码。</p>
            {challenge.is_new_user ? (
              <>
                <label className="comm-field-label" htmlFor="comm-register-nick-input">
                  论坛展示昵称（已默认填入海豹昵称，可自由修改，最长 32 字）
                </label>
                <div className="comm-input-row">
                  <input
                    id="comm-register-nick-input"
                    data-testid="bot-login-nickname-input"
                    type="text"
                    maxLength={32}
                    placeholder="输入你在论坛展示的称呼…"
                    value={nickname}
                    onChange={e => setNickname(e.target.value)}
                    style={{ fontFamily: 'inherit' }}
                  />
                </div>
              </>
            ) : <p className="comm-auth-desc">论坛昵称：{challenge.suggested_nickname}</p>}

            {error && <p className="comm-auth-error" role="alert">{error}</p>}

            <div className="comm-auth-actions">
              <button
                type="button"
                disabled={submitting}
                onClick={() => requestCode(true)}
              >
                换一个登录码
              </button>
              <button type="submit" className="comm-primary-btn" disabled={submitting || (challenge.is_new_user && !nickname.trim())}>
                {submitting ? '正在登录…' : '确认并登录'}
              </button>
            </div>
          </form>
        )}
      </div>
    </dialog>
  );
}
