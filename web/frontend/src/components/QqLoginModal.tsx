import { useEffect, useRef, useState } from 'react';
import { api } from '../api';
import { copyText } from '../library-state';
import type { BotInspectCodeResponse, CommunityViewerState } from '../types';

type Props = {
  open: boolean;
  viewer: CommunityViewerState | null;
  onClose: () => void;
  onLoggedIn: (viewer: CommunityViewerState) => void;
  onLoggedOut: () => void;
};

type StepState =
  | { kind: 'code' }
  | {
      kind: 'confirm';
      code: string;
      userMasked: string;
      isNewUser: boolean;
      sealdiceNickname: string;
      avatarUrl: string;
    };

export function QqLoginModal({ open, viewer, onClose, onLoggedIn, onLoggedOut }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [code, setCode] = useState('');
  const [nickname, setNickname] = useState('');
  const [step, setStep] = useState<StepState>({ kind: 'code' });
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
    if (viewer?.authenticated) {
      setStep({ kind: 'code' });
      setCode('');
    }
  }, [viewer?.authenticated]);

  const handleCopyCommand = async () => {
    try {
      await copyText('/登录');
      setNotice('已复制指令 /登录，前往主群 @' + botName + ' 发送即可');
      window.setTimeout(() => setNotice(''), 3000);
    } catch {
      setNotice('请手动在主群发送 /登录');
    }
  };

  const handleInspectCode = async (event: React.FormEvent) => {
    event.preventDefault();
    const cleaned = code.replace(/\D/g, '').slice(0, 6);
    if (cleaned.length !== 6) {
      setError('请输入主群机器人回复的 6 位数字登录码。');
      return;
    }
    setSubmitting(true);
    setError('');
    setNotice('');
    try {
      const info = await api<BotInspectCodeResponse>('/api/auth/bot/inspect', {
        method: 'POST',
        body: JSON.stringify({ code: cleaned }),
      });

      // 已注册老用户直接完成登录，无需每次重复确认昵称（登录后可随时在个人面板修改昵称）
      if (!info.is_new_user) {
        const res = await api<CommunityViewerState & { ok: boolean }>('/api/auth/bot/consume', {
          method: 'POST',
          body: JSON.stringify({ code: cleaned, nickname: '' }),
        });
        if (res.authenticated && res.user) {
          setCode('');
          setStep({ kind: 'code' });
          onLoggedIn(res);
          onClose();
          return;
        }
      }

      setNickname(info.suggested_nickname || '');
      setStep({
        kind: 'confirm',
        code: cleaned,
        userMasked: info.user_key_masked,
        isNewUser: info.is_new_user,
        sealdiceNickname: info.sealdice_nickname,
        avatarUrl: info.avatar_url,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : '登录码核验失败，请重新获取');
    } finally {
      setSubmitting(false);
    }
  };

  const handleCompleteRegistration = async (event: React.FormEvent) => {
    event.preventDefault();
    if (step.kind !== 'confirm') return;
    const trimmedNick = nickname.trim();
    if (!trimmedNick) {
      setError('请输入要在论坛展示的昵称。');
      return;
    }
    setSubmitting(true);
    setError('');
    try {
      const res = await api<CommunityViewerState & { ok: boolean }>('/api/auth/bot/consume', {
        method: 'POST',
        body: JSON.stringify({ code: step.code, nickname: trimmedNick }),
      });
      if (res.authenticated && res.user) {
        setCode('');
        setStep({ kind: 'code' });
        onLoggedIn(res);
        onClose();
      } else {
        setError('未能建立登录会话，请重新获取登录码。');
        setStep({ kind: 'code' });
      }
    } catch (err) {
      setError(err instanceof Error ? err.message : '完成登录失败');
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
        onClose();
      }}
      onClick={e => {
        if (e.target === e.currentTarget) onClose();
      }}
    >
      <div className="comm-auth-card">
        <header className="comm-auth-head">
          <div>
            <span className="comm-kicker">宏观界域 · 轮回者终端认证</span>
            <h2 id="comm-auth-title">
              {viewer?.authenticated
                ? '当前已认证轮回者身份'
                : step.kind === 'confirm'
                  ? '确认轮回者展示资料'
                  : '主群机器人登录码认证'}
            </h2>
          </div>
          <button type="button" className="comm-dialog-close" onClick={onClose} aria-label="关闭登录窗口">
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
        ) : step.kind === 'code' ? (
          <form className="comm-auth-form" onSubmit={handleInspectCode}>
            <p className="comm-auth-desc">
              仅限<strong>宏观界域主群</strong>成员发言。无需密码，在主群内唤醒官方机器人 <strong>@{botName}</strong> 发送{' '}
              <code>/登录</code>（或点击输入框旁指令菜单），将收到的 6 位数字登录码填入下方即可。
            </p>
            {viewer && !viewer.auth_configured && (
              <p className="comm-auth-warn" role="status">
                提示：当前服务器尚未配置机器人认证密钥，请联系维护者检查配置。
              </p>
            )}

            <div className="comm-bot-row" style={{ marginTop: '4px' }}>
              <span>主群指令：<code>@{botName} /登录</code></span>
              <button type="button" onClick={handleCopyCommand}>
                复制 /登录
              </button>
            </div>

            <label className="comm-field-label" htmlFor="comm-code-input">
              6 位一次性登录码（5 分钟内有效）
            </label>
            <div className="comm-input-row">
              <input
                id="comm-code-input"
                data-testid="bot-login-code-input"
                type="text"
                inputMode="numeric"
                autoComplete="one-time-code"
                maxLength={6}
                placeholder="输入机器人回复的 6 位数字码…"
                value={code}
                onChange={e => setCode(e.target.value.replace(/\D/g, '').slice(0, 6))}
              />
              <button type="submit" className="comm-primary-btn" disabled={submitting || code.trim().length !== 6}>
                {submitting ? '正在核验…' : '下一步'}
              </button>
            </div>

            <ul className="comm-auth-tips">
              <li>在主群点击机器人指令菜单选择 <code>/登录</code>（或发送 <code>@{botName} .登录</code>）获取 6 位短码。</li>
              <li>首次登录会自动拉取你的 QQ 头像与海豹角色昵称，你可以在下一步自由修改展示昵称。</li>
              <li>验证通过后当前浏览器保持 30 天免重复登录。</li>
            </ul>
            {notice && <p className="comm-copied-hint" role="status">{notice}</p>}
            {error && <p className="comm-auth-error" role="alert">{error}</p>}
          </form>
        ) : (
          <form className="comm-auth-form" onSubmit={handleCompleteRegistration}>
            <div className="comm-profile-row" style={{ marginBottom: '8px' }}>
              {step.avatarUrl && (
                <img
                  src={step.avatarUrl}
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
                  标识：{step.userMasked}
                  {step.sealdiceNickname ? ` · 海豹原名：${step.sealdiceNickname}` : ''}
                </p>
              </div>
            </div>

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

            {error && <p className="comm-auth-error" role="alert">{error}</p>}

            <div className="comm-auth-actions">
              <button
                type="button"
                disabled={submitting}
                onClick={() => {
                  setStep({ kind: 'code' });
                  setError('');
                }}
              >
                重新输码
              </button>
              <button type="submit" className="comm-primary-btn" disabled={submitting || !nickname.trim()}>
                {submitting ? '正在激活…' : '确认并登录'}
              </button>
            </div>
          </form>
        )}
      </div>
    </dialog>
  );
}
