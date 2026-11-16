import { useEffect, useRef, useState } from 'react';
import { api } from '../api';
import { copyText } from '../library-state';
import type { CommunityViewerState, QqChallengeBeginResponse, QqChallengePollResponse } from '../types';

type Props = {
  open: boolean;
  viewer: CommunityViewerState | null;
  onClose: () => void;
  onLoggedIn: (viewer: CommunityViewerState) => void;
  onLoggedOut: () => void;
};

type StepState =
  | { kind: 'input' }
  | {
      kind: 'waiting';
      qq: string;
      botQq: string;
      code: string;
      expiresAt: string;
      statusText: string;
      consuming: boolean;
    };

export function QqLoginModal({ open, viewer, onClose, onLoggedIn, onLoggedOut }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [qq, setQq] = useState('');
  const [step, setStep] = useState<StepState>({ kind: 'input' });
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [copiedNotice, setCopiedNotice] = useState('');
  const [remainSeconds, setRemainSeconds] = useState<number | null>(null);
  const consumingRef = useRef(false);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (open && !dialog.open) {
      setError('');
      setCopiedNotice('');
      dialog.showModal();
    } else if (!open && dialog.open) {
      dialog.close();
    }
  }, [open]);

  // 倒计时刷新
  useEffect(() => {
    if (!open || step.kind !== 'waiting') {
      setRemainSeconds(null);
      return;
    }
    const update = () => {
      const target = Date.parse(step.expiresAt);
      if (Number.isNaN(target)) {
        setRemainSeconds(null);
        return;
      }
      const diff = Math.max(0, Math.floor((target - Date.now()) / 1000));
      setRemainSeconds(diff);
    };
    update();
    const id = window.setInterval(update, 1000);
    return () => window.clearInterval(id);
  }, [open, step]);

  // 每 2.5 秒轮询已绑定的服务端验证状态，变为 verified 后自动触发一次性 consume
  useEffect(() => {
    if (!open || step.kind !== 'waiting') return;
    let cancelled = false;

    const doConsume = async () => {
      if (consumingRef.current) return;
      consumingRef.current = true;
      setStep(prev => (prev.kind === 'waiting' ? { ...prev, consuming: true, statusText: '已收到验证码，正在核验群成员与角色卡资格…' } : prev));
      try {
        const res = await api<CommunityViewerState & { ok: boolean }>('/api/auth/qq/challenges/consume', {
          method: 'POST',
        });
        if (cancelled) return;
        if (res.authenticated && res.user) {
          setStep({ kind: 'input' });
          onLoggedIn(res);
          onClose();
        } else {
          setError('未能通过登录资格核验，请重新申请验证码。');
          setStep({ kind: 'input' });
        }
      } catch (err) {
        if (cancelled) return;
        const msg = err instanceof Error ? err.message : '登录核验失败';
        setError(msg);
        setStep({ kind: 'input' });
      } finally {
        consumingRef.current = false;
      }
    };

    const tick = async () => {
      if (cancelled || consumingRef.current) return;
      try {
        const poll = await api<QqChallengePollResponse>('/api/auth/qq/challenges/current');
        if (cancelled) return;
        if (poll.status === 'verified') {
          await doConsume();
        } else if (poll.status === 'expired') {
          setError('验证码已过期（有效期 5 分钟），请重新获取新验证码。');
          setStep({ kind: 'input' });
        } else if (poll.status === 'denied') {
          setError('群成员或角色卡资格未通过，请确认在群且已上传角色卡后重新获取验证码。');
          setStep({ kind: 'input' });
        } else if (poll.status === 'consumed') {
          setError('该验证码已被使用，请重新申请新验证码。');
          setStep({ kind: 'input' });
        }
      } catch (err) {
        if (cancelled) return;
        const msg = err instanceof Error ? err.message : '';
        if (msg.includes('没有待处理的登录验证') || msg.includes('不存在或已失效')) {
          setError(msg);
          setStep({ kind: 'input' });
        }
      }
    };

    const timer = window.setInterval(tick, 2500);
    return () => {
      cancelled = true;
      window.clearInterval(timer);
    };
  }, [open, step.kind, onLoggedIn, onClose]);

  const handleBegin = async (event: React.FormEvent) => {
    event.preventDefault();
    const cleaned = qq.trim();
    if (!/^[1-9][0-9]{4,11}$/.test(cleaned)) {
      setError('请输入有效的 QQ 号（5～12 位数字，不以 0 开头）。');
      return;
    }
    setSubmitting(true);
    setError('');
    setCopiedNotice('');
    try {
      const data = await api<QqChallengeBeginResponse>('/api/auth/qq/challenges', {
        method: 'POST',
        body: JSON.stringify({ qq: cleaned }),
      });
      setStep({
        kind: 'waiting',
        qq: data.qq,
        botQq: data.bot_qq,
        code: data.verification_code,
        expiresAt: data.expires_at,
        statusText: '等待对应 QQ 向 BOT 发送私聊验证码…（自动检测中，每 2.5 秒刷新）',
        consuming: false,
      });
    } catch (err) {
      setError(err instanceof Error ? err.message : '无法获取验证码，请稍后重试');
    } finally {
      setSubmitting(false);
    }
  };

  const handleCopy = async (text: string, label: string) => {
    try {
      await copyText(text);
      setCopiedNotice(`已复制${label}`);
      window.setTimeout(() => setCopiedNotice(''), 2500);
    } catch {
      setCopiedNotice('浏览器不支持自动复制，请手动选中复制');
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

  const formatCountdown = (sec: number | null) => {
    if (sec === null) return '5 分钟内有效';
    if (sec <= 0) return '已到期';
    const m = Math.floor(sec / 60);
    const s = String(sec % 60).padStart(2, '0');
    return `剩余 ${m}:${s}`;
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
            <span className="comm-kicker">宏观界域 · 群成员身份核验</span>
            <h2 id="comm-auth-title">{viewer?.authenticated ? '当前已绑定调查员身份' : 'QQ 私聊验证码登录'}</h2>
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
              <div className="comm-profile-meta">
                <div className="comm-profile-name">
                  <strong>{viewer.user.nickname}</strong>
                  {viewer.user.badge && <span className="comm-role-badge">{viewer.user.badge}</span>}
                </div>
                <p>绑定 QQ：{viewer.user.qq} · 群内已登记角色卡 {viewer.user.card_count} 份</p>
                {viewer.user.is_banned && (
                  <p className="comm-ban-alert" role="alert">
                    发言受限：{viewer.user.ban_reason || '已被管理员禁言'}
                  </p>
                )}
              </div>
            </div>
            <div className="comm-auth-actions">
              <button type="button" onClick={onClose}>返回浏览</button>
              <button type="button" className="comm-danger-btn" disabled={submitting} onClick={handleLogout}>
                退出当前账号
              </button>
            </div>
          </div>
        ) : step.kind === 'input' ? (
          <form className="comm-auth-form" onSubmit={handleBegin}>
            <p className="comm-auth-desc">
              仅限<strong>宏观界域主群（{viewer?.group_id || 672204828}）</strong>成员且在群文件上传过角色卡（<code>.xlsx</code>）的调查员发言。
              无需密码或第三方授权，使用本人 QQ 向指定 BOT 发送一次性验证码即可登录。
            </p>
            {viewer && !viewer.auth_configured && (
              <p className="comm-auth-warn" role="status">
                提示：当前服务器尚未配置 <code>ONEBOT_SEQUENCE_TOKEN</code>，请联系维护者配置后再发起验证。
              </p>
            )}
            <label className="comm-field-label" htmlFor="comm-qq-input">
              群内本人 QQ 号
            </label>
            <div className="comm-input-row">
              <input
                id="comm-qq-input"
                type="text"
                inputMode="numeric"
                autoComplete="off"
                placeholder="输入你在主群内的 QQ 号…"
                value={qq}
                onChange={e => setQq(e.target.value.replace(/[^\d]/g, '').slice(0, 12))}
              />
              <button type="submit" className="comm-primary-btn" disabled={submitting || !qq.trim()}>
                {submitting ? '正在校验在群状态…' : '获取验证码'}
              </button>
            </div>
            <ul className="comm-auth-tips">
              <li>第 1 步先校验你的 QQ 当前在群 <code>672204828</code>，不在群不发码。</li>
              <li>第 2 步使用刚填的 QQ 向 BOT（<code>{viewer?.bot_qq || '3239195029'}</code>）私聊发送完整验证码（支持好友或群临时会话）。</li>
              <li>BOT 只收码不回复、不加好友；网页检测到后自动核验角色卡并完成登录。</li>
            </ul>
            {error && <p className="comm-auth-error" role="alert">{error}</p>}
          </form>
        ) : (
          <div className="comm-auth-waiting">
            <div className="comm-code-panel">
              <div className="comm-code-meta">
                <span>申请 QQ：<b>{step.qq}</b></span>
                <span className={remainSeconds !== null && remainSeconds <= 30 ? 'is-urgent' : ''}>
                  {formatCountdown(remainSeconds)}
                </span>
              </div>
              <div className="comm-code-box">
                <code data-testid="qq-verification-code">{step.code}</code>
                <button type="button" className="comm-primary-btn" onClick={() => handleCopy(step.code, '验证码')}>
                  复制验证码
                </button>
              </div>
              <div className="comm-bot-row">
                <span>接收 BOT QQ：<code>{step.botQq}</code>（可在群成员列表找到它发起临时会话）</span>
                <button type="button" onClick={() => handleCopy(step.botQq, ' BOT QQ')}>
                  复制 BOT 号
                </button>
              </div>
            </div>

            <div className="comm-poll-status" role="status" aria-live="polite">
              <span className="comm-pulse-dot" aria-hidden="true" />
              <span>{step.statusText}</span>
            </div>

            <p className="comm-auth-subnote">
              请务必使用 <b>{step.qq}</b> 本人账号单独发送上方 <code>TRPG-...</code> 完整字符串，<strong>不要附加其他文字、引用或图片</strong>。BOT 不会回复消息，发送后留在本页等待 2～5 秒即可自动跳转。
            </p>

            {copiedNotice && <p className="comm-copied-hint" role="status">{copiedNotice}</p>}
            {error && <p className="comm-auth-error" role="alert">{error}</p>}

            <div className="comm-auth-actions">
              <button
                type="button"
                disabled={step.consuming}
                onClick={() => {
                  setError('');
                  setStep({ kind: 'input' });
                }}
              >
                更换 QQ / 重新获取
              </button>
              <button type="button" onClick={onClose}>稍后继续</button>
            </div>
          </div>
        )}
      </div>
    </dialog>
  );
}
