import { useCallback, useEffect, useRef, useState } from 'react';
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

const PENDING_STORAGE_KEY = 'seqlib-pending-qq-auth-v1';

function readCachedPending(): { qq: string; botQq: string; code: string; expiresAt: string } | null {
  try {
    const raw = sessionStorage.getItem(PENDING_STORAGE_KEY) || localStorage.getItem(PENDING_STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as { qq?: string; botQq?: string; code?: string; expiresAt?: string };
    if (!parsed.qq || !parsed.expiresAt) return null;
    const exp = Date.parse(parsed.expiresAt);
    if (!Number.isNaN(exp) && exp <= Date.now()) {
      clearCachedPending();
      return null;
    }
    return {
      qq: parsed.qq,
      botQq: parsed.botQq || '3239195029',
      code: parsed.code || '',
      expiresAt: parsed.expiresAt,
    };
  } catch {
    return null;
  }
}

function writeCachedPending(info: { qq: string; botQq: string; code: string; expiresAt: string }): void {
  const payload = JSON.stringify(info);
  try {
    sessionStorage.setItem(PENDING_STORAGE_KEY, payload);
  } catch {
    /* ignore */
  }
  try {
    localStorage.setItem(PENDING_STORAGE_KEY, payload);
  } catch {
    /* ignore */
  }
}

function clearCachedPending(): void {
  try {
    sessionStorage.removeItem(PENDING_STORAGE_KEY);
  } catch {
    /* ignore */
  }
  try {
    localStorage.removeItem(PENDING_STORAGE_KEY);
  } catch {
    /* ignore */
  }
}

export function QqLoginModal({ open, viewer, onClose, onLoggedIn, onLoggedOut }: Props) {
  const dialogRef = useRef<HTMLDialogElement>(null);
  const [qq, setQq] = useState('');
  const [step, setStep] = useState<StepState>({ kind: 'input' });
  const [submitting, setSubmitting] = useState(false);
  const [error, setError] = useState('');
  const [copiedNotice, setCopiedNotice] = useState('');
  const [remainSeconds, setRemainSeconds] = useState<number | null>(null);
  const consumingRef = useRef(false);

  // 当服务端会话或本地存储存在未过期的挂起验证（例如在 QQ 内置浏览器切出去发码返回导致页面重载）时，自动恢复等待界面
  useEffect(() => {
    if (viewer?.authenticated) {
      clearCachedPending();
      setStep({ kind: 'input' });
      return;
    }
    if (step.kind === 'waiting') return;
    const serverPending = viewer?.pending_challenge;
    const cached = readCachedPending();
    const targetExp = serverPending?.expires_at || cached?.expiresAt || '';
    if (targetExp) {
      const expMs = Date.parse(targetExp);
      if (Number.isNaN(expMs) || expMs > Date.now()) {
        const restoredQq = serverPending?.qq || cached?.qq || '';
        const restoredCode = serverPending?.verification_code || cached?.code || '';
        const restoredBot = serverPending?.bot_qq || cached?.botQq || '3239195029';
        if (restoredQq) {
          setQq(restoredQq);
          setStep({
            kind: 'waiting',
            qq: restoredQq,
            botQq: restoredBot,
            code: restoredCode,
            expiresAt: targetExp,
            statusText: '已恢复未过期的验证会话，正在自动检查验证码投递状态…',
            consuming: false,
          });
        }
      }
    }
  }, [viewer?.authenticated, viewer?.pending_challenge, step.kind]);

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
    if (step.kind !== 'waiting') {
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
  }, [step]);

  const checkAndConsume = useCallback(async () => {
    if (consumingRef.current) return;
    try {
      const poll = await api<QqChallengePollResponse>('/api/auth/qq/challenges/current');
      if (poll.status === 'verified') {
        if (consumingRef.current) return;
        consumingRef.current = true;
        setStep(prev =>
          prev.kind === 'waiting'
            ? { ...prev, consuming: true, statusText: '已收到验证码，正在核验交易群成员与角色卡资格…' }
            : prev
        );
        try {
          const res = await api<CommunityViewerState & { ok: boolean }>('/api/auth/qq/challenges/consume', {
            method: 'POST',
          });
          clearCachedPending();
          if (res.authenticated && res.user) {
            setStep({ kind: 'input' });
            onLoggedIn(res);
            onClose();
          } else {
            setError('未能通过轮回者资格核验，请重新申请验证码。');
            setStep({ kind: 'input' });
          }
        } catch (err) {
          clearCachedPending();
          setError(err instanceof Error ? err.message : '登录核验失败');
          setStep({ kind: 'input' });
        } finally {
          consumingRef.current = false;
        }
      } else if (poll.status === 'expired') {
        clearCachedPending();
        setError('验证码已过期（有效期 5 分钟），请重新获取新验证码。');
        setStep({ kind: 'input' });
      } else if (poll.status === 'denied') {
        clearCachedPending();
        setError('交易群成员或角色卡资格未通过，请确认在群且已上传角色卡后重新获取验证码。');
        setStep({ kind: 'input' });
      } else if (poll.status === 'consumed') {
        clearCachedPending();
        setError('该验证码已被使用，请重新申请新验证码。');
        setStep({ kind: 'input' });
      }
    } catch (err) {
      const msg = err instanceof Error ? err.message : '';
      if (msg.includes('没有待处理的登录验证') || msg.includes('不存在或已失效')) {
        clearCachedPending();
        setError(msg);
        setStep({ kind: 'input' });
      }
    }
  }, [onLoggedIn, onClose]);

  // 即使弹窗未打开，只要处于 waiting 状态（如页面从 QQ 切回重载恢复后），也在后台立即检测并完成自动登录；可见切回时立即触发一次检查
  useEffect(() => {
    if (step.kind !== 'waiting') return;
    checkAndConsume();
    const timer = window.setInterval(checkAndConsume, 2500);
    const onVisible = () => {
      if (document.visibilityState === 'visible') checkAndConsume();
    };
    document.addEventListener('visibilitychange', onVisible);
    window.addEventListener('focus', checkAndConsume);
    return () => {
      window.clearInterval(timer);
      document.removeEventListener('visibilitychange', onVisible);
      window.removeEventListener('focus', checkAndConsume);
    };
  }, [step.kind, checkAndConsume]);

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
      writeCachedPending({
        qq: data.qq,
        botQq: data.bot_qq,
        code: data.verification_code,
        expiresAt: data.expires_at,
      });
      setStep({
        kind: 'waiting',
        qq: data.qq,
        botQq: data.bot_qq,
        code: data.verification_code,
        expiresAt: data.expires_at,
        statusText: '等待对应 QQ 向 BOT 发送私聊验证码…（切回本页会自动恢复并完成验证）',
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
      setCopiedNotice('当前环境限制自动复制，请长按上方文字手动复制');
    }
  };

  const handleLogout = async () => {
    setSubmitting(true);
    try {
      await api('/api/auth/logout', { method: 'POST' });
      clearCachedPending();
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
            <span className="comm-kicker">宏观界域 · 轮回者终端认证</span>
            <h2 id="comm-auth-title">{viewer?.authenticated ? '当前已认证轮回者身份' : 'QQ 私聊验证码登录'}</h2>
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
                <p>绑定 QQ：{viewer.user.qq} · 交易群内已登记角色卡 {viewer.user.card_count} 份</p>
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
              仅限<strong>交易群（{viewer?.group_id || 672204828}）</strong>成员且在群文件上传过角色卡（<code>.xlsx</code>）的轮回者发言。
              无需密码，使用本人 QQ 向指定 BOT 发送一次性验证码即可登录。
            </p>
            {viewer && !viewer.auth_configured && (
              <p className="comm-auth-warn" role="status">
                提示：当前服务器尚未配置 <code>ONEBOT_SEQUENCE_TOKEN</code>，请联系维护者配置后再发起验证。
              </p>
            )}
            <label className="comm-field-label" htmlFor="comm-qq-input">
              交易群内本人 QQ 号
            </label>
            <div className="comm-input-row">
              <input
                id="comm-qq-input"
                type="text"
                inputMode="numeric"
                autoComplete="off"
                placeholder="输入你在交易群内的 QQ 号…"
                value={qq}
                onChange={e => setQq(e.target.value.replace(/[^\d]/g, '').slice(0, 12))}
              />
              <button type="submit" className="comm-primary-btn" disabled={submitting || !qq.trim()}>
                {submitting ? '正在校验在群状态…' : '获取验证码'}
              </button>
            </div>
            <ul className="comm-auth-tips">
              <li>第 1 步先校验你的 QQ 当前在交易群 <code>672204828</code>，不在群不发码。</li>
              <li>第 2 步使用刚填的 QQ 向 BOT（<code>{viewer?.bot_qq || '3239195029'}</code>）私聊发送完整验证码（好友或群临时会话均可）。</li>
              <li>即使从 QQ 内置网页切出去发消息后再重新点开链接，未过期的验证也会自动续接并完成登录。</li>
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
              {step.code ? (
                <div className="comm-code-box">
                  <code data-testid="qq-verification-code">{step.code}</code>
                  <button type="button" className="comm-primary-btn" onClick={() => handleCopy(step.code, '验证码')}>
                    复制验证码
                  </button>
                </div>
              ) : (
                <p className="comm-auth-subnote">正在等待刚才生成的验证码私聊投递…</p>
              )}
              <div className="comm-bot-row">
                <span>接收 BOT QQ：<code>{step.botQq}</code>（可在交易群成员列表找到它发起私聊）</span>
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
              请使用 <b>{step.qq}</b> 本人账号单独发送上方 <code>TRPG-...</code> 完整字符串，<strong>不要附加其他文字、引用或图片</strong>。BOT 不回复消息；若你在 QQ 内置浏览器操作，发完重新点进链接即可自动完成登录。
            </p>

            {copiedNotice && <p className="comm-copied-hint" role="status">{copiedNotice}</p>}
            {error && <p className="comm-auth-error" role="alert">{error}</p>}

            <div className="comm-auth-actions">
              <button type="button" disabled={step.consuming} onClick={checkAndConsume}>
                我已发送，立即检查
              </button>
              <button
                type="button"
                disabled={step.consuming}
                onClick={() => {
                  clearCachedPending();
                  setError('');
                  setStep({ kind: 'input' });
                }}
              >
                更换 QQ / 重新获取
              </button>
            </div>
          </div>
        )}
      </div>
    </dialog>
  );
}
