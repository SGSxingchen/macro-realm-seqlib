import { ThemeToggle } from './ThemeToggle';
import { ResponsiveMenu } from './ResponsiveMenu';
import type { AppTab, CommunityViewerState } from '../types';

type Props = {
  recordCount: number;
  tab: AppTab;
  onTab: (tab: AppTab) => void;
  viewer?: CommunityViewerState | null;
  onOpenAuth?: () => void;
};

export function Header({ recordCount, tab, onTab, viewer, onOpenAuth }: Props) {
  const user = viewer?.user;
  return (
    <header className="topbar realm-topbar">
      <button type="button" className="realm-brand" onClick={() => onTab('board')} aria-label="宏观界域 · 返回轮回者论坛">
        <span className="realm-brand-mark" aria-hidden="true">∞</span>
        <span className="realm-brand-text">
          <strong>宏观界域<span>强化序列库</span></strong>
        </span>
      </button>
      {recordCount > 0 && <div className="realm-library-count">已收录 <b>{recordCount}</b> 项序列</div>}
      <ResponsiveMenu label="导航" className="realm-nav-menu" closeOnSelect>
        <nav aria-label="主导航">
          {([
            ['board', '轮回者论坛'],
            ['read', '序列检索'],
            ['updates', '最近更新'],
            ['stats', '结团统计'],
            ['admin', '后台'],
          ] as const).map(([value, label]) => (
            <button
              key={value}
              type="button"
              aria-current={tab === value ? 'page' : undefined}
              className={tab === value ? 'active' : ''}
              onClick={() => onTab(value)}
            >
              {label}
            </button>
          ))}
          <a href="/tabletop.html" className="tabletop-entry">战术桌 ↗</a>
          {onOpenAuth && (
            <button
              type="button"
              className={`comm-topbar-auth${user ? ' is-logged' : ''}`}
              onClick={onOpenAuth}
              title={user ? `${user.nickname}（${user.qq_masked}）` : '使用主群机器人 /登录 短码认证'}
            >
              {user ? (
                <>
                  <img src={user.avatar_url} alt="" className="comm-topbar-avatar" referrerPolicy="no-referrer" />
                  <span className="comm-topbar-nick">{user.nickname}</span>
                  {user.badge && <em className="comm-topbar-badge">{user.badge}</em>}
                </>
              ) : (
                '轮回者登录'
              )}
            </button>
          )}
          <ThemeToggle />
        </nav>
      </ResponsiveMenu>
    </header>
  );
}
