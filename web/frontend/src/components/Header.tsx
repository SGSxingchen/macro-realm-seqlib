import { ThemeToggle } from './ThemeToggle';
import { ResponsiveMenu } from './ResponsiveMenu';
import type { AppTab } from '../types';

type Props = { recordCount: number; tab: AppTab; onTab: (tab: AppTab) => void };

export function Header({ recordCount, tab, onTab }: Props) {
  return (
    <header className="topbar realm-topbar">
      <button type="button" className="realm-brand" onClick={() => onTab('read')} aria-label="宏观界域 · 返回查阅">
        <span className="realm-brand-mark" aria-hidden="true">∞</span>
        <span className="realm-brand-text">
          <strong>宏观界域<span>强化序列库</span></strong>
        </span>
      </button>
      {recordCount > 0 && <div className="realm-library-count">已收录 <b>{recordCount}</b> 项序列</div>}
      <ResponsiveMenu label="导航" className="realm-nav-menu" closeOnSelect>
        <nav aria-label="主导航">
          {([
            ['read', '查阅档案'],
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
          <ThemeToggle />
        </nav>
      </ResponsiveMenu>
    </header>
  );
}
