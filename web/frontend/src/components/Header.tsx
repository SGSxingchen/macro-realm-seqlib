import { ThemeToggle } from './ThemeToggle';
import type { AppTab } from '../types';

type Props = { recordCount: number; tab: AppTab; onTab: (tab: AppTab) => void };
export function Header({ recordCount, tab, onTab }: Props) {
  return <header className="topbar realm-topbar">
    <button className="realm-brand" onClick={() => onTab('read')} aria-label="宏观界域 · 返回查阅">
      <span className="realm-brand-mark" aria-hidden="true">∞</span>
      <span><strong>宏观界域<span>强化序列库</span></strong><small>MACRO REALM / ARCHIVE</small></span>
    </button>
    <div className="realm-library-count">{recordCount ? `${recordCount} 份公开档案` : '公开档案终端'}</div>
    <nav aria-label="主导航">
      <a href="/tabletop.html" className="tabletop-entry">战术桌 ↗</a>
      {([['read', '查阅档案'], ['updates', '最近更新'], ['stats', '结团统计'], ['admin', '后台']] as const).map(([value, label]) => <button key={value} type="button" aria-current={tab === value ? 'page' : undefined} className={tab === value ? 'active' : ''} onClick={() => onTab(value)}>{label}</button>)}
      <ThemeToggle />
    </nav>
  </header>;
}
