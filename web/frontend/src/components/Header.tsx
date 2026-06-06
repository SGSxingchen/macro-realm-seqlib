import { ThemeToggle } from './ThemeToggle';
import { AppTab } from '../types';

type Props = {
  recordCount: number;
  tab: AppTab;
  onTab: (t: AppTab) => void;
};

export function Header({ recordCount, tab, onTab }: Props) {
  return (
    <header className="topbar">
      <div className="brand-mark" aria-label="Macro Realm">MR</div>
      <div className="brand-copy">
        <p className="eyebrow">收录 {recordCount} 条记录</p>
        <h1>宏观界域强化序列库</h1>
      </div>
      <nav>
        <ThemeToggle />
        <button type="button" className={tab === 'read' ? 'active' : ''} onClick={() => onTab('read')}>查阅</button>
        <button type="button" className={tab === 'updates' ? 'active' : ''} onClick={() => onTab('updates')}>最近更新</button>
        <button type="button" className={tab === 'stats' ? 'active' : ''} onClick={() => onTab('stats')}>结团统计</button>
        <button type="button" className={tab === 'admin' ? 'active admin-entry' : 'admin-entry'} onClick={() => onTab('admin')}>后台</button>
      </nav>
    </header>
  );
}
