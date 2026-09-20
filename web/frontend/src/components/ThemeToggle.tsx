import { useEffect, useState } from 'react';

type Theme = 'dark' | 'light';
const KEY = 'seqlib-theme';
const CHANGE_EVENT = 'seqlib-theme-change';

function savedTheme(): Theme {
  try {
    const saved = localStorage.getItem(KEY);
    if (saved === 'dark' || saved === 'light') return saved;
  } catch { /* The visual toggle still works without storage. */ }
  return window.matchMedia?.('(prefers-color-scheme: light)').matches ? 'light' : 'dark';
}
function currentTheme(): Theme {
  const current = document.documentElement.dataset.theme;
  return current === 'light' || current === 'dark' ? current : savedTheme();
}
function applyTheme(theme: Theme) {
  document.documentElement.dataset.theme = theme;
  window.dispatchEvent(new Event(CHANGE_EVENT));
}

/** The header and compact-layout controls must reflect the SAME theme. A
 * newly mounted compact control must not overwrite the current appearance. */
export function ThemeToggle() {
  const [theme, setTheme] = useState(currentTheme);
  useEffect(() => {
    const sync = () => setTheme(currentTheme());
    const storage = (event: StorageEvent) => { if (event.key === KEY || event.key === null) applyTheme(savedTheme()); };
    window.addEventListener(CHANGE_EVENT, sync);
    window.addEventListener('storage', storage);
    if (!document.documentElement.dataset.theme) applyTheme(currentTheme());
    sync();
    return () => { window.removeEventListener(CHANGE_EVENT, sync); window.removeEventListener('storage', storage); };
  }, []);
  return <button type="button" className="theme-toggle" onClick={() => {
    const next = currentTheme() === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    try { localStorage.setItem(KEY, next); } catch { /* Keep the in-page choice. */ }
  }} title={theme === 'dark' ? '切到亮色' : '切到暗色'} aria-label="切换主题">{theme === 'dark' ? '☀' : '☾'}</button>;
}
