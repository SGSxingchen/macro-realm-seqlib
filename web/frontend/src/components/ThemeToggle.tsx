import { useEffect, useState } from 'react';

type Theme = 'dark' | 'light';

const KEY = 'seqlib-theme';

function readTheme(): Theme {
  try {
    const saved = localStorage.getItem(KEY) as Theme | null;
    if (saved === 'dark' || saved === 'light') return saved;
  } catch { /* ignore */ }
  if (typeof window !== 'undefined' && window.matchMedia?.('(prefers-color-scheme: light)').matches) return 'light';
  return 'dark';
}

function applyTheme(t: Theme) {
  document.documentElement.setAttribute('data-theme', t);
}

export function ThemeToggle() {
  const [theme, setTheme] = useState<Theme>(() => readTheme());

  useEffect(() => { applyTheme(theme); }, [theme]);
  useEffect(() => {
    try { localStorage.setItem(KEY, theme); } catch { /* ignore */ }
  }, [theme]);

  const next = theme === 'dark' ? 'light' : 'dark';
  return (
    <button
      type="button"
      className="theme-toggle"
      onClick={() => setTheme(next)}
      title={theme === 'dark' ? '切到亮色' : '切到暗色'}
      aria-label="切换主题"
    >
      {theme === 'dark' ? '☀' : '☾'}
    </button>
  );
}
