import type { SavedResource } from './library-state';

export type PanePreferences = { index: boolean; results: boolean };
const PANE_KEY = 'seqlib-panes-v1';
const TAB_KEY = 'seqlib-open-tabs-v1';

export function readPanePreferences(): PanePreferences {
  try {
    const saved = JSON.parse(localStorage.getItem(PANE_KEY) || '{}');
    return { index: saved?.index !== false, results: saved?.results !== false };
  } catch { return { index: true, results: true }; }
}
export function storePanePreferences(value: PanePreferences): void {
  try { localStorage.setItem(PANE_KEY, JSON.stringify(value)); } catch { /* Session controls still work. */ }
}
export function readOpenTabs(): SavedResource[] {
  try {
    const value: unknown = JSON.parse(sessionStorage.getItem(TAB_KEY) || '[]');
    if (!Array.isArray(value)) return [];
    const valid = value.filter((item): item is SavedResource => Boolean(item && typeof item.path === 'string' && item.path.startsWith('序列库/') && typeof item.title === 'string'));
    return valid.filter((item, index) => valid.findIndex(other => other.path === item.path) === index).slice(-30);
  } catch { return []; }
}
export function storeOpenTabs(tabs: SavedResource[]): void {
  // Persist only navigation metadata, not resource text or API credentials.
  try { sessionStorage.setItem(TAB_KEY, JSON.stringify(tabs.slice(-30))); } catch { /* In-memory tabs remain usable. */ }
}
export function rememberTab(tabs: SavedResource[], path: string, title?: string): SavedResource[] {
  if (!path) return tabs;
  const existing = tabs.find(item => item.path === path);
  if (existing) return title && existing.title !== title ? tabs.map(item => item.path === path ? { ...item, title } : item) : tabs;
  return [...tabs, { path, title: title || path.split('/').pop()?.replace(/\.txt$/i, '') || path }];
}
export function closeTab(tabs: SavedResource[], path: string, active: string): { tabs: SavedResource[]; active: string } {
  const index = tabs.findIndex(item => item.path === path);
  const remaining = tabs.filter(item => item.path !== path);
  return { tabs: remaining, active: active !== path ? active : remaining[Math.min(Math.max(index - 1, 0), remaining.length - 1)]?.path || '' };
}
