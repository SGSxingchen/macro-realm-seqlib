import type { AppTab, SearchFilters } from './types';

export type LocationState = { tab: AppTab; filters: SearchFilters; openPath: string; anchor: string };
export const emptyFilters = (): SearchFilters => ({ q: '', category: '', kinds: [], sides: [], authors: [] });

export function readLocation(href = window.location.href): LocationState {
  const url = new URL(href);
  const sp = url.searchParams;
  const tab = sp.get('tab');
  return {
    tab: tab === 'admin' || tab === 'updates' || tab === 'stats' ? tab : 'read',
    filters: { q: sp.get('q') || '', category: sp.get('cat') || '', kinds: sp.getAll('kinds').filter(Boolean), sides: sp.getAll('sides').filter(Boolean), authors: sp.getAll('authors').filter(Boolean) },
    openPath: sp.get('open') || '',
    anchor: url.hash.slice(1),
  };
}

export function locationUrl(state: LocationState): string {
  const sp = new URLSearchParams();
  if (state.tab !== 'read') sp.set('tab', state.tab);
  if (state.filters.q) sp.set('q', state.filters.q);
  if (state.filters.category) sp.set('cat', state.filters.category);
  for (const key of ['kinds', 'sides', 'authors'] as const) state.filters[key].forEach(value => sp.append(key, value));
  if (state.openPath) sp.set('open', state.openPath);
  const qs = sp.toString();
  return `${window.location.pathname}${qs ? '?' + qs : ''}${state.openPath && state.anchor ? '#' + state.anchor : ''}`;
}

export type SavedResource = { path: string; title: string };
const STORAGE_KEY = 'seqlib-saved-v1';
export function loadSaved(): SavedResource[] {
  try {
    const value: unknown = JSON.parse(localStorage.getItem(STORAGE_KEY) || '[]');
    if (!Array.isArray(value)) return [];
    return value.filter((item): item is SavedResource => Boolean(item && typeof item.path === 'string' && item.path.startsWith('序列库/') && typeof item.title === 'string')).slice(0, 100);
  } catch { return []; }
}
export function saveResources(items: SavedResource[]): boolean {
  try { localStorage.setItem(STORAGE_KEY, JSON.stringify(items)); return true; }
  catch { return false; }
}

/** Async clipboard only: no silent success and no selection-destroying fallback. */
export async function copyText(text: string): Promise<void> {
  if (!navigator.clipboard?.writeText) throw new Error('浏览器不允许复制，请切换到原文后手动选择。');
  await navigator.clipboard.writeText(text);
}
