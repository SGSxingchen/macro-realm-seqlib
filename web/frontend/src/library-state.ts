import type { AppTab, SearchFilters } from './types';

export type LocationState = { tab: AppTab; filters: SearchFilters; openPath: string; anchor: string; threadId?: number };
export const emptyFilters = (): SearchFilters => ({ q: '', category: '', kinds: [], sides: [], authors: [] });

export function readLocation(href = window.location.href): LocationState {
  const url = new URL(href);
  const sp = url.searchParams;
  const tab = sp.get('tab');
  const openPath = sp.get('open') || '';
  const rawThread = sp.get('thread');
  const parsedThread = rawThread && /^\d+$/.test(rawThread) ? Number(rawThread) : 0;
  const filters: SearchFilters = {
    q: sp.get('q') || '',
    category: sp.get('cat') || '',
    kinds: sp.getAll('kinds').filter(Boolean),
    sides: sp.getAll('sides').filter(Boolean),
    authors: sp.getAll('authors').filter(Boolean),
  };
  const hasSearchIntent = Boolean(
    openPath ||
    filters.q ||
    filters.category ||
    filters.kinds.length ||
    filters.sides.length ||
    filters.authors.length
  );
  let resolvedTab: AppTab = 'board';
  if (tab === 'admin' || tab === 'updates' || tab === 'stats' || tab === 'read' || tab === 'board') {
    resolvedTab = tab;
  } else if (parsedThread > 0) {
    resolvedTab = 'board';
  } else if (hasSearchIntent) {
    resolvedTab = 'read';
  }
  return {
    tab: resolvedTab,
    filters,
    openPath,
    anchor: url.hash.slice(1),
    ...(parsedThread > 0 ? { threadId: parsedThread } : {}),
  };
}

export function locationUrl(state: LocationState): string {
  const sp = new URLSearchParams();
  const hasSearchIntent = Boolean(
    state.openPath ||
    state.filters.q ||
    state.filters.category ||
    state.filters.kinds.length ||
    state.filters.sides.length ||
    state.filters.authors.length
  );
  if (state.tab === 'read') {
    if (!hasSearchIntent) sp.set('tab', 'read');
  } else if (state.tab !== 'board') {
    sp.set('tab', state.tab);
  }
  if (state.tab === 'board' && state.threadId && state.threadId > 0) {
    sp.set('thread', String(state.threadId));
  }
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
