import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { api, buildQuery, routePath } from './api';
import { AppTab, Detail, Resource, ResourceListResponse, SearchFilters, TreeNode } from './types';
import { Header } from './components/Header';
import { SearchBar } from './components/SearchBar';
import { FilterRail } from './components/FilterRail';
import { ResourceList } from './components/ResourceList';
import { Reader } from './components/Reader';
import { AdminPanel } from './components/Admin';
import { RecentUpdates } from './components/RecentUpdates';
import { NormalizationReviewPage } from './components/NormalizationReview';
import { SessionStats } from './components/SessionStats';

const DEFAULT_FILTERS: SearchFilters = { q: '', category: '', kinds: [], sides: [], authors: [] };

function readUrlState(): { tab: AppTab; filters: SearchFilters; openPath: string } {
  const url = new URL(window.location.href);
  const sp = url.searchParams;
  const arr = (k: string) => sp.getAll(k).filter(Boolean);
  const rawTab = sp.get('tab');
  const tab: AppTab = rawTab === 'admin' || rawTab === 'updates' || rawTab === 'stats' ? rawTab : 'read';
  return {
    tab,
    filters: {
      q: sp.get('q') || '',
      category: sp.get('cat') || '',
      kinds: arr('kinds'),
      sides: arr('sides'),
      authors: arr('authors'),
    },
    openPath: sp.get('open') || '',
  };
}

function writeUrlState(tab: AppTab, f: SearchFilters, openPath: string) {
  const sp = new URLSearchParams();
  if (tab !== 'read') sp.set('tab', tab);
  if (f.q) sp.set('q', f.q);
  if (f.category) sp.set('cat', f.category);
  f.kinds.forEach(k => sp.append('kinds', k));
  f.sides.forEach(s => sp.append('sides', s));
  f.authors.forEach(a => sp.append('authors', a));
  if (openPath) sp.set('open', openPath);
  const qs = sp.toString();
  const url = qs ? `${window.location.pathname}?${qs}` : window.location.pathname;
  window.history.replaceState(null, '', url);
}

export function App() {
  if (window.location.pathname === '/normalize-review') {
    return <NormalizationReviewPage />;
  }

  return <LibraryApp />;
}

function LibraryApp() {
  const initial = useRef(readUrlState());
  const [tab, setTab] = useState<AppTab>(initial.current.tab);
  const [filters, setFilters] = useState<SearchFilters>(initial.current.filters);
  const [items, setItems] = useState<Resource[]>([]);
  const [tree, setTree] = useState<TreeNode[]>([]);
  const [response, setResponse] = useState<ResourceListResponse | null>(null);
  const [loading, setLoading] = useState(true);
  const [detail, setDetail] = useState<Detail | null>(null);
  const [drawerOpen, setDrawerOpen] = useState(false);

  const debounceRef = useRef<number | undefined>(undefined);
  const reqIdRef = useRef(0);

  const load = useCallback((f: SearchFilters) => {
    const id = ++reqIdRef.current;
    setLoading(true);
    const qs = buildQuery({
      q: f.q,
      category: f.category,
      kinds: f.kinds,
      sides: f.sides,
      authors: f.authors,
      include_content: false,
      limit: 500,
    });
    api<ResourceListResponse>('/api/resources' + qs)
      .then(r => {
        if (id !== reqIdRef.current) return;
        setItems(r.items);
        setResponse(r);
        setLoading(false);
      })
      .catch(err => {
        if (id !== reqIdRef.current) return;
        console.error(err);
        setLoading(false);
      });
  }, []);

  useEffect(() => {
    api<{ items: TreeNode[] }>('/api/tree').then(r => setTree(r.items)).catch(console.error);
  }, []);

  // debounce 搜索
  useEffect(() => {
    if (debounceRef.current) clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => load(filters), filters.q ? 220 : 0);
    return () => { if (debounceRef.current) clearTimeout(debounceRef.current); };
  }, [filters, load]);

  // 同步 URL
  useEffect(() => { writeUrlState(tab, filters, detail?.path || ''); }, [tab, filters, detail?.path]);

  // 处理 URL 里的 open=path 初始打开
  useEffect(() => {
    if (initial.current.openPath) openResource(initial.current.openPath);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const openResource = async (path: string) => {
    try {
      const data = await api<Detail>('/api/resources/' + routePath(path));
      setDetail(data);
      setDrawerOpen(false);
    } catch (e: unknown) {
      alert(e instanceof Error ? e.message : String(e));
    }
  };

  const onPickCat = (path: string) => setFilters(f => ({ ...f, category: path.startsWith('序列库') ? path.split('/').slice(1).join('/') : '' }));
  const onToggleFacet = (group: 'kinds' | 'sides' | 'authors', name: string) => {
    setFilters(f => {
      const cur = f[group];
      return { ...f, [group]: cur.includes(name) ? cur.filter(x => x !== name) : [...cur, name] };
    });
  };
  const onClearFacets = () => setFilters(f => ({ ...f, kinds: [], sides: [], authors: [] }));
  const updateQ = (q: string) => setFilters(f => ({ ...f, q }));

  const highlightTokens = useMemo(
    () => (response?.tokens || []).filter(Boolean),
    [response?.tokens],
  );

  const selectedCatFull = filters.category ? `序列库/${filters.category}` : '';

  return (
    <main className={tab === 'admin' ? 'app admin-mode' : 'app'}>
      <Header recordCount={response?.total ?? 0} tab={tab} onTab={setTab} />

      {tab === 'read' ? (
        <div className="reader-shell">
          <div className="reader-bar">
            <button className="drawer-btn" onClick={() => setDrawerOpen(true)} aria-label="打开筛选">⊞</button>
            <SearchBar
              value={filters.q}
              onChange={updateQ}
              count={response?.total ?? 0}
              searching={loading}
              pinyinReady={!!response?.engine?.pinyin}
              onClear={() => updateQ('')}
            />
          </div>
          <div className="reader-grid">
            <div className={`rail-shell ${drawerOpen ? 'open' : ''}`}>
              <button className="drawer-close" onClick={() => setDrawerOpen(false)} aria-label="关闭">×</button>
              <FilterRail
                tree={tree}
                selectedCat={selectedCatFull}
                onPickCat={onPickCat}
                facets={response?.facets || null}
                selectedKinds={filters.kinds}
                selectedSides={filters.sides}
                selectedAuthors={filters.authors}
                onToggle={onToggleFacet}
                onClearFacets={onClearFacets}
              />
              <ResourceList
                items={items}
                activePath={detail?.path || ''}
                onOpen={openResource}
                highlightTokens={highlightTokens}
                loading={loading}
              />
            </div>
            <Reader detail={detail} />
          </div>
          {drawerOpen && <div className="drawer-backdrop" onClick={() => setDrawerOpen(false)} />}
        </div>
      ) : tab === 'updates' ? (
        <RecentUpdates onOpen={async path => { await openResource(path); setTab('read'); }} />
      ) : tab === 'stats' ? (
        <SessionStats />
      ) : (
        <AdminPanel
          detail={detail}
          reload={() => load(filters)}
          onResourceMoved={() => { setDetail(null); load(filters); }}
          onBackToRead={() => setTab('read')}
        />
      )}
    </main>
  );
}
