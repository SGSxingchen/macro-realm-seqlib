import { useCallback, useEffect, useLayoutEffect, useRef, useState } from 'react';
import { api, buildQuery, routePath } from './api';
import type { Detail, ResourceListResponse, SearchFilters, TreeNode } from './types';
import { Header } from './components/Header';
import { SearchBar } from './components/SearchBar';
import { FilterRail } from './components/FilterRail';
import { ResourceList } from './components/ResourceList';
import { ReaderWorkbench } from './components/Reader/Workbench';
import { AdminPanel } from './components/Admin';
import { RecentUpdates } from './components/RecentUpdates';
import { NormalizationReviewPage } from './components/NormalizationReview';
import { SessionStats } from './components/SessionStats';
import { emptyFilters, loadSaved, locationUrl, readLocation, saveResources } from './library-state';
import type { LocationState } from './library-state';
import { readPanePreferences, storePanePreferences } from './workbench-state';

export function App() {
  return window.location.pathname === '/normalize-review' ? <NormalizationReviewPage /> : <LibraryApp />;
}

function LibraryApp() {
  const [nav, setNav] = useState(readLocation);
  const navRef = useRef(nav);
  const [result, setResult] = useState<{ key: string; data: ResourceListResponse } | null>(null);
  const [loading, setLoading] = useState(true);
  const [listError, setListError] = useState('');
  const [tree, setTree] = useState<TreeNode[]>([]);
  const [treeError, setTreeError] = useState(false);
  const [detail, setDetail] = useState<Detail | null>(null);
  const [detailLoading, setDetailLoading] = useState(false);
  const [detailError, setDetailError] = useState<{ path: string; message: string } | null>(null);
  const [refresh, setRefresh] = useState(0);
  const [drawer, setDrawer] = useState(false);
  const [saved, setSaved] = useState(loadSaved);
  const [storageNotice, setStorageNotice] = useState('');
  const [panes, setPanes] = useState(readPanePreferences);
  const [focus, setFocus] = useState(false);
  const [compareFocus, setCompareFocus] = useState(false);
  const [searchFocusTick, setSearchFocusTick] = useState(0);
  // Comparison can free the sidebars without pretending to be explicit focus.
  // With no document open, always leave a usable search surface.
  const readingFocus = Boolean(nav.openPath) && focus;
  const focused = Boolean(nav.openPath) && (focus || compareFocus);
  const showIndex = panes.index && !focused;
  const showResults = (!nav.openPath || panes.results) && !focused;
  const dialogRef = useRef<HTMLDialogElement>(null);
  const dialogReturnFocus = useRef<HTMLButtonElement | null>(null);
  const listRequest = useRef<AbortController | null>(null);
  const pending = useRef(false);
  const focusSearchAfterNavigation = useRef(false);
  const nextOffset = useRef(0);
  const detailCache = useRef(new Map<string, Detail>());
  const key = JSON.stringify(nav.filters);
  const response = result?.key === key ? result.data : null;
  const items = response?.items || [];
  const searching = loading || result?.key !== key;

  const navigate = useCallback((patch: Partial<LocationState>, replace = false) => {
    const previous = navRef.current;
    const next = { ...previous, ...patch };
    navRef.current = next;
    setNav(next);
    const url = locationUrl(next);
    if (url !== window.location.pathname + window.location.search + window.location.hash) {
      window.history[replace ? 'replaceState' : 'pushState']({ ...window.history.state, realm: true, ...(!replace ? { realmFromResults: previous.tab === 'read' && !previous.openPath && !!next.openPath } : {}) }, '', url);
    }
  }, []);

  useEffect(() => { storePanePreferences(panes); }, [panes]);
  useEffect(() => {
    const restore = () => { const next = readLocation(); navRef.current = next; setNav(next); setDrawer(false); };
    window.addEventListener('popstate', restore);
    window.addEventListener('hashchange', restore);
    return () => { window.removeEventListener('popstate', restore); window.removeEventListener('hashchange', restore); };
  }, []);

  const loadPage = useCallback((offset: number, append: boolean) => {
    listRequest.current?.abort();
    const controller = new AbortController();
    listRequest.current = controller;
    pending.current = true;
    setLoading(true); setListError('');
    const f = navRef.current.filters;
    const requestKey = JSON.stringify(f);
    const qs = buildQuery({ q: f.q, category: f.category, kinds: f.kinds, sides: f.sides, authors: f.authors, include_content: false, limit: 100, offset });
    api<ResourceListResponse>('/api/resources' + qs, { signal: controller.signal }).then(data => {
      if (controller.signal.aborted || requestKey !== JSON.stringify(navRef.current.filters)) return;
      nextOffset.current = data.items.length ? data.offset + data.items.length : data.total;
      setResult(previous => ({ key: requestKey, data: { ...data, items: append && previous?.key === requestKey ? [...previous.data.items, ...data.items] : data.items } }));
    }).catch(error => {
      if (!controller.signal.aborted) setListError(error instanceof Error ? error.message : '请求失败');
    }).finally(() => {
      if (!controller.signal.aborted) { setLoading(false); pending.current = false; }
    });
  }, []);

  useEffect(() => {
    listRequest.current?.abort(); pending.current = false; setListError(''); setLoading(true);
    const timer = window.setTimeout(() => loadPage(0, false), navRef.current.filters.q ? 220 : 0);
    return () => { clearTimeout(timer); listRequest.current?.abort(); };
  }, [key, refresh, loadPage]);

  useEffect(() => {
    const controller = new AbortController(); setTreeError(false);
    api<{ items: TreeNode[] }>('/api/tree', { signal: controller.signal }).then(r => setTree(r.items)).catch(() => { if (!controller.signal.aborted) setTreeError(true); });
    return () => controller.abort();
  }, [refresh]);

  useEffect(() => {
    const path = nav.openPath;
    const controller = new AbortController(); setDetailError(null);
    if (!path) { setDetail(null); setDetailLoading(false); return () => controller.abort(); }
    const cached = detailCache.current.get(path);
    if (cached) { setDetail(cached); setDetailLoading(false); return () => controller.abort(); }
    setDetailLoading(true);
    api<Detail>('/api/resources/' + routePath(path), { signal: controller.signal }).then(data => {
      if (controller.signal.aborted || navRef.current.openPath !== path) return;
      if (detailCache.current.size >= 24) detailCache.current.delete(detailCache.current.keys().next().value!);
      detailCache.current.set(path, data); setDetail(data);
    }).catch(error => {
      if (!controller.signal.aborted) setDetailError({ path, message: error instanceof Error ? error.message : '请求失败' });
    }).finally(() => { if (!controller.signal.aborted) setDetailLoading(false); });
    return () => controller.abort();
  }, [nav.openPath, refresh]);

  useEffect(() => {
    const dialog = dialogRef.current;
    if (!dialog) return;
    if (drawer && !dialog.open) dialog.showModal();
    if (!drawer && dialog.open) {
      dialog.close();
      if (focusSearchAfterNavigation.current) {
        focusSearchAfterNavigation.current = false;
        document.querySelector<HTMLInputElement>('#realm-query')?.focus();
      } else if (dialogReturnFocus.current?.getBoundingClientRect().width) dialogReturnFocus.current.focus();
    }
  }, [drawer]);
  useEffect(() => {
    const mq = window.matchMedia('(min-width: 1181px)');
    const resize = () => { if (mq.matches) setDrawer(false); };
    mq.addEventListener('change', resize);
    return () => mq.removeEventListener('change', resize);
  }, []);
  useEffect(() => {
    const escapeFocus = (event: KeyboardEvent) => {
      if (event.key !== 'Escape' || event.defaultPrevented || !readingFocus || drawer) return;
      if (event.target instanceof HTMLElement && event.target.closest('input, textarea, select, [contenteditable="true"], details[open]')) return;
      setFocus(false);
    };
    window.addEventListener('keydown', escapeFocus);
    return () => window.removeEventListener('keydown', escapeFocus);
  }, [readingFocus, drawer]);

  const revealSearch = useCallback(() => {
    focusSearchAfterNavigation.current = true;
    dialogReturnFocus.current = null; setDrawer(false);
    setFocus(false); setCompareFocus(false);
    setPanes(value => ({ ...value, results: true }));
    setSearchFocusTick(value => value + 1);
    if (window.matchMedia('(max-width: 820px)').matches) navigate({ openPath: '', anchor: '' }, true);
  }, [navigate]);
  useLayoutEffect(() => {
    if (focusSearchAfterNavigation.current && !dialogRef.current?.open && nav.tab === 'read' && showResults && (!nav.openPath || window.matchMedia('(min-width: 821px)').matches)) {
      focusSearchAfterNavigation.current = false;
      document.querySelector<HTMLInputElement>('#realm-query')?.focus();
    }
  }, [nav.openPath, nav.tab, showResults, searchFocusTick]);
  useEffect(() => {
    const searchShortcut = (event: KeyboardEvent) => {
      if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k' && navRef.current.tab === 'read') {
        event.preventDefault(); revealSearch();
      }
    };
    window.addEventListener('keydown', searchShortcut);
    return () => window.removeEventListener('keydown', searchShortcut);
  }, [revealSearch]);

  const backToResults = () => {
    setFocus(false); setCompareFocus(false); setPanes(value => ({ ...value, results: true }));
    if (window.history.state?.realmFromResults) window.history.back();
    else navigate({ openPath: '', anchor: '' }, true);
  };
  const openResource = (path: string) => { navigate({ tab: 'read', openPath: path, anchor: '' }); setDrawer(false); };
  const changeFilters = (filters: SearchFilters) => {
    const phone = window.matchMedia('(max-width: 820px)').matches;
    navigate({ filters, ...(phone ? { openPath: '', anchor: '' } : {}) }, true);
  };
  const toggleFacet = (group: 'kinds' | 'sides' | 'authors', name: string) => {
    const f = navRef.current.filters;
    changeFilters({ ...f, [group]: f[group].includes(name) ? f[group].filter(v => v !== name) : [...f[group], name] });
  };
  const reload = () => { detailCache.current.clear(); setRefresh(v => v + 1); };
  const currentDetail = detail?.path === nav.openPath ? detail : null;
  const currentError = detailError?.path === nav.openPath ? detailError.message : '';
  const selectedCount = nav.filters.kinds.length + nav.filters.sides.length + nav.filters.authors.length + Number(Boolean(nav.filters.category));
  const filterRail = <FilterRail tree={tree} selectedCat={nav.filters.category ? `序列库/${nav.filters.category}` : ''}
    onPickCat={path => changeFilters({ ...navRef.current.filters, category: path.startsWith('序列库/') ? path.slice('序列库/'.length) : '' })}
    facets={response?.facets || null} selectedKinds={nav.filters.kinds} selectedSides={nav.filters.sides} selectedAuthors={nav.filters.authors}
    onToggle={toggleFacet} onClearFacets={() => changeFilters({ ...emptyFilters(), q: nav.filters.q })} />;
  const openFilters = (button: HTMLButtonElement) => { dialogReturnFocus.current = button; setDrawer(true); };
  const togglePane = (pane: 'index' | 'results') => {
    const shown = pane === 'index' ? showIndex : showResults;
    setFocus(false); setCompareFocus(false);
    setPanes(value => ({ ...value, [pane]: !shown }));
  };

  return <main className={`app realm-app ${nav.tab === 'read' ? `realm-library${readingFocus ? ' realm-focus-reading' : ''}` : nav.tab === 'admin' ? 'admin-mode' : ''}`}>
    <Header recordCount={tree.reduce((sum, node) => sum + node.count, 0)} tab={nav.tab} onTab={tab => { setDrawer(false); navigate({ tab }); }} />
    {nav.tab === 'read' ? <div className="realm-reading-shell">
      <div className="realm-layout-controls" aria-label="阅读布局">
        <span className="realm-overline">READING DESK</span>
        <button className="realm-index-toggle" aria-expanded={showIndex} aria-controls="realm-index-panel" onClick={() => togglePane('index')}>{showIndex ? '收起索引' : '展开索引'}</button>
        <button className="realm-global-filter" aria-haspopup="dialog" onClick={event => openFilters(event.currentTarget)}>索引 / 收藏</button>
        <button className="realm-results-toggle" disabled={!nav.openPath} title={!nav.openPath ? '打开档案后可收起结果' : undefined} aria-expanded={showResults} aria-controls="realm-results-panel" onClick={() => togglePane('results')}>{showResults ? '收起结果' : '展开结果'}</button>
        <button className="realm-focus-toggle" aria-pressed={readingFocus} disabled={!nav.openPath} title={!nav.openPath ? '打开一份档案后进入专注阅读' : '专注阅读，按 Esc 或再次点击退出'} onClick={() => setFocus(value => !value)}>{readingFocus ? '退出专注' : '专注阅读'}</button>
        <button className="realm-search-toggle" onClick={revealSearch}>搜索档案 <kbd>⌘ / Ctrl K</kbd></button>
      </div>
      <div className={`realm-workspace ${nav.openPath ? 'has-resource' : ''} ${showIndex ? '' : 'index-hidden'} ${showResults ? '' : 'results-hidden'}`}>
        <aside id="realm-index-panel" className="realm-navigation" aria-label="档案筛选">
          <div className="realm-panel-heading"><span className="realm-overline">ARCHIVE INDEX</span><h2>界域索引</h2></div>
          {treeError && <p role="alert" className="realm-error">分类暂不可用 <button onClick={reload}>重试</button></p>}
          {filterRail}
          <div className="realm-saved"><h3>随行档案 <span>{saved.length}</span></h3><p>仅保存在此浏览器</p>
            {saved.length ? saved.map(item => <div key={item.path}><button className="saved-title" onClick={() => openResource(item.path)}>{item.title}</button><button aria-label={`取消收藏：${item.title}`} onClick={() => { const next = saved.filter(v => v.path !== item.path); setSaved(next); if (!saveResources(next)) setStorageNotice('收藏未能保存，关闭页面后可能丢失。'); }}>×</button></div>) : <p>阅读时点「收藏」，把常用规则带在身边。</p>}
            {storageNotice && <p role="status">{storageNotice}</p>}
          </div>
        </aside>
        <section id="realm-results-panel" className="realm-results" aria-label="搜索结果">
          <div className="realm-search-head"><div className="realm-panel-heading"><span className="realm-overline">SEQUENCE LIBRARY</span><h2>序列检索</h2></div>
            <button className="realm-filter-trigger" onClick={event => openFilters(event.currentTarget)} aria-haspopup="dialog">筛选{selectedCount ? ` · ${selectedCount}` : ''}</button>
          </div>
          <SearchBar value={nav.filters.q} onChange={q => changeFilters({ ...navRef.current.filters, q })} count={response?.total ?? 0} searching={searching && !listError} pinyinReady={!!response?.engine?.pinyin} onClear={() => changeFilters({ ...navRef.current.filters, q: '' })} />
          <div className="realm-filter-chips" aria-label="已选条件">
            {nav.filters.category && <button onClick={() => changeFilters({ ...nav.filters, category: '' })}>{nav.filters.category} ×</button>}
            {(['kinds', 'sides', 'authors'] as const).flatMap(group => nav.filters[group].map(value => <button key={`${group}-${value}`} onClick={() => toggleFacet(group, value)}>{value} ×</button>))}
            {(selectedCount > 0 || nav.filters.q) && <button className="realm-reset" onClick={() => changeFilters(emptyFilters())}>重置全部</button>}
          </div>
          <div className="realm-results-meta" role="status">{searching && !listError ? '正在检索档案…' : `找到 ${response?.total ?? 0} 份档案`}<span>公开序列库</span></div>
          {listError && <div role="alert" className="realm-error">检索失败，未将错误当作空结果。<button onClick={() => loadPage(response ? nextOffset.current : 0, !!response)}>重试</button><details><summary>错误详情</summary>{listError}</details></div>}
          <ResourceList key={key} items={items} activePath={nav.openPath} onOpen={openResource} highlightTokens={response?.tokens || []} loading={searching && !listError} scrollKey={key} error={!!listError} />
          {response && nextOffset.current < response.total && <button className="realm-load-more" disabled={loading} onClick={() => { if (!pending.current) loadPage(nextOffset.current, true); }}>{loading ? '加载中…' : `加载更多 · 已显示 ${items.length} / ${response.total}`}</button>}
        </section>
        <ReaderWorkbench openPath={nav.openPath} onOpen={openResource} onFocusMode={setCompareFocus} revision={refresh}
          detail={currentDetail} loading={Boolean(nav.openPath && (detailLoading || (!currentDetail && !currentError)))} error={currentError}
          anchor={nav.anchor} query={nav.filters.q} onBack={backToResults} onRetry={reload} onAnchor={anchor => navigate({ anchor }, true)} onBrowse={revealSearch}
          saved={saved.some(item => item.path === nav.openPath)} onSave={() => {
            if (!currentDetail) return;
            const next = saved.some(item => item.path === currentDetail.path) ? saved.filter(item => item.path !== currentDetail.path) : [...saved.slice(-99), { path: currentDetail.path, title: currentDetail.title }];
            setSaved(next); const ok = saveResources(next); setStorageNotice(ok ? '' : '收藏未能保存，关闭页面后可能丢失。'); return ok;
          }} />
        <dialog ref={dialogRef} className="realm-filter-dialog" aria-labelledby="realm-filter-title" onCancel={e => { e.preventDefault(); setDrawer(false); }} onClose={() => setDrawer(false)} onClick={e => {
          const rect = e.currentTarget.getBoundingClientRect();
          if (e.target === e.currentTarget && (e.clientX < rect.left || e.clientX > rect.right || e.clientY < rect.top || e.clientY > rect.bottom)) setDrawer(false);
        }}>
          <header><h2 id="realm-filter-title">筛选档案</h2><button autoFocus onClick={() => setDrawer(false)} aria-label="关闭筛选">×</button></header>
          {treeError && <p role="alert">分类暂不可用 <button onClick={reload}>重试</button></p>}{filterRail}
          {saved.length > 0 && <div className="realm-saved"><h3>随行档案</h3>{saved.map(item => <div key={item.path}><button className="saved-title" onClick={() => openResource(item.path)}>{item.title}</button><button aria-label={`取消收藏：${item.title}`} onClick={() => { const next = saved.filter(v => v.path !== item.path); setSaved(next); if (!saveResources(next)) setStorageNotice('收藏未能保存。'); }}>×</button></div>)}</div>}
          {storageNotice && <p role="status">{storageNotice}</p>}
          <button className="realm-primary" onClick={() => { setDrawer(false); if (window.matchMedia('(max-width: 820px)').matches) revealSearch(); }}>查看结果{response ? ` · ${response.total}` : ''}</button>
        </dialog>
      </div>
    </div> : nav.tab === 'updates' ? <RecentUpdates onOpen={async path => openResource(path)} /> : nav.tab === 'stats' ? <SessionStats /> : <AdminPanel detail={detail} reload={reload} onResourceMoved={() => { navigate({ openPath: '', anchor: '' }, true); reload(); }} onBackToRead={() => navigate({ tab: 'read' })} />}
  </main>;
}
