import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { Detail } from '../../types';
import { copyText } from '../../library-state';
import { ActionPopover } from '../ui/ActionPopover';
import { buildDocument, resourceLink } from './adaptive';
import { AdaptiveSection } from './AdaptiveDocument';

type Props = {
  detail: Detail | null; loading?: boolean; error?: string; anchor?: string; query?: string; saved?: boolean;
  onBack?: () => void; onRetry?: () => void; onAnchor?: (id: string) => void; onBrowse?: () => void; onSave?: () => boolean | undefined;
  instance?: string; idPrefix?: string; compact?: boolean; paneLabel?: string;
};
type ReadingPreferences = { mode: 'adaptive' | 'raw'; find: string; findOpen: boolean; initialQuery: string };
type ViewAnchor = { id: string; offset: number };
const positions = new Map<string, number>();
const preferences = new Map<string, ReadingPreferences>();

export function Reader(props: Props) {
  if (props.error) return <article className="realm-reader">{props.onBack && <button className="realm-back" onClick={props.onBack}>← 返回结果</button>}<div className="realm-reader-state" role="alert"><h2>暂时无法打开这份档案</h2><p>资源可能已改名、下架，或网络暂不可用。不会自动替换成同名条目。</p><button onClick={props.onRetry}>重新加载</button><details><summary>错误详情</summary><pre>{props.error}</pre></details></div></article>;
  if (props.loading) return <article className="realm-reader" aria-busy="true">{props.onBack && <button className="realm-back" onClick={props.onBack}>← 返回结果</button>}<div className="realm-reader-state"><p role="status">正在展开档案…</p><div className="realm-skeleton"><div /><div /><div /></div></div></article>;
  if (!props.detail) return <article className="realm-reader realm-welcome">
    <div className="realm-overline">BETWEEN WORLDS · BEYOND LIMITS</div>
    <div className="realm-gate" aria-hidden="true"><i /><i /><span>∞</span></div>
    <div className="realm-overline">宏观界域 / 强化档案终端</div>
    <h2>诸界之间，<br />寻找你的下一种可能。</h2>
    <p>从职业、技能与特质中查阅完整规则。<br />跨越不同体系，不必迁就同一种模板。</p>
    <button className="realm-primary" onClick={props.onBrowse}>开始检索 <span aria-hidden="true">↗</span></button>
    <div className="realm-welcome-notes"><span><kbd>Ctrl / ⌘ K</kbd> 快速搜索</span><span>多词检索 · 拼音能力以搜索框提示为准</span><span>原文保留 · 规则不作改写</span></div>
  </article>;
  return <ReaderContent key={`${props.instance || 'primary'}:${props.detail.path}`} {...props} detail={props.detail} />;
}

function ReaderContent({ detail, anchor = '', query = '', saved, onBack, onAnchor, onSave, instance = 'primary', idPrefix = '', compact = false, paneLabel }: Props & { detail: Detail }) {
  const sections = useMemo(() => buildDocument(detail.content), [detail.content]);
  const identity = `${instance}:${detail.path}`;
  const remembered = useRef(preferences.get(identity));
  const [mode, setMode] = useState<'adaptive' | 'raw'>(remembered.current?.mode || 'adaptive');
  const sameQuery = remembered.current?.initialQuery === query;
  const [findOpen, setFindOpen] = useState(sameQuery ? remembered.current!.findOpen : Boolean(query));
  const [find, setFind] = useState(sameQuery ? remembered.current!.find : query);
  const [matchCount, setMatchCount] = useState(0);
  const [matchIndex, setMatchIndex] = useState(0);
  const [notice, setNotice] = useState('');
  const scroll = useRef<HTMLDivElement>(null);
  const title = useRef<HTMLHeadingElement>(null);
  const toc = useRef<HTMLDetailsElement>(null);
  const findInput = useRef<HTMLInputElement>(null);
  const matches = useRef<HTMLElement[]>([]);
  const activeQuery = findOpen ? find : '';
  const previousQuery = useRef<string | null>(remembered.current ? activeQuery : null);
  const initialSearchQuery = useRef(query);
  const timer = useRef<number | undefined>(undefined);
  const viewAnchor = useRef<ViewAnchor | null>(null);
  const previousCompact = useRef(compact);

  const reveal = (element: HTMLElement) => {
    const node = scroll.current;
    if (!node) return;
    node.scrollTop += element.getBoundingClientRect().top - node.getBoundingClientRect().top - 20;
  };
  const getSection = (id: string) => scroll.current?.querySelector<HTMLElement>(`#${CSS.escape(idPrefix + id)}`);
  const rememberViewport = () => {
    const node = scroll.current;
    if (!node?.clientHeight) return;
    const top = node.getBoundingClientRect().top;
    const visible = Array.from(node.querySelectorAll<HTMLElement>('.archive-section')).find(section => section.getBoundingClientRect().bottom > top + 20);
    viewAnchor.current = visible ? { id: visible.id, offset: visible.getBoundingClientRect().top - top } : null;
  };
  const jump = (id: string) => {
    const target = getSection(id);
    if (target) reveal(target);
    if (toc.current) toc.current.open = false;
    onAnchor?.(id);
  };
  useLayoutEffect(() => {
    const node = scroll.current;
    if (!node) return;
    const key = `${identity}:${mode}`;
    const target = anchor ? getSection(anchor) : null;
    if (target) reveal(target); else node.scrollTop = positions.get(key) || 0;
    rememberViewport();
    return () => { if (positions.size > 100) positions.delete(positions.keys().next().value!); positions.set(key, node.scrollTop); };
  }, [identity, mode]);
  // Keep the visible source section stationary when the decorative document
  // heading disappears/reappears. Do not remount the reader to change density.
  useLayoutEffect(() => {
    const node = scroll.current;
    const previous = viewAnchor.current;
    if (node && previous && previousCompact.current !== compact) {
      const target = node.querySelector<HTMLElement>(`#${CSS.escape(previous.id)}`);
      if (target) node.scrollTop += target.getBoundingClientRect().top - node.getBoundingClientRect().top - previous.offset;
    }
    previousCompact.current = compact;
    rememberViewport();
  }, [compact]);
  useEffect(() => {
    if (!anchor) return;
    const target = getSection(anchor);
    if (target) reveal(target);
  }, [anchor]);
  useEffect(() => {
    if (instance === 'primary' && window.matchMedia('(max-width: 820px)').matches) title.current?.focus({ preventScroll: true });
    return () => clearTimeout(timer.current);
  }, [detail.path, instance]);
  useEffect(() => {
    if (query !== initialSearchQuery.current) {
      initialSearchQuery.current = query;
      setFind(query); if (query) setFindOpen(true);
    }
  }, [query]);
  useEffect(() => {
    if (preferences.size > 60) preferences.delete(preferences.keys().next().value!);
    preferences.set(identity, { mode, find, findOpen, initialQuery: query });
  }, [identity, mode, find, findOpen, query]);
  useEffect(() => {
    matches.current = Array.from(scroll.current?.querySelectorAll<HTMLElement>('[data-reader-match]') || []);
    matches.current.forEach(node => node.removeAttribute('data-current'));
    setMatchCount(matches.current.length);
    setMatchIndex(0);
    if (matches.current[0]) {
      matches.current[0].dataset.current = 'true';
      if (activeQuery && previousQuery.current !== activeQuery && (!anchor || previousQuery.current !== null)) reveal(matches.current[0]);
    }
    previousQuery.current = activeQuery;
  }, [activeQuery, mode, sections]);
  useLayoutEffect(() => { if (findOpen) findInput.current?.focus({ preventScroll: true }); }, [findOpen]);

  const notify = (text: string) => { clearTimeout(timer.current); setNotice(text); timer.current = window.setTimeout(() => setNotice(''), 4500); };
  const copy = async (text: string, message = '已复制原文') => {
    try { await copyText(text); notify(message); }
    catch (error) { notify(error instanceof Error ? error.message : '复制失败，请在原文视图手动选择。'); }
  };
  const nextMatch = (step: number) => {
    if (!matches.current.length) return;
    matches.current[matchIndex]?.removeAttribute('data-current');
    const index = (matchIndex + step + matches.current.length) % matches.current.length;
    matches.current[index].dataset.current = 'true';
    setMatchIndex(index);
    reveal(matches.current[index]);
  };
  const modeControls = <div className="realm-mode" role="group" aria-label="阅读方式"><button aria-pressed={mode === 'adaptive'} onClick={() => setMode('adaptive')}>自适应</button><button aria-pressed={mode === 'raw'} onClick={() => setMode('raw')}>原文</button></div>;
  const saveControl = onSave && <button aria-pressed={!!saved} onClick={() => { const ok = onSave(); notify(ok === false ? '浏览器未能持久保存，本次更改仅在当前页面有效。' : saved ? '已取消收藏' : '已加入随行档案，仅保存在此浏览器'); }}>{saved ? '★ 已收藏' : '☆ 收藏'}</button>;
  const copyControls = <><button data-close-popover onClick={() => copy(detail.content, '已复制完整原文')}>复制全文</button><button data-close-popover onClick={() => copy(resourceLink(detail.path), '已复制档案链接')}>分享档案</button></>;
  const fileInformation = <details><summary>档案信息</summary><div><p>{detail.path}</p><p>{detail.encoding} · {detail.size} 字节</p><button data-close-popover onClick={() => copy(detail.path, '已复制文件路径')}>复制路径</button></div></details>;

  return <article className={`realm-reader${compact ? ' compact-reader' : ''}`}>
    <div className="realm-reader-toolbar">
      {onBack && <button className="realm-back" aria-label="返回结果" onClick={onBack}>{compact ? '←' : '← 结果'}</button>}
      {compact && <div className="reader-identity">{paneLabel && <span className="reader-pane-badge">{paneLabel}</span>}<h1 ref={title} className="reader-current-title" tabIndex={-1} title={detail.title}>{detail.title}</h1></div>}
      <details ref={toc} className="realm-toc"><summary>目录 <span>{sections.length}</span></summary><nav aria-label="文档目录">{sections.map(section => <button key={section.id} onClick={() => jump(section.id)}>{section.title}</button>)}</nav></details>
      <button aria-label="文内查找" aria-expanded={findOpen} onClick={() => setFindOpen(value => !value)}>{compact ? '查找' : '文内查找'}</button>
      {compact ? <ActionPopover label="阅读工具" trigger="更多">
        <p className="reader-tools-title">{detail.title}</p>
        <div className="reader-tools-row">{modeControls}{saveControl}</div>
        <div className="reader-tools-row">{copyControls}</div>
        {fileInformation}
        <p className="reader-tools-note">{detail.category}{detail.authors?.length ? ` · ${detail.authors.join(' / ')}` : ''}</p>
        <p className="reader-tools-note">自适应仅辅助排版；非标准内容按原顺序保留，规则以原文为准。</p>
      </ActionPopover> : <>{modeControls}{saveControl}</>}
    </div>
    {findOpen && <div className="realm-find"><input ref={findInput} type="search" aria-label="文内精确查找" placeholder="在当前正文精确查找…" value={find} onChange={event => setFind(event.target.value)} onKeyDown={event => { if (event.nativeEvent.isComposing) return; if (event.key === 'Enter') { event.preventDefault(); nextMatch(event.shiftKey ? -1 : 1); } if (event.key === 'Escape') setFindOpen(false); }} /><span role="status">{matchCount ? `${matchIndex + 1} / ${matchCount}` : find ? '无精确匹配' : '输入关键词'}</span><button disabled={!matchCount} aria-label="上一个匹配" onClick={() => nextMatch(-1)}>↑</button><button disabled={!matchCount} aria-label="下一个匹配" onClick={() => nextMatch(1)}>↓</button><button aria-label="关闭文内查找" onClick={() => setFindOpen(false)}>×</button></div>}
    <div ref={scroll} className="realm-reader-scroll" onScroll={rememberViewport}>
      <header className="realm-document-heading" hidden={compact}><div className="realm-overline">RESOURCE ARCHIVE / {detail.top_kind || '序列档案'}</div><h1 ref={compact ? undefined : title} tabIndex={-1}>{detail.title}</h1>
        <div className="realm-document-meta">{detail.side && <span>{detail.side}</span>}<span>{detail.category}</span>{detail.authors?.length ? <span>{detail.authors.join(' / ')}</span> : null}</div>
        <div className="realm-document-actions">{copyControls}{fileInformation}</div>
      </header>
      <div className="realm-reading-note" hidden={compact}>{mode === 'adaptive' ? '按明确字段辅助排版；非标准内容按原有顺序保留。规则以原文为准。' : '原文视图：保留完整内容、换行与顺序。'}</div>
      <div className={`realm-document-content${mode === 'raw' ? ' is-raw' : ''}`}>
        {sections.length ? sections.map(section => <AdaptiveSection key={section.id} section={section} idPrefix={idPrefix} rawMode={mode === 'raw'} query={activeQuery} onCopy={text => copy(text)} onShare={id => copy(resourceLink(detail.path, id), '已复制条目定位链接')} />) : <p>这份档案暂无正文。</p>}
      </div>
      <footer className="realm-document-footer">END OF ARCHIVE <span>宏观界域 · 强化序列库</span><button onClick={() => { if (scroll.current) scroll.current.scrollTop = 0; onAnchor?.(''); }}>回到顶部 ↑</button></footer>
    </div>
    <div className={`realm-toast${notice ? ' visible' : ''}`} role="status" aria-live="polite">{notice}</div>
  </article>;
}
