import { useEffect, useLayoutEffect, useMemo, useRef, useState } from 'react';
import type { Detail } from '../../types';
import { copyText } from '../../library-state';
import { buildDocument, resourceLink } from './adaptive';
import { AdaptiveSection, extractRank } from './AdaptiveDocument';
import { ResponsiveMenu } from '../ResponsiveMenu';
import { sideKey } from '../ResourceCard';

export type WelcomeHubData = {
  totalCount: number;
  categories: Array<{ name: string; path: string; count: number }>;
  sides: Array<{ name: string; count: number }>;
  savedItems: Array<{ path: string; title: string }>;
  sampleItems: Array<{ path: string; title: string; top_kind?: string; side?: string }>;
  onPickCategory: (category: string) => void;
  onPickSide: (side: string) => void;
  onOpenResource: (path: string) => void;
  onOpenRandom: () => void;
};

type Props = {
  detail: Detail | null;
  loading?: boolean;
  error?: string;
  anchor?: string;
  query?: string;
  saved?: boolean;
  onBack?: () => void;
  onRetry?: () => void;
  onAnchor?: (id: string) => void;
  onBrowse?: () => void;
  onSave?: () => boolean | undefined;
  instance?: string;
  idPrefix?: string;
  welcome?: WelcomeHubData;
};

type ReadingPreferences = { mode: 'adaptive' | 'raw'; find: string; findOpen: boolean; initialQuery: string };
const positions = new Map<string, number>();
const preferences = new Map<string, ReadingPreferences>();

const SIDE_DESCRIPTIONS: Record<string, string> = {
  '战技侧': '武技、兵器、体术与正面交锋体系',
  '神秘侧': '咒术、仪式、血脉、神秘学与异能体系',
  '科技侧': '义体、军工、力场、智能与科学造物',
  '特殊侧': '跨体系机制、概念特质与特殊规则序列',
};

export function Reader(props: Props) {
  if (props.error) {
    return (
      <article className="realm-reader">
        {props.onBack && <button className="realm-back" onClick={props.onBack}>← 返回结果</button>}
        <div className="realm-reader-state" role="alert">
          <h2>暂时无法打开这份档案</h2>
          <p>资源可能已改名、下架，或网络暂不可用。不会自动替换成同名条目。</p>
          <button onClick={props.onRetry}>重新加载</button>
          <details><summary>错误详情</summary><pre>{props.error}</pre></details>
        </div>
      </article>
    );
  }
  if (props.loading) {
    return (
      <article className="realm-reader" aria-busy="true">
        {props.onBack && <button className="realm-back" onClick={props.onBack}>← 返回结果</button>}
        <div className="realm-reader-state">
          <p role="status">正在展开档案…</p>
          <div className="realm-skeleton"><div /><div /><div /></div>
        </div>
      </article>
    );
  }
  if (!props.detail) {
    const w = props.welcome;
    return (
      <article className="realm-reader realm-welcome">
        <div className="realm-welcome-container">
          <header className="realm-welcome-hero">
            <div className="realm-welcome-badge">
              <span className="welcome-sigil" aria-hidden="true">∞</span>
              <span>宏观界域 · 强化序列档案库{w?.totalCount ? ` · ${w.totalCount} 份公开档案` : ''}</span>
            </div>
            <h2>诸界之间，查阅你的下一份强化序列</h2>
            <p>完整收录职业、技能表、特质改造、能量池与公共建筑原文。结构化速读与无损原文随时切换，支持多档案并排对照。</p>
            <div className="realm-welcome-cta">
              <button className="realm-primary" onClick={props.onBrowse}>开始检索 <kbd>⌘K</kbd></button>
              {w && w.sampleItems.length > 0 && (
                <button type="button" className="realm-secondary-cta" onClick={w.onOpenRandom}>随机抽选档案</button>
              )}
            </div>
          </header>

          {w && w.sides.length > 0 && (
            <section className="realm-welcome-section" aria-label="按作品侧浏览">
              <h3>按作品侧快速筛选</h3>
              <div className="welcome-side-grid">
                {w.sides.map(side => (
                  <button
                    key={side.name}
                    type="button"
                    className="welcome-side-card"
                    data-side={sideKey(side.name)}
                    onClick={() => w.onPickSide(side.name)}
                  >
                    <div className="welcome-side-top">
                      <strong>{side.name}</strong>
                      <span>{side.count} 份</span>
                    </div>
                    <small>{SIDE_DESCRIPTIONS[side.name] || '点击筛选该侧别下的全部档案'}</small>
                  </button>
                ))}
              </div>
            </section>
          )}

          {w && w.categories.length > 0 && (
            <section className="realm-welcome-section" aria-label="按分类浏览">
              <h3>核心序列大类</h3>
              <div className="welcome-cat-pills">
                {w.categories.map(cat => (
                  <button key={cat.path} type="button" onClick={() => w.onPickCategory(cat.path)}>
                    <span>{cat.name}</span>
                    <em>{cat.count}</em>
                  </button>
                ))}
              </div>
            </section>
          )}

          {w && (w.savedItems.length > 0 || w.sampleItems.length > 0) && (() => {
            const quickItems: Array<{ path: string; title: string; side?: string }> =
              w.savedItems.length > 0 ? w.savedItems.slice(0, 6) : w.sampleItems.slice(0, 6);
            return (
              <section className="realm-welcome-section" aria-label={w.savedItems.length > 0 ? '随行收藏' : '档案直达'}>
                <h3>{w.savedItems.length > 0 ? '随行收藏直达' : '当前列表速览'}</h3>
                <div className="welcome-quick-list">
                  {quickItems.map(item => (
                    <button key={item.path} type="button" className="welcome-quick-item" onClick={() => w.onOpenResource(item.path)}>
                      <span className="quick-item-title">{item.title}</span>
                      {item.side ? (
                        <small data-side={sideKey(item.side)}>{item.side}</small>
                      ) : (
                        <small>打开 ↵</small>
                      )}
                    </button>
                  ))}
                </div>
              </section>
            );
          })()}

          <footer className="realm-welcome-notes">
            <span><kbd>Ctrl / ⌘ K</kbd> 聚焦全局搜索</span>
            <span>支持多词组合与拼音检索</span>
            <span>原文逐字保留 · 不改写任何规则</span>
          </footer>
        </div>
      </article>
    );
  }
  return <ReaderContent key={`${props.instance || 'primary'}:${props.detail.path}`} {...props} detail={props.detail} />;
}

function cleanOutlineTitle(title: string): string {
  return title.replace(/[（(\[【]\s*(?:EX|[SABCDEF][+-]?)\s*级?\s*[）)\]】]/gi, '').trim() || title;
}

function ReaderContent({
  detail,
  anchor = '',
  query = '',
  saved,
  onBack,
  onAnchor,
  onSave,
  instance = 'primary',
  idPrefix = '',
}: Props & { detail: Detail }) {
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
  const [activeSectionId, setActiveSectionId] = useState(() => anchor || sections[0]?.id || '');
  const scroll = useRef<HTMLDivElement>(null);
  const title = useRef<HTMLHeadingElement>(null);
  const toc = useRef<HTMLDetailsElement>(null);
  const findInput = useRef<HTMLInputElement>(null);
  const matches = useRef<HTMLElement[]>([]);
  const activeQuery = findOpen ? find : '';
  const previousQuery = useRef<string | null>(remembered.current ? activeQuery : null);
  const initialSearchQuery = useRef(query);
  const timer = useRef<number | undefined>(undefined);

  const reveal = (element: HTMLElement) => {
    const node = scroll.current;
    if (!node) return;
    node.scrollTop += element.getBoundingClientRect().top - node.getBoundingClientRect().top - 20;
  };
  const getSection = (id: string) => scroll.current?.querySelector<HTMLElement>(`#${CSS.escape(idPrefix + id)}`);
  const jump = (id: string) => {
    const target = getSection(id);
    if (target) reveal(target);
    if (toc.current) toc.current.open = false;
    setActiveSectionId(id);
    onAnchor?.(id);
  };

  useLayoutEffect(() => {
    const node = scroll.current;
    if (!node) return;
    const key = `${identity}:${mode}`;
    const target = anchor ? getSection(anchor) : null;
    if (target) reveal(target);
    else node.scrollTop = positions.get(key) || 0;
    return () => {
      if (positions.size > 100) positions.delete(positions.keys().next().value!);
      positions.set(key, node.scrollTop);
    };
  }, [identity, mode]);

  useEffect(() => {
    if (!anchor) return;
    const target = getSection(anchor);
    if (target) {
      reveal(target);
      setActiveSectionId(anchor);
    }
  }, [anchor]);

  useEffect(() => {
    const node = scroll.current;
    if (!node || sections.length < 2) return;
    let raf = 0;
    const updateActive = () => {
      raf = 0;
      const topLimit = node.getBoundingClientRect().top + 96;
      let current = sections[0]?.id || '';
      for (const sec of sections) {
        const el = node.querySelector<HTMLElement>(`#${CSS.escape(idPrefix + sec.id)}`);
        if (!el || el.clientHeight === 0) continue;
        if (el.getBoundingClientRect().top <= topLimit) current = sec.id;
        else break;
      }
      setActiveSectionId(prev => (prev === current ? prev : current));
    };
    const onScroll = () => {
      if (!raf) raf = requestAnimationFrame(updateActive);
    };
    node.addEventListener('scroll', onScroll, { passive: true });
    updateActive();
    return () => {
      node.removeEventListener('scroll', onScroll);
      if (raf) cancelAnimationFrame(raf);
    };
  }, [sections, idPrefix, mode]);

  useEffect(() => {
    if (instance === 'primary' && window.matchMedia('(max-width: 820px)').matches) title.current?.focus({ preventScroll: true });
    return () => clearTimeout(timer.current);
  }, [detail.path, instance]);

  useEffect(() => {
    const dismiss = (event: PointerEvent) => {
      if (toc.current?.open && !toc.current.contains(event.target as Node)) toc.current.open = false;
    };
    document.addEventListener('pointerdown', dismiss);
    return () => document.removeEventListener('pointerdown', dismiss);
  }, []);

  useEffect(() => {
    if (query !== initialSearchQuery.current) {
      initialSearchQuery.current = query;
      setFind(query);
      if (query) setFindOpen(true);
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

  const notify = (text: string) => {
    clearTimeout(timer.current);
    setNotice(text);
    timer.current = window.setTimeout(() => setNotice(''), 4500);
  };
  const copy = async (text: string, message = '已复制原文') => {
    try {
      await copyText(text);
      notify(message);
    } catch (error) {
      notify(error instanceof Error ? error.message : '复制失败，请在原文视图手动选择。');
    }
  };
  const downloadTxt = () => {
    let url: string | undefined;
    const link = document.createElement('a');
    try {
      const filename = detail.path.split('/').pop() || detail.title || '资源';
      link.download = filename.replace(/\.[^.]+$/, '').replace(/[<>:"/\\|?*\u0000-\u001f]/g, '_') + '.txt';
      url = URL.createObjectURL(new Blob([detail.content], { type: 'text/plain;charset=utf-8' }));
      link.href = url;
      document.body.appendChild(link);
      link.click();
      notify('已发起 TXT 下载');
    } catch (error) {
      console.error('资源 TXT 下载失败', error);
      notify('下载失败，请重试或使用复制全文。');
    } finally {
      link.remove();
      if (url) {
        const downloadUrl = url;
        window.setTimeout(() => URL.revokeObjectURL(downloadUrl), 30000);
      }
    }
  };
  const nextMatch = (step: number) => {
    if (!matches.current.length) return;
    matches.current[matchIndex]?.removeAttribute('data-current');
    const index = (matchIndex + step + matches.current.length) % matches.current.length;
    matches.current[index].dataset.current = 'true';
    setMatchIndex(index);
    reveal(matches.current[index]);
  };

  const seqMatch = detail.title.match(/^(\d{1,4}】)(.+)$/);
  const showDesktopOutline = instance === 'primary' && mode === 'adaptive' && sections.length >= 2;

  return (
    <article className="realm-reader">
      <div className="realm-reader-toolbar">
        {onBack && <button className="realm-back" onClick={onBack}>← 结果</button>}
        <details
          ref={toc}
          className="realm-toc"
          onKeyDown={event => {
            if (event.key === 'Escape' && event.currentTarget.open) {
              event.preventDefault();
              event.stopPropagation();
              event.currentTarget.open = false;
              event.currentTarget.querySelector('summary')?.focus();
            }
          }}
        >
          <summary>目录 <span>{sections.length}</span></summary>
          <nav aria-label="文档目录">
            {sections.map(section => (
              <button key={section.id} className={activeSectionId === section.id ? 'is-active' : undefined} onClick={() => jump(section.id)}>
                {section.title}
              </button>
            ))}
          </nav>
        </details>
        <button
          aria-expanded={findOpen}
          onClick={() => {
            setFindOpen(v => !v);
            requestAnimationFrame(() => findInput.current?.focus());
          }}
        >
          文内查找
        </button>
        <ResponsiveMenu label="阅读设置" className="realm-reader-options">
          <div className="realm-mode" role="group" aria-label="阅读方式">
            <button aria-pressed={mode === 'adaptive'} onClick={() => setMode('adaptive')}>自适应</button>
            <button aria-pressed={mode === 'raw'} onClick={() => setMode('raw')}>原文</button>
          </div>
          {onSave && (
            <button
              aria-pressed={!!saved}
              onClick={() => {
                const ok = onSave();
                notify(ok === false ? '浏览器未能持久保存，本次更改仅在当前页面有效。' : saved ? '已取消收藏' : '已加入随行档案，仅保存在此浏览器');
              }}
            >
              {saved ? '★ 已收藏' : '☆ 收藏'}
            </button>
          )}
        </ResponsiveMenu>
      </div>
      {findOpen && (
        <div className="realm-find">
          <input
            ref={findInput}
            type="search"
            aria-label="文内精确查找"
            placeholder="在当前正文精确查找…"
            value={find}
            onChange={e => setFind(e.target.value)}
            onKeyDown={e => {
              if (e.nativeEvent.isComposing) return;
              if (e.key === 'Enter') {
                e.preventDefault();
                nextMatch(e.shiftKey ? -1 : 1);
              }
              if (e.key === 'Escape') {
                e.stopPropagation();
                setFindOpen(false);
              }
            }}
          />
          <span role="status">{matchCount ? `${matchIndex + 1} / ${matchCount}` : find ? '无精确匹配' : '输入关键词'}</span>
          <button disabled={!matchCount} aria-label="上一个匹配" onClick={() => nextMatch(-1)}>↑</button>
          <button disabled={!matchCount} aria-label="下一个匹配" onClick={() => nextMatch(1)}>↓</button>
          <button aria-label="关闭文内查找" onClick={() => setFindOpen(false)}>×</button>
        </div>
      )}
      <div ref={scroll} className="realm-reader-scroll">
        <div className={`realm-reader-layout${showDesktopOutline ? ' has-outline' : ''}`}>
          <div className="realm-reader-main">
            <header className="realm-document-heading">
              <div className="realm-document-meta">
                {detail.top_kind && <span className="doc-meta-kind">{detail.top_kind}</span>}
                {detail.side && <span className="doc-meta-side" data-side={sideKey(detail.side)}>{detail.side}</span>}
                {detail.category && detail.category !== detail.top_kind && <span className="doc-meta-cat">{detail.category}</span>}
                {detail.authors?.length ? <span className="doc-meta-authors">创作者：{detail.authors.join(' / ')}</span> : null}
              </div>
              <h1 ref={title} tabIndex={-1}>
                {seqMatch ? (
                  <>
                    <span className="doc-seq-prefix">{seqMatch[1]}</span>
                    <span>{seqMatch[2]}</span>
                  </>
                ) : (
                  detail.title
                )}
              </h1>
              <div className="realm-document-actions">
                <button onClick={() => copy(detail.content, '已复制完整原文')}>复制全文</button>
                <button onClick={downloadTxt}>下载 TXT</button>
                <button onClick={() => copy(resourceLink(detail.path), '已复制档案链接')}>分享档案</button>
                <details className="doc-info-disclosure">
                  <summary>档案信息</summary>
                  <div className="doc-info-popover">
                    <p>{detail.path}</p>
                    <p>{detail.encoding} · {detail.size} 字节</p>
                    <button onClick={() => copy(detail.path, '已复制文件路径')}>复制路径</button>
                  </div>
                </details>
              </div>
            </header>
            <div className="realm-reading-note">
              {mode === 'adaptive'
                ? '结构化阅读视图：自动提炼属性标签与能力卡，非标准段落按原序保留，实际判定以原文为准。'
                : '纯文本原文视图：100% 保留原始换行、缩进与符号顺序。'}
            </div>
            <div className={`realm-document-content${mode === 'raw' ? ' is-raw' : ''}`}>
              {sections.length ? (
                sections.map(section => (
                  <AdaptiveSection
                    key={section.id}
                    section={section}
                    docTitle={detail.title}
                    idPrefix={idPrefix}
                    rawMode={mode === 'raw'}
                    query={activeQuery}
                    onCopy={text => copy(text)}
                    onShare={id => copy(resourceLink(detail.path, id), '已复制条目定位链接')}
                  />
                ))
              ) : (
                <p>这份档案暂无正文。</p>
              )}
            </div>
            <footer className="realm-document-footer">
              <span>档案末尾 · 宏观界域强化序列库</span>
              <button
                onClick={() => {
                  if (scroll.current) scroll.current.scrollTop = 0;
                  onAnchor?.('');
                }}
              >
                回到顶部 ↑
              </button>
            </footer>
          </div>
          {showDesktopOutline && (
            <aside className="realm-doc-outline" aria-label="本文章节速览">
              <div className="outline-head">
                <strong>本文目录</strong>
                <span>{sections.length}</span>
              </div>
              <nav className="outline-list">
                {sections.map(section => {
                  const rank = section.kind === 'entry' ? extractRank(section.title) : undefined;
                  const label = section.kind === 'intro' ? '档案概览' : cleanOutlineTitle(section.title);
                  return (
                    <button
                      key={section.id}
                      type="button"
                      className={activeSectionId === section.id ? 'is-active' : undefined}
                      onClick={() => jump(section.id)}
                      title={section.title}
                    >
                      <span className="outline-item-text">{label}</span>
                      {rank && <em className={`outline-rank rank-${rank.tier}`}>{rank.badge}</em>}
                    </button>
                  );
                })}
              </nav>
            </aside>
          )}
        </div>
      </div>
      <div className={`realm-toast${notice ? ' visible' : ''}`} role="status" aria-live="polite">{notice}</div>
    </article>
  );
}
