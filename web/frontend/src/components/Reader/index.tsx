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

const CANONICAL_SIDES = ['战技侧', '神秘侧', '科技侧', '特殊侧'] as const;

const SIDE_DESCRIPTIONS: Record<string, string> = {
  '战技侧': '冷热兵器技艺、格斗流派、极意体术与正面交锋序列',
  '神秘侧': '魔术礼装、神代血脉、咒术仪式、灵魂与异能体系',
  '科技侧': '义体殖装、动力装甲、高能军工、战术智能与力场造物',
  '特殊侧': '跨世界法则、概念类特质、替身/异常机制与特殊序列',
};

export function Reader(props: Props) {
  if (props.error) {
    return (
      <article className="realm-reader">
        {props.onBack && <button type="button" className="realm-back" onClick={props.onBack}>← 返回结果</button>}
        <div className="realm-reader-state" role="alert">
          <h2>暂时无法打开这份档案</h2>
          <p>资源可能已改名、下架，或网络暂不可用。不会自动替换成同名条目。</p>
          <button type="button" onClick={props.onRetry}>重新加载</button>
          <details><summary>错误详情</summary><pre>{props.error}</pre></details>
        </div>
      </article>
    );
  }
  if (props.loading) {
    return (
      <article className="realm-reader" aria-busy="true">
        {props.onBack && <button type="button" className="realm-back" onClick={props.onBack}>← 返回结果</button>}
        <div className="realm-reader-state">
          <p role="status">正在展开档案…</p>
          <div className="realm-skeleton"><div /><div /><div /></div>
        </div>
      </article>
    );
  }
  if (!props.detail) {
    const w = props.welcome;
    const sideCountMap = new Map((w?.sides || []).map(s => [s.name, s.count]));
    const quadrantSides = w
      ? CANONICAL_SIDES.map(name => ({ name, count: sideCountMap.get(name) ?? 0 }))
      : [];
    const sortedCategories = w
      ? [...w.categories].sort((a, b) => b.count - a.count)
      : [];

    return (
      <article className="realm-reader realm-welcome">
        <div className="realm-welcome-container">
          <header className="realm-welcome-hero">
            <div className="realm-welcome-kicker">
              {w?.totalCount ? `宏观界域 · 现存 ${w.totalCount} 项公开强化序列` : '宏观界域 · 公开强化序列总枢'}
            </div>
            <h1>轮回强化序列总枢</h1>
            <p>覆盖万界副本的职业、技能表、特质改造、能量池与界域公共设施。支持多词组合与拼音检索、模块化参数速读、原始文本逐字对照及多档案同屏比对。</p>
            <div className="realm-welcome-cta">
              <button
                type="button"
                className="realm-welcome-search"
                style={{
                  flex: '1 1 240px',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: '12px',
                  textAlign: 'left',
                }}
                onClick={props.onBrowse}
              >
                <span>检索序列名称、规则词条、出处世界或创作者…</span>
                <kbd aria-hidden="true">⌘K</kbd>
              </button>
              {w && w.sampleItems.length > 0 && (
                <button type="button" className="realm-secondary-cta realm-welcome-random" onClick={w.onOpenRandom}>
                  随机抽取序列
                </button>
              )}
            </div>
          </header>

          {w && (w.sides.length > 0 || w.totalCount > 0) && (
            <section className="realm-welcome-section" aria-label="四大强化体系">
              <h3>四大强化体系</h3>
              <div className="realm-quadrant-matrix">
                {quadrantSides.map(side => {
                  const sKey = sideKey(side.name);
                  return (
                    <button
                      key={side.name}
                      type="button"
                      className="realm-quad-cell"
                      onClick={() => w.onPickSide(side.name)}
                    >
                      <div
                        className="realm-quad-head"
                        style={{
                          display: 'flex',
                          justifyContent: 'space-between',
                          alignItems: 'baseline',
                          gap: '8px',
                          width: '100%',
                        }}
                      >
                        <span
                          className="realm-quad-name"
                          style={{ display: 'inline-flex', alignItems: 'center', gap: '8px' }}
                        >
                          <i
                            className="realm-quad-bar"
                            style={{
                              width: '3px',
                              height: '0.95em',
                              display: 'inline-block',
                              flexShrink: 0,
                              ...(sKey ? { background: `var(--side-${sKey})` } : {}),
                            }}
                            aria-hidden="true"
                          />
                          <strong>{side.name}</strong>
                        </span>
                        {side.count > 0 && <span className="realm-quad-count">{side.count} 项</span>}
                      </div>
                      <p className="realm-quad-desc">{SIDE_DESCRIPTIONS[side.name]}</p>
                    </button>
                  );
                })}
              </div>
            </section>
          )}

          {w && sortedCategories.length > 0 && (
            <section className="realm-welcome-section" aria-label="六大兑换分区">
              <h3>六大兑换分区</h3>
              <div className="realm-terminal-strip">
                {sortedCategories.map(cat => (
                  <button
                    key={cat.path}
                    type="button"
                    className="realm-term-cell"
                    onClick={() => w.onPickCategory(cat.path)}
                  >
                    <strong>{cat.count}</strong>
                    <span>{cat.name}</span>
                  </button>
                ))}
              </div>
            </section>
          )}

          {w && (w.savedItems.length > 0 || w.sampleItems.length > 0) && (() => {
            const quickItems: Array<{ path: string; title: string; side?: string }> =
              w.savedItems.length > 0 ? w.savedItems.slice(0, 6) : w.sampleItems.slice(0, 6);
            return (
              <section className="realm-welcome-section" aria-label={w.savedItems.length > 0 ? '随行收藏' : '序列速览'}>
                <h3>{w.savedItems.length > 0 ? '随行收藏直达' : '当前列表速览'}</h3>
                <div className="welcome-quick-list">
                  {quickItems.map(item => {
                    const qSeq = item.title.match(/^(\d{1,4})】\s*(.+)$/);
                    return (
                      <button
                        key={item.path}
                        type="button"
                        className="welcome-quick-item"
                        onClick={() => w.onOpenResource(item.path)}
                      >
                        <span className="quick-item-title">
                          {qSeq ? `${qSeq[1]} · ${qSeq[2]}` : item.title}
                        </span>
                        {item.side ? (
                          <small data-side={sideKey(item.side)}>{item.side}</small>
                        ) : (
                          <small>查阅 ↵</small>
                        )}
                      </button>
                    );
                  })}
                </div>
              </section>
            );
          })()}

          <footer className="realm-welcome-notes">
            <span>快捷键聚焦全局检索</span>
            <span>支持多词组合与拼音匹配</span>
            <span>原文逐字无损保留 · 不改写任何规则</span>
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

  const seqMatch = detail.title.match(/^(\d{1,4})(】)(.+)$/);
  const sKey = sideKey(detail.side);
  const kickerParts: Array<{ text: string; accent?: boolean; side?: string }> = [];
  if (seqMatch) kickerParts.push({ text: `编号 ${seqMatch[1]}`, accent: true });
  if (detail.top_kind) kickerParts.push({ text: detail.top_kind });
  if (detail.side && detail.side !== detail.top_kind) kickerParts.push({ text: detail.side, side: sKey });
  if (!kickerParts.length && detail.category) kickerParts.push({ text: detail.category });

  const showDesktopOutline = instance === 'primary' && mode === 'adaptive' && sections.length >= 2;

  return (
    <article className="realm-reader">
      <div className="realm-reader-toolbar">
        {onBack && <button type="button" className="realm-back" onClick={onBack}>← 结果</button>}
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
              <button type="button" key={section.id} className={activeSectionId === section.id ? 'is-active' : undefined} onClick={() => jump(section.id)}>
                {section.title}
              </button>
            ))}
          </nav>
        </details>
        <button
          type="button"
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
            <button type="button" aria-pressed={mode === 'adaptive'} onClick={() => setMode('adaptive')}>自适应</button>
            <button type="button" aria-pressed={mode === 'raw'} onClick={() => setMode('raw')}>原文</button>
          </div>
          {onSave && (
            <button
              type="button"
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
          <button type="button" disabled={!matchCount} aria-label="上一个匹配" onClick={() => nextMatch(-1)}>↑</button>
          <button type="button" disabled={!matchCount} aria-label="下一个匹配" onClick={() => nextMatch(1)}>↓</button>
          <button type="button" aria-label="关闭文内查找" onClick={() => setFindOpen(false)}>×</button>
        </div>
      )}
      <div ref={scroll} className="realm-reader-scroll">
        <div className={`realm-reader-layout${showDesktopOutline ? ' has-outline' : ''}`}>
          <div className="realm-reader-main">
            <header className="realm-document-heading">
              {kickerParts.length > 0 && (
                <div className="realm-seq-kicker">
                  {kickerParts.map((part, idx) => (
                    <span key={idx} className="kicker-segment">
                      {idx > 0 && <span className="kicker-sep" aria-hidden="true"> · </span>}
                      <span
                        className={part.accent ? 'kicker-code' : part.side ? 'kicker-side' : undefined}
                        style={part.side ? { color: `var(--side-${part.side})` } : undefined}
                      >
                        {part.text}
                      </span>
                    </span>
                  ))}
                </div>
              )}
              <h1 ref={title} tabIndex={-1}>
                {seqMatch ? (
                  <>
                    <span className="doc-seq-prefix">{seqMatch[1] + seqMatch[2]}</span>
                    <span>{seqMatch[3]}</span>
                  </>
                ) : (
                  detail.title
                )}
              </h1>
              <div
                className="realm-document-subrow"
                style={{
                  display: 'flex',
                  flexWrap: 'wrap',
                  alignItems: 'center',
                  justifyContent: 'space-between',
                  gap: '8px 16px',
                }}
              >
                <div className="realm-document-meta" style={{ marginBottom: 0 }}>
                  {detail.category && <span className="doc-meta-cat">分区：{detail.category}</span>}
                  {detail.authors?.length ? <span className="doc-meta-authors">署名：{detail.authors.join(' / ')}</span> : null}
                </div>
                <div className="realm-document-actions">
                  <button type="button" onClick={() => copy(detail.content, '已复制完整原文')}>复制全文</button>
                  <button type="button" onClick={downloadTxt}>下载 TXT</button>
                  <button type="button" onClick={() => copy(resourceLink(detail.path), '已复制档案链接')}>分享档案</button>
                  <details className="doc-info-disclosure">
                    <summary>档案信息</summary>
                    <div className="doc-info-popover">
                      <p>{detail.path}</p>
                      <p>{detail.encoding} · {detail.size} 字节</p>
                      <button type="button" onClick={() => copy(detail.path, '已复制文件路径')}>复制路径</button>
                    </div>
                  </details>
                </div>
              </div>
            </header>
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
              <span>
                {mode === 'adaptive'
                  ? '档案末尾 · 模块化速读视图，实际判定以原文为准'
                  : '档案末尾 · 纯文本原文视图，逐字保留原始换行与符号'}
              </span>
              <button
                type="button"
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
                  const isPhase = section.kind === 'section';
                  const label = section.kind === 'intro' ? '序列导言' : cleanOutlineTitle(section.title);
                  const cls = `${isPhase ? 'is-phase ' : ''}${activeSectionId === section.id ? 'is-active' : ''}`.trim();
                  return (
                    <button
                      key={section.id}
                      type="button"
                      className={cls || undefined}
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
