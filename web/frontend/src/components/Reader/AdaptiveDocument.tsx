import { useState } from 'react';
import type { Field, Part, Section } from './adaptive';

export function MarkedText({ text, query }: { text: string; query: string }) {
  if (!query.trim()) return <>{text}</>;
  const needle = query.toLocaleLowerCase();
  const haystack = text.toLocaleLowerCase();
  const nodes = [];
  let start = 0;
  let match = haystack.indexOf(needle);
  while (match !== -1) {
    nodes.push(text.slice(start, match));
    nodes.push(<mark data-reader-match key={match}>{text.slice(match, match + query.length)}</mark>);
    start = match + query.length;
    match = haystack.indexOf(needle, start);
  }
  nodes.push(text.slice(start));
  return <>{nodes}</>;
}

const RULE_KEY_RE = /效果|限制|要求|条件|副作用|规则/;
const STAT_KEY_RE = /^(释放类型|施展时间|施法时间|动作类型|打击类型|攻击类型|伤害类型|属性类型|段位\/等级|等级|段位|阶级|消耗|消耗能量|能力消耗|法力消耗|体力消耗|消耗规则|冷却|冷却时间|技能冷却|持续|持续时间|维持时间|射程|范围|作用范围|距离|目标|恢复方式|充能|次数|使用次数)$/;
const NONE_VAL_RE = /^(无|无。|暂无|-|—|——|None|N\/A)$/i;
const CREDIT_TOKEN_RE = /[（(]\s*(制作人|制作者|投稿人|调整人|修改人|重置人|审核人|审核者|复查人|文本优化|作者|协力)\s*[:：]\s*([^）)]+?)\s*[）)]/g;
const CHANGELOG_LINE_RE = /^[【\[]\s*(\d{2,4}\s*[./年-]\s*\d{1,2}[^：:】\]]*)\s*[:：]\s*([^】\]]+?)\s*[】\]]$/;

export type RankInfo = { badge: string; tier: 'ex' | 's' | 'a' | 'b' | 'c' | 'd' | 'e' | 'f' };

export function extractRank(title: string): RankInfo | undefined {
  const m = title.match(/[（(\[【]\s*(EX|[SABCDEF][+-]?)\s*级?\s*[）)\]】]/i);
  if (!m) return undefined;
  const raw = m[1].toUpperCase();
  const badge = raw === 'EX' ? 'EX' : `${raw}级`;
  const letter = raw === 'EX' ? 'ex' : (raw[0].toLowerCase() as RankInfo['tier']);
  return { badge, tier: letter };
}

function isCompactStat(field: Field): boolean {
  const key = field.key.trim();
  const val = field.value.trim();
  if (!val || val.includes('\n') || val.includes('\r') || val.length > 36) return false;
  if (/名称$/.test(key) || RULE_KEY_RE.test(key)) return false;
  return STAT_KEY_RE.test(key);
}

type IntroParsed = {
  credits: Array<{ role: string; name: string }>;
  changelogs: Array<{ date: string; note: string }>;
  restFirstRaw: string;
  restParts: Part[];
};

function parseIntroParts(parts: Part[], docTitle?: string): IntroParsed | null {
  if (!parts.length || parts[0].field) return null;
  const lines = parts[0].raw.match(/[^\r\n]*(?:\r\n|\n|\r|$)/g)?.filter(Boolean) || [];
  let idx = 0;
  const credits: Array<{ role: string; name: string }> = [];
  const changelogs: Array<{ date: string; note: string }> = [];

  while (idx < lines.length && !lines[idx].trim()) idx++;
  if (idx < lines.length && docTitle) {
    const cleanFirst = lines[idx].replace(/^\uFEFF/, '').trim();
    if (cleanFirst === docTitle.trim()) {
      idx++;
      while (idx < lines.length && !lines[idx].trim()) idx++;
    }
  }

  while (idx < lines.length) {
    const trimmed = lines[idx].trim();
    if (!trimmed) {
      idx++;
      continue;
    }
    const strippedCredits = trimmed.replace(CREDIT_TOKEN_RE, '').trim();
    if (!strippedCredits && CREDIT_TOKEN_RE.test(trimmed)) {
      CREDIT_TOKEN_RE.lastIndex = 0;
      let m: RegExpExecArray | null;
      while ((m = CREDIT_TOKEN_RE.exec(trimmed)) !== null) {
        credits.push({ role: m[1], name: m[2].trim() });
      }
      idx++;
      continue;
    }
    CREDIT_TOKEN_RE.lastIndex = 0;

    const logMatch = trimmed.match(CHANGELOG_LINE_RE);
    if (logMatch) {
      changelogs.push({ date: logMatch[1].trim(), note: logMatch[2].trim() });
      idx++;
      continue;
    }
    break;
  }

  const restFirstRaw = lines.slice(idx).join('');
  return { credits, changelogs, restFirstRaw, restParts: parts.slice(1) };
}

type RenderGroup =
  | { type: 'stats'; items: Field[] }
  | { type: 'part'; part: Part };

function groupParts(parts: Part[]): RenderGroup[] {
  const groups: RenderGroup[] = [];
  for (const part of parts) {
    if (part.field && isCompactStat(part.field)) {
      const prev = groups[groups.length - 1];
      if (prev && prev.type === 'stats') {
        prev.items.push(part.field);
      } else {
        groups.push({ type: 'stats', items: [part.field] });
      }
    } else {
      groups.push({ type: 'part', part });
    }
  }
  return groups;
}

function renderPartGroup(group: RenderGroup, idx: number, query: string) {
  if (group.type === 'stats') {
    return (
      <div className="archive-stat-grid" key={`stats-${idx}`}>
        {group.items.map((field, j) => {
          const val = field.value.trim();
          const muted = NONE_VAL_RE.test(val);
          return (
            <div className={`archive-stat-chip${muted ? ' is-muted' : ''}`} key={j}>
              <span className="archive-stat-label"><MarkedText text={field.key} query={query} /></span>
              <span className="archive-stat-value"><MarkedText text={val} query={query} /></span>
            </div>
          );
        })}
      </div>
    );
  }
  const { part } = group;
  if (part.field) {
    const isName = /名称$/.test(part.field.key);
    const isRule = RULE_KEY_RE.test(part.field.key);
    return (
      <div className={`archive-field${isName ? ' is-name' : ''}${isRule ? ' is-rule' : ''}`} key={idx}>
        <div className="archive-field-label"><MarkedText text={part.field.key} query={query} /></div>
        <div className="archive-field-value"><MarkedText text={part.field.value.trimEnd()} query={query} /></div>
      </div>
    );
  }
  if (!part.raw.trim()) return null;
  return <div className="archive-prose" key={idx}><MarkedText text={part.raw.replace(/^(?:\r\n|\n|\r)+|(?:\r\n|\n|\r)+$/g, '')} query={query} /></div>;
}

export function AdaptiveSection({
  section,
  query,
  rawMode,
  docTitle,
  onCopy,
  onShare,
  idPrefix = '',
}: {
  section: Section;
  query: string;
  rawMode: boolean;
  docTitle?: string;
  idPrefix?: string;
  onCopy: (raw: string) => void;
  onShare: (id: string) => void;
}) {
  const [showSource, setShowSource] = useState(false);

  if (rawMode) {
    return (
      <section
        id={idPrefix + section.id}
        data-section-id={section.id}
        className={`archive-section archive-${section.kind} archive-raw`}
        aria-label={section.title}
      >
        <pre className="archive-source"><MarkedText text={section.raw} query={query} /></pre>
      </section>
    );
  }

  // Intro section: clean duplicate title + creator & changelog strip when not actively searching text inside them
  if (section.kind === 'intro') {
    const hasQuery = Boolean(query.trim());
    const parsed = !hasQuery ? parseIntroParts(section.parts, docTitle) : null;
    if (parsed) {
      const hasProse = Boolean(parsed.restFirstRaw.trim()) || parsed.restParts.length > 0;
      const hasMeta = parsed.credits.length > 0 || parsed.changelogs.length > 0;
      if (!hasMeta && !hasProse) {
        return <section id={idPrefix + section.id} data-section-id={section.id} className="archive-section archive-intro is-empty" aria-label={section.title} />;
      }
      const remainingParts: Part[] = [
        ...(parsed.restFirstRaw.trim() ? [{ raw: parsed.restFirstRaw }] : []),
        ...parsed.restParts,
      ];
      return (
        <section id={idPrefix + section.id} data-section-id={section.id} className="archive-section archive-intro" aria-label={section.title}>
          {hasMeta && (
            <div className="archive-intro-meta">
              {parsed.credits.length > 0 && (
                <div className="archive-credits">
                  {parsed.credits.map((c, i) => (
                    <span className="archive-credit-pill" key={i}>
                      <em>{c.role}</em>
                      <b>{c.name}</b>
                    </span>
                  ))}
                </div>
              )}
              {parsed.changelogs.length > 0 && (
                <div className="archive-changelogs">
                  {parsed.changelogs.map((log, i) => (
                    <div className="archive-changelog-item" key={i}>
                      <span className="changelog-date">{log.date}</span>
                      <span className="changelog-note">{log.note}</span>
                    </div>
                  ))}
                </div>
              )}
            </div>
          )}
          {remainingParts.length > 0 && (
            <div className="archive-parts archive-intro-prose">
              {groupParts(remainingParts).map((g, i) => renderPartGroup(g, i, query))}
            </div>
          )}
        </section>
      );
    }

    return (
      <section id={idPrefix + section.id} data-section-id={section.id} className="archive-section archive-intro" aria-label={section.title}>
        <div className="archive-parts archive-intro-prose">
          {groupParts(section.parts).map((g, i) => renderPartGroup(g, i, query))}
        </div>
      </section>
    );
  }

  // Standalone 1-line changelog section like 【25.10.5：为所有技能明确词条】
  const isSoloChangelog =
    section.kind === 'section' &&
    /^\d{2,4}\s*[./年-]\s*\d{1,2}/.test(section.title) &&
    section.parts.length === 1 &&
    !section.parts[0].field &&
    section.parts[0].raw.trim() === `【${section.title}】`;

  if (isSoloChangelog) {
    return (
      <section
        id={idPrefix + section.id}
        data-section-id={section.id}
        className="archive-section archive-section-changelog"
        aria-label={section.title}
      >
        <div className="archive-changelog-row">
          <span className="changelog-tag">修订记录</span>
          <span className="changelog-text"><MarkedText text={section.title} query={query} /></span>
          <div className="archive-inline-actions">
            <button type="button" onClick={() => onCopy(section.raw)}>复制原文</button>
            <button type="button" onClick={() => onShare(section.id)} aria-label={`分享定位：${section.title}`}>定位链接</button>
          </div>
        </div>
      </section>
    );
  }

  // Entry section: promote starter [*名称] into card header and group short stat chips
  const firstPart = section.parts[0];
  const hasPromotedTitle =
    section.kind === 'entry' &&
    Boolean(firstPart?.field && /名称$/.test(firstPart.field.key) && firstPart.field.value.trim());

  const rank = section.kind === 'entry' ? extractRank(section.title) : undefined;
  const entryKindLabel = hasPromotedTitle
    ? firstPart.field!.key.replace(/名称$/, '') || '条目'
    : section.kind === 'entry'
      ? '能力 / 条目'
      : '章节';

  // For 'section' kind, strip redundant first heading line (e.g. 【非标准章节】 or # 标题) if promoted into header
  let bodyParts = hasPromotedTitle ? section.parts.slice(1) : section.parts;
  let promotedSectionHeading = '';
  if (section.kind === 'section' && firstPart && !firstPart.field) {
    const lines = firstPart.raw.match(/[^\r\n]*(?:\r\n|\n|\r|$)/g)?.filter(Boolean) || [];
    if (lines.length > 0) {
      const line0 = lines[0].trim();
      if (line0 === `【${section.title}】` || /^#{1,3}\s+/.test(line0)) {
        promotedSectionHeading = section.title;
        const restRaw = lines.slice(1).join('');
        bodyParts = [
          ...(restRaw.trim() ? [{ raw: restRaw }] : []),
          ...section.parts.slice(1),
        ];
      }
    }
  }

  const keyMatchesQuery =
    hasPromotedTitle && Boolean(query.trim()) &&
    firstPart.field!.key.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase());

  return (
    <section
      id={idPrefix + section.id}
      data-section-id={section.id}
      data-rank={rank?.tier}
      className={`archive-section archive-${section.kind}`}
      aria-label={section.title}
    >
      <header className="archive-section-bar">
        <div className="archive-heading-group">
          <span className="archive-kind-pill">
            {keyMatchesQuery ? <MarkedText text={firstPart.field!.key} query={query} /> : entryKindLabel}
          </span>
          {hasPromotedTitle ? (
            <h3 className="archive-entry-heading">
              <MarkedText text={firstPart.field!.value.trim()} query={query} />
            </h3>
          ) : promotedSectionHeading ? (
            <h2 className="archive-section-heading">
              <MarkedText text={promotedSectionHeading} query={query} />
            </h2>
          ) : null}
          {rank && <span className={`archive-rank-badge rank-${rank.tier}`}>{rank.badge}</span>}
        </div>
        <div className="archive-card-actions">
          <button type="button" onClick={() => onCopy(section.raw)}>复制原文</button>
          <button type="button" onClick={() => onShare(section.id)} aria-label={`分享定位：${section.title}`}>定位链接</button>
          <button type="button" aria-expanded={showSource} onClick={() => setShowSource(v => !v)}>
            {showSource ? '关闭对照' : '原文对照'}
          </button>
        </div>
      </header>
      {bodyParts.length > 0 && (
        <div className="archive-parts">
          {groupParts(bodyParts).map((g, i) => renderPartGroup(g, i, query))}
        </div>
      )}
      {showSource && <pre className="archive-source source-compare">{section.raw}</pre>}
    </section>
  );
}
