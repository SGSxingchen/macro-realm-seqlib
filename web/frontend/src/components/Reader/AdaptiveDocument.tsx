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
const CREDIT_TOKEN_RE = /[（(【\[]\s*(制作人|制作者|投稿人|调整人|修改人|重置人|审核人|审核者|复查人|文本优化|作者|协力)\s*[:：]\s*([^）)】\]]+?)\s*[）)】\]]/g;
const CHANGELOG_BRACKET_RE = /^[【\[]\s*(\d{2,4}\s*[./年-]\s*\d{1,2}[^：:】\]]*)\s*[:：]\s*([^】\]]+?)\s*[】\]]$/;
const CHANGELOG_BARE_RE = /^(\d{2,4}\s*[./-]\s*\d{1,2}(?:\s*[./-]\s*\d{1,2})?)\s*[:：]\s*(.+)$/;
const PHASE_NOTE_HINT_RE = /(积分|徽章|消耗|兑换|解锁|前置|条件|习得|购买|视作|需要|仅当|限制|要求|开启|免费|升级)/;

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

function parseChangelogLine(trimmed: string): { date: string; note: string; rawLine: string } | null {
  const m1 = trimmed.match(CHANGELOG_BRACKET_RE);
  if (m1) return { date: m1[1].trim(), note: m1[2].trim(), rawLine: trimmed };
  const m2 = trimmed.match(CHANGELOG_BARE_RE);
  if (m2) return { date: m2[1].trim(), note: m2[2].trim(), rawLine: trimmed };
  return null;
}

type IntroParsed = {
  credits: Array<{ role: string; name: string; rawToken: string }>;
  changelogs: Array<{ date: string; note: string; rawLine: string }>;
  restFirstRaw: string;
  restParts: Part[];
};

export function parseIntroContent(parts: Part[], docTitle?: string, highlightQuery = ''): IntroParsed | null {
  if (!parts.length || parts[0].field) return null;
  const lines = parts[0].raw.match(/[^\r\n]*(?:\r\n|\n|\r|$)/g)?.filter(Boolean) || [];
  let idx = 0;
  const credits: Array<{ role: string; name: string; rawToken: string }> = [];
  const changelogs: Array<{ date: string; note: string; rawLine: string }> = [];

  while (idx < lines.length && !lines[idx].trim()) idx++;
  if (idx < lines.length && docTitle) {
    const cleanDocTitle = docTitle.trim();
    const bareDocTitle = cleanDocTitle.replace(/^\d{1,4}】\s*/, '').trim();
    const cleanFirst = lines[idx].replace(/^\uFEFF/, '').trim();
    if (cleanFirst === cleanDocTitle || (bareDocTitle && cleanFirst === bareDocTitle)) {
      idx++;
      while (idx < lines.length && !lines[idx].trim()) idx++;
      if (idx < lines.length && bareDocTitle && !highlightQuery.trim()) {
        const secondTrimmed = lines[idx].trim();
        const strippedSecond = secondTrimmed
          .replace(/[（(\[【]\s*(?:职业|技能表|特质|特质改造|能量池|公共建筑|战技侧|神秘侧|科技侧|特殊侧|新过审序列|[·/\s])+?\s*[）)\]】]/g, '')
          .trim();
        if (strippedSecond === bareDocTitle || strippedSecond === cleanDocTitle) {
          idx++;
          while (idx < lines.length && !lines[idx].trim()) idx++;
        }
      }
    }
  }

  while (idx < lines.length) {
    const trimmed = lines[idx].trim();
    if (!trimmed) {
      idx++;
      continue;
    }
    CREDIT_TOKEN_RE.lastIndex = 0;
    const strippedCredits = trimmed.replace(CREDIT_TOKEN_RE, '').trim();
    CREDIT_TOKEN_RE.lastIndex = 0;
    if (!strippedCredits && CREDIT_TOKEN_RE.test(trimmed)) {
      CREDIT_TOKEN_RE.lastIndex = 0;
      let m: RegExpExecArray | null;
      while ((m = CREDIT_TOKEN_RE.exec(trimmed)) !== null) {
        credits.push({ role: m[1], name: m[2].trim(), rawToken: m[0] });
      }
      idx++;
      continue;
    }
    CREDIT_TOKEN_RE.lastIndex = 0;

    const log = parseChangelogLine(trimmed);
    if (log) {
      changelogs.push(log);
      idx++;
      continue;
    }
    break;
  }

  const bodyLines = lines.slice(idx);
  if (parts.length === 1 && bodyLines.length > 0) {
    let endIdx = bodyLines.length - 1;
    while (endIdx >= 0 && !bodyLines[endIdx].trim()) endIdx--;
    const trailingLogs: Array<{ date: string; note: string; rawLine: string }> = [];
    while (endIdx >= 0) {
      const t = bodyLines[endIdx].trim();
      if (!t) {
        endIdx--;
        continue;
      }
      const log = parseChangelogLine(t);
      if (log) {
        trailingLogs.push(log);
        endIdx--;
        continue;
      }
      break;
    }
    if (trailingLogs.length > 0) {
      trailingLogs.reverse();
      changelogs.push(...trailingLogs);
      bodyLines.splice(endIdx + 1);
    }
  }

  const restFirstRaw = bodyLines.join('');
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

function buildEntryGroups(bodyParts: Part[]): RenderGroup[] {
  const statFields: Field[] = [];
  const seenKeys = new Set<string>();
  let hasDuplicateStatKey = false;
  const otherParts: Part[] = [];
  for (const part of bodyParts) {
    if (part.field && isCompactStat(part.field)) {
      const k = part.field.key.trim();
      if (seenKeys.has(k)) {
        hasDuplicateStatKey = true;
        break;
      }
      seenKeys.add(k);
      statFields.push(part.field);
    } else {
      otherParts.push(part);
    }
  }
  if (hasDuplicateStatKey || statFields.length === 0) {
    return groupParts(bodyParts);
  }
  return [{ type: 'stats', items: statFields }, ...groupParts(otherParts)];
}

function renderPartGroup(group: RenderGroup, idx: number, query: string) {
  if (group.type === 'stats') {
    return (
      <div className="archive-stat-grid" key={`stats-${idx}`}>
        {group.items.map((field, j) => {
          const val = field.value.trim();
          const isNil = NONE_VAL_RE.test(val);
          return (
            <div className={`archive-stat-chip${isNil ? ' is-nil is-muted' : ''}`} key={j}>
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
  return (
    <div className="archive-prose" key={idx}>
      <MarkedText text={part.raw.replace(/^(?:\r\n|\n|\r)+|(?:\r\n|\n|\r)+$/g, '')} query={query} />
    </div>
  );
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
        data-archive-anchor={section.id}
        data-section-id={section.id}
        className={`archive-section archive-${section.kind} archive-raw`}
        aria-label={section.title}
      >
        <pre className="archive-source"><MarkedText text={section.raw} query={query} /></pre>
      </section>
    );
  }

  // Intro section: always fold duplicate title & extract credits / changelogs with inline match highlighting
  if (section.kind === 'intro') {
    const parsed = parseIntroContent(section.parts, docTitle, query);
    if (parsed) {
      const hasProse = Boolean(parsed.restFirstRaw.trim()) || parsed.restParts.length > 0;
      const hasMeta = parsed.credits.length > 0 || parsed.changelogs.length > 0;
      if (!hasMeta && !hasProse) {
        return (
          <section
            id={idPrefix + section.id}
            data-archive-anchor={section.id}
            data-section-id={section.id}
            className="archive-section archive-intro is-empty"
            aria-label={section.title}
          />
        );
      }
      const remainingParts: Part[] = [
        ...(parsed.restFirstRaw.trim() ? [{ raw: parsed.restFirstRaw }] : []),
        ...parsed.restParts,
      ];
      const qLower = query.trim().toLocaleLowerCase();
      return (
        <section
          id={idPrefix + section.id}
          data-archive-anchor={section.id}
          data-section-id={section.id}
          className="archive-section archive-intro"
          aria-label={section.title}
        >
          {parsed.credits.length > 0 && (
            <div className="archive-intro-credits archive-credits">
              {parsed.credits.map((c, i) => {
                const tokenMatchOnly =
                  Boolean(qLower) &&
                  c.rawToken.toLocaleLowerCase().includes(qLower) &&
                  !c.role.toLocaleLowerCase().includes(qLower) &&
                  !c.name.toLocaleLowerCase().includes(qLower);
                return (
                  <span className="archive-credit-pill" key={i}>
                    {tokenMatchOnly ? (
                      <MarkedText text={c.rawToken} query={query} />
                    ) : (
                      <>
                        <em><MarkedText text={c.role} query={query} /></em>
                        <b><MarkedText text={c.name} query={query} /></b>
                      </>
                    )}
                  </span>
                );
              })}
            </div>
          )}
          {remainingParts.length > 0 && (
            <div className="archive-parts archive-intro-prose">
              {groupParts(remainingParts).map((g, i) => renderPartGroup(g, i, query))}
            </div>
          )}
          {parsed.changelogs.length > 0 && (
            <div className="archive-intro-logs archive-changelogs">
              {parsed.changelogs.map((log, i) => {
                const rawMatchOnly =
                  Boolean(qLower) &&
                  log.rawLine.toLocaleLowerCase().includes(qLower) &&
                  !log.date.toLocaleLowerCase().includes(qLower) &&
                  !log.note.toLocaleLowerCase().includes(qLower);
                return (
                  <div className="archive-changelog-item" key={i}>
                    {rawMatchOnly ? (
                      <span className="changelog-note"><MarkedText text={log.rawLine} query={query} /></span>
                    ) : (
                      <>
                        <span className="changelog-date"><MarkedText text={log.date} query={query} /></span>
                        <span className="changelog-note"><MarkedText text={log.note} query={query} /></span>
                      </>
                    )}
                  </div>
                );
              })}
            </div>
          )}
        </section>
      );
    }

    return (
      <section
        id={idPrefix + section.id}
        data-archive-anchor={section.id}
        data-section-id={section.id}
        className="archive-section archive-intro"
        aria-label={section.title}
      >
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
    const colonIdx = section.title.search(/[:：]/);
    const splitLog = colonIdx > 0 && !/[:：]/.test(query);
    return (
      <section
        id={idPrefix + section.id}
        data-archive-anchor={section.id}
        data-section-id={section.id}
        className="archive-section archive-section-changelog"
        aria-label={section.title}
      >
        <div className="archive-changelog-row">
          {splitLog ? (
            <>
              <span className="changelog-date"><MarkedText text={section.title.slice(0, colonIdx).trim()} query={query} /></span>
              <span className="changelog-text"><MarkedText text={section.title.slice(colonIdx + 1).trim()} query={query} /></span>
            </>
          ) : (
            <span className="changelog-text"><MarkedText text={section.title} query={query} /></span>
          )}
          <div className="archive-inline-actions">
            <button type="button" onClick={() => onCopy(section.raw)}>复制原文</button>
            <button type="button" onClick={() => onShare(section.id)} aria-label={`分享定位：${section.title}`}>定位链接</button>
          </div>
        </div>
      </section>
    );
  }

  // Ordinary phase / chapter section -> sleek horizontal beam (.archive-section.is-phase)
  if (section.kind === 'section') {
    let bodyParts = section.parts;
    const headingText = section.title;
    const firstPart = section.parts[0];
    if (firstPart && !firstPart.field) {
      const lines = firstPart.raw.match(/[^\r\n]*(?:\r\n|\n|\r|$)/g)?.filter(Boolean) || [];
      if (lines.length > 0) {
        const line0 = lines[0].trim();
        if (line0 === `【${section.title}】` || /^#{1,3}\s+/.test(line0)) {
          const restRaw = lines.slice(1).join('');
          bodyParts = [
            ...(restRaw.trim() ? [{ raw: restRaw }] : []),
            ...section.parts.slice(1),
          ];
        }
      }
    }

    let phaseNote = '';
    if (bodyParts.length === 1 && !bodyParts[0].field) {
      const candidate = bodyParts[0].raw.trim();
      if (
        candidate &&
        !/[\r\n]/.test(candidate) &&
        (candidate.length <= 42 || (candidate.length <= 58 && PHASE_NOTE_HINT_RE.test(candidate)))
      ) {
        phaseNote = candidate;
        bodyParts = [];
      }
    }

    return (
      <section
        id={idPrefix + section.id}
        data-archive-anchor={section.id}
        data-section-id={section.id}
        className="archive-section is-phase"
        aria-label={section.title}
      >
        <header className="archive-phase-head">
          <h2 className="archive-phase-title archive-section-heading">
            <MarkedText text={headingText} query={query} />
          </h2>
          {phaseNote && (
            <span className="archive-phase-note" style={{ marginLeft: 'auto' }}>
              <MarkedText text={phaseNote} query={query} />
            </span>
          )}
          <div className="archive-phase-actions archive-card-actions">
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

  // Entry section: ability / module card (.archive-entry)
  const firstPart = section.parts[0];
  const hasPromotedTitle =
    Boolean(firstPart?.field && /名称$/.test(firstPart.field.key) && firstPart.field.value.trim());

  const rank = extractRank(section.title);
  const rawStarterKind = hasPromotedTitle ? firstPart.field!.key.replace(/名称$/, '') : '';
  const showKindPill =
    hasPromotedTitle &&
    ((Boolean(query.trim()) && firstPart.field!.key.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase())) ||
      (Boolean(rawStarterKind) && rawStarterKind !== '能力' && rawStarterKind !== '技能'));
  const keyMatchesQuery =
    hasPromotedTitle && Boolean(query.trim()) &&
    firstPart.field!.key.toLocaleLowerCase().includes(query.trim().toLocaleLowerCase());

  const rawEntryTitle = hasPromotedTitle ? firstPart.field!.value.trim() : section.title;
  const strippedEntryTitle = rank
    ? rawEntryTitle.replace(/[（(\[【]\s*(?:EX|[SABCDEF][+-]?)\s*级?\s*[）)\]】]/i, '').trim()
    : rawEntryTitle;
  const qNorm = query.trim().toLocaleLowerCase();
  const displayEntryTitle =
    Boolean(strippedEntryTitle) &&
    (!qNorm ||
      strippedEntryTitle.toLocaleLowerCase().includes(qNorm) ||
      !rawEntryTitle.toLocaleLowerCase().includes(qNorm))
      ? strippedEntryTitle
      : rawEntryTitle;

  const bodyParts = hasPromotedTitle ? section.parts.slice(1) : section.parts;

  return (
    <section
      id={idPrefix + section.id}
      data-archive-anchor={section.id}
      data-section-id={section.id}
      data-rank={rank?.tier}
      className="archive-section archive-entry"
      aria-label={section.title}
    >
      <header className="archive-section-bar">
        <div className="archive-heading-group">
          {showKindPill && (
            <span className="archive-kind-pill">
              {keyMatchesQuery ? <MarkedText text={firstPart.field!.key} query={query} /> : rawStarterKind}
            </span>
          )}
          {hasPromotedTitle && (
            <h3 className="archive-entry-heading">
              <MarkedText text={displayEntryTitle} query={query} />
            </h3>
          )}
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
          {buildEntryGroups(bodyParts).map((g, i) => renderPartGroup(g, i, query))}
        </div>
      )}
      {showSource && <pre className="archive-source source-compare">{section.raw}</pre>}
    </section>
  );
}
