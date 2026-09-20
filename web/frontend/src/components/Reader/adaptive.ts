/** Lossless presentation model. Never edits the resource or guesses missing rules. */
export type Field = { key: string; prefix: string; value: string };
export type Part = { raw: string; field?: Field };
export type Section = {
  id: string;
  title: string;
  kind: 'intro' | 'entry' | 'section';
  start: number;
  end: number;
  raw: string;
  parts: Part[];
};

const STARTERS = new Set([
  '能力名称', '技能名称', '称号名称', '道具名称', '能量池名称', '技艺名称',
  '模块名称', '建筑名称', '公共建筑名称', '奖励名称', '魔药名称', '序列名称',
]);
const LEGACY_FIELDS = new Set([
  ...STARTERS, '能力简介', '能力效果', '技能效果', '能力形容', '能力消耗',
  '释放类型', '打击类型', '伤害类型', '段位/等级', '消耗', '消耗能量',
  '消耗规则', '冷却', '冷却时间', '技能冷却', '持续', '持续时间',
  '恢复方式', '使用要求', '开放条件', '解锁条件', '基础限制', '补充说明',
  '风险/副作用', '技艺规则', '效果', '称号效果', '道具效果', '能量池效果',
]);

// An explicit bracketed field may use an unknown key. Bare prose is only a
// field when its key is known; e.g. “注意：” remains untouched prose.
export function readField(line: string): Field | undefined {
  const explicit = line.match(/^(\s*(?:\[([^\]\r\n]{1,32})\]|【([^】\r\n]{1,32})】)\s*[:：][ \t]*)([^\r\n]*)(\r\n|\n|\r)?$/);
  if (explicit) return { key: (explicit[2] || explicit[3]).trim(), prefix: explicit[1], value: explicit[4] + (explicit[5] || '') };
  const bare = line.match(/^(\s*([^\s:：\[\]【】]{1,16})\s*[:：][ \t]*)([^\r\n]*)(\r\n|\n|\r)?$/);
  if (bare && LEGACY_FIELDS.has(bare[2])) return { key: bare[2], prefix: bare[1], value: bare[3] + (bare[4] || '') };
  // Old sheets sometimes put a known field on a line without a colon.
  const solo = line.match(/^(\s*(?:\[([^\]]+)\]|【([^】]+)】)[ \t]*)(\r\n|\n|\r)?$/);
  if (solo && LEGACY_FIELDS.has((solo[2] || solo[3]).trim())) {
    return { key: (solo[2] || solo[3]).trim(), prefix: solo[1], value: solo[4] || '' };
  }
  return undefined;
}

function linesOf(text: string): string[] {
  return text.match(/[^\r\n]*(?:\r\n|\n|\r|$)/g)?.filter(Boolean) || [];
}

export function splitParts(raw: string): Part[] {
  const parts: Part[] = [];
  let afterBlank = false;
  for (const line of linesOf(raw)) {
    const field = readField(line);
    if (field) parts.push({ raw: line, field });
    else if (parts.length) {
      const last = parts[parts.length - 1];
      const completeName = last.field && STARTERS.has(last.field.key) && last.field.value.trim();
      // Blank-separated prose and text after a name are not silently assigned
      // to the previous field. They remain visible, in their original position.
      if (line.trim() && last.field && (afterBlank || completeName)) parts.push({ raw: line });
      else { last.raw += line; if (last.field) last.field.value += line; }
    } else parts.push({ raw: line });
    afterBlank = !line.trim();
  }
  return parts;
}

function heading(line: string): { title: string; kind: Section['kind'] } | undefined {
  const field = readField(line);
  if (field) {
    if (STARTERS.has(field.key) && field.value.trim()) return { title: field.value.trim(), kind: 'entry' };
    return undefined;
  }
  const text = line.trim();
  const markdown = text.match(/^#{1,3}\s+(.+)$/);
  if (markdown) return { title: markdown[1], kind: 'section' };
  const bracket = text.match(/^【([^】]{1,100})】$/);
  if (bracket) return { title: bracket[1], kind: 'section' };
  return undefined;
}

function hash(text: string): string {
  let value = 2166136261;
  for (let i = 0; i < text.length; i++) value = Math.imul(value ^ text.charCodeAt(i), 16777619);
  return (value >>> 0).toString(36);
}

export function buildDocument(content: string): Section[] {
  if (!content) return [];
  const starts: Array<{ start: number; title: string; kind: Section['kind'] }> = [];
  let offset = 0;
  const lines = linesOf(content);
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];
    let h = heading(line);
    const field = readField(line);
    const next = lines[i + 1];
    if (!h && field && STARTERS.has(field.key) && !field.value.trim() && next?.trim() && !readField(next) && !heading(next)) {
      h = { title: next.trim(), kind: 'entry' };
    }
    if (h) starts.push({ start: offset, ...h });
    offset += line.length;
  }
  if (!starts.length || starts[0].start !== 0) starts.unshift({ start: 0, title: '档案正文', kind: 'intro' });
  const occurrences = new Map<string, number>();
  return starts.map((s, i) => {
    const end = starts[i + 1]?.start ?? content.length;
    const raw = content.slice(s.start, end);
    const key = hash(`${s.kind}:${s.title}`);
    const occurrence = (occurrences.get(key) || 0) + 1;
    occurrences.set(key, occurrence);
    return { ...s, end, raw, id: `entry-${key}-${occurrence}`, parts: splitParts(raw) };
  });
}

export function resourceLink(path: string, anchor = ''): string {
  const url = new URL(window.location.href);
  url.searchParams.delete('tab');
  url.searchParams.set('open', path);
  url.hash = anchor;
  return url.href;
}
