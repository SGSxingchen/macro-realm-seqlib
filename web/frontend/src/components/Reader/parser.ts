/** 阅读区结构化解析器：泛用、不折叠 section（避免误切）。
 *
 * 设计原则：
 * - 凡是包在 `【...】` 里的整行内容（小节标题、修订标签、更新条目），默认作为 'banner' 渲染。
 * - 紧贴在单个能力卡前、描述最近修改的 `【...】` 行渲染成 ability-note，不当作首部审核记录。
 * - 明确属于规则说明的小节渲染成 rule-section，称号/道具/技艺等专用字段渲染成 special-card。
 * - 真正结构化只识别 `[能力名称]:`、`[字段]:值` 形式的能力卡，把它们渲染成属性表。
 * - 其余一律 paragraph 平铺。
 *
 * 等用户内容侧统一格式后，再决定是否启用 section 折叠。
 */

export type Block =
  | { kind: 'title'; text: string }
  | { kind: 'meta'; text: string }
  | { kind: 'banner'; text: string }                                   // 【...】 行
  | { kind: 'ability-note'; text: string }                              // 能力前的最近修改备注
  | { kind: 'level-tree'; items: Array<{ level: string; text: string }> }
  | { kind: 'rule-section'; title: string; lines: string[] }
  | { kind: 'ability'; name: string; level?: string; tags: string[]; fields: Array<{ key: string; value: string }>; tail?: string }
  | { kind: 'special-card'; cardType: string; name: string; fields: Array<{ key: string; value: string }> }
  | { kind: 'paragraph'; text: string };

const BANNER_BARE_RE = /^【([^】]+)】\s*$/;
const BANNER_INLINE_RE = /^【([^】]+)】(.+)$/;
const ABILITY_NAME_RE = /^\[能力名称\]\s*[:：]\s*(.+)$/;
const ABILITY_FIELD_RE = /^\[([^\]]+)\]\s*[:：]\s*(.*)$/;
const META_LINE_RE = /^[（(]\s*(?:制作人|制作者|作者|原作者|投稿人|审核|审核人|审核者|修改人|修改内容|调整人|重置人|重置内容|复查人|策划|文本优化)\s*[:：].*[)）]\s*$/;
const LEVEL_TREE_RE = /^(?:【\s*)?(EX|S|A|B|C|D|E|F)(?:\s*】)?\s*(?:级)?\s*[:：]\s*(.+)$/i;

const CARD_STARTERS: Record<string, string> = {
  '称号名称': '称号',
  '道具名称': '道具',
  '能量池名称': '能量池',
  '技艺名称': '技艺',
  '模块名称': '模块',
  '建筑名称': '公共建筑',
  '公共建筑名称': '公共建筑',
  '奖励名称': '奖励',
  '魔药名称': '魔药',
  '序列名称': '序列',
};
const RULE_SECTION_TITLES = new Set([
  '通用规则',
  '通用说明',
  '购买规则',
  '升级规则',
  '兑换规则',
  '学习规则',
  '获取规则',
  '使用规则',
  '维护规则',
  '建造规则',
  '规则说明',
  '总规则',
  '技能表规则',
  '技能表通用规则',
  '技艺规则',
  '道具规则',
  '能量池规则',
  '建筑规则',
  '公共建筑规则',
  '称号规则',
  '职业规则',
  '职业通用规则',
  '基础规则',
  '基础性能',
  '资源规则',
  '历史记录',
  '历史更新记录',
]);

const LEVEL_TAGS = ['EX', 'S', 'A', 'B', 'C', 'D', 'E', 'F'];
const LEVEL_RE = new RegExp(`[（(]\\s*(${LEVEL_TAGS.join('|')})级?\\s*[)）]`);
const ABILITY_NOTE_RE = /(\d{1,4}[.\/年]\d{1,2}|加强|削弱|调整|重写|新增|删除|修复|明确|补充|优化|数值|效果|冷却|描述|限制|降低|提升|下调|上调|去除|改为|增加)/;

function nextMeaningfulLine(lines: string[], start: number): string {
  let i = start;
  while (i < lines.length) {
    const t = lines[i].trim();
    if (t) return t;
    i++;
  }
  return '';
}

function isAbilityNote(title: string, lines: string[], nextIndex: number): boolean {
  if (!ABILITY_NOTE_RE.test(title)) return false;
  return ABILITY_NAME_RE.test(nextMeaningfulLine(lines, nextIndex));
}

function parseLevelTreeLine(line: string): { level: string; text: string } | undefined {
  const m = line.match(LEVEL_TREE_RE);
  if (!m) return undefined;
  const text = m[2].trim();
  if (!text) return undefined;
  return { level: m[1].toUpperCase(), text };
}

function readLevelTree(lines: string[], start: number): { items: Array<{ level: string; text: string }>; next: number } | undefined {
  const items: Array<{ level: string; text: string }> = [];
  let i = start;
  while (i < lines.length) {
    const current = lines[i].trim();
    if (!current) {
      const next = nextMeaningfulLine(lines, i + 1);
      if (parseLevelTreeLine(next)) {
        i++;
        continue;
      }
      break;
    }
    const item = parseLevelTreeLine(current);
    if (!item) break;
    items.push(item);
    i++;
  }
  if (items.length < 3) return undefined;
  return { items, next: i };
}

function extractAbilityHeader(name: string): { name: string; level?: string; tags: string[] } {
  let lvl: string | undefined;
  const tags: string[] = [];
  let cleaned = name;
  const m = cleaned.match(LEVEL_RE);
  if (m) {
    lvl = m[1];
    cleaned = cleaned.replace(LEVEL_RE, '').trim();
  }
  cleaned = cleaned.replace(/【([^】]+)】/g, (_m, t) => { tags.push(t.trim()); return ''; }).trim();
  return { name: cleaned, level: lvl, tags };
}

function readFieldCard(lines: string[], start: number, firstKey: string, firstValue: string) {
  const fields: Array<{ key: string; value: string }> = [{ key: firstKey, value: firstValue.trim() }];
  let i = start + 1;
  while (i < lines.length) {
    const t = lines[i].trim();
    if (!t) break;
    if (ABILITY_NAME_RE.test(t) || BANNER_BARE_RE.test(t) || BANNER_INLINE_RE.test(t)) break;
    const fm = t.match(ABILITY_FIELD_RE);
    if (fm) {
      const key = fm[1].trim();
      if (key === firstKey) break;
      fields.push({ key: fm[1].trim(), value: fm[2].trim() });
      i++;
      continue;
    }
    const last = fields[fields.length - 1];
    last.value = (last.value ? last.value + '\n' : '') + t;
    i++;
  }
  return { fields, next: i };
}

function filledFields(fields: Array<{ key: string; value: string }>): Array<{ key: string; value: string }> {
  return fields.filter((field) => field.value.trim());
}

export function parseDocument(content: string): { title: string; blocks: Block[] } {
  if (!content) return { title: '', blocks: [] };
  const lines = content.split(/\r?\n/);
  const blocks: Block[] = [];

  let i = 0;
  while (i < lines.length && !lines[i].trim()) i++;
  const title = (lines[i] || '').trim();
  if (title) {
    blocks.push({ kind: 'title', text: title });
    i++;
  }

  let pendingPara: string[] = [];

  const flushPara = () => {
    if (!pendingPara.length) return;
    const text = pendingPara.join('\n').trim();
    if (text) blocks.push({ kind: 'paragraph', text });
    pendingPara = [];
  };

  while (i < lines.length) {
    const t = lines[i].trim();
    if (!t) { flushPara(); i++; continue; }

    if (META_LINE_RE.test(t)) {
      flushPara();
      blocks.push({ kind: 'meta', text: t });
      i++;
      continue;
    }

    const levelTree = readLevelTree(lines, i);
    if (levelTree) {
      flushPara();
      blocks.push({ kind: 'level-tree', items: levelTree.items });
      i = levelTree.next;
      continue;
    }

    const am = t.match(ABILITY_NAME_RE);
    if (am) {
      flushPara();
      const head = extractAbilityHeader(am[1].trim());
      const fields: Array<{ key: string; value: string }> = [];
      const tail: string[] = [];
      i++;
      while (i < lines.length) {
        const t2 = lines[i].trim();
        if (!t2) {
          const next = nextMeaningfulLine(lines, i + 1);
          if (next && ABILITY_FIELD_RE.test(next) && !ABILITY_NAME_RE.test(next)) {
            i++;
            continue;
          }
          if (next && parseLevelTreeLine(next) && fields.length) {
            i++;
            continue;
          }
          break;
        }
        if (ABILITY_NAME_RE.test(t2)) break;
        if (BANNER_BARE_RE.test(t2)) break;
        const fm = t2.match(ABILITY_FIELD_RE);
        if (fm) {
          fields.push({ key: fm[1].trim(), value: fm[2].trim() });
          i++;
          // 续行
          while (i < lines.length) {
            const next = lines[i].trim();
            if (!next) break;
            if (
              ABILITY_NAME_RE.test(next) ||
              ABILITY_FIELD_RE.test(next) ||
              BANNER_BARE_RE.test(next) ||
              BANNER_INLINE_RE.test(next)
            ) break;
            const last = fields[fields.length - 1];
            last.value = (last.value ? last.value + '\n' : '') + next;
            i++;
          }
        } else if (parseLevelTreeLine(t2) && fields.length) {
          const last = fields[fields.length - 1];
          last.value = (last.value ? last.value + '\n' : '') + t2;
          i++;
        } else {
          tail.push(t2);
          i++;
        }
      }
      blocks.push({ kind: 'ability', ...head, fields: filledFields(fields), tail: tail.join('\n') || undefined });
      continue;
    }

    const cardMatch = t.match(ABILITY_FIELD_RE);
    if (cardMatch) {
      const firstKey = cardMatch[1].trim();
      const cardType = CARD_STARTERS[firstKey];
      if (cardType) {
        flushPara();
        const { fields, next } = readFieldCard(lines, i, firstKey, cardMatch[2]);
        const name = fields[0]?.value || cardType;
        blocks.push({ kind: 'special-card', cardType, name, fields: filledFields(fields.slice(1)) });
        i = next;
        continue;
      }
    }

    const sb = t.match(BANNER_BARE_RE);
    if (sb) {
      flushPara();
      const titleText = sb[1].trim();
      if (RULE_SECTION_TITLES.has(titleText)) {
        const linesOut: string[] = [];
        i++;
        while (i < lines.length) {
          const next = lines[i].trim();
          if (!next) break;
          if (ABILITY_NAME_RE.test(next) || BANNER_BARE_RE.test(next) || BANNER_INLINE_RE.test(next)) break;
          linesOut.push(next);
          i++;
        }
        blocks.push({ kind: 'rule-section', title: titleText, lines: linesOut });
        continue;
      }
      if (isAbilityNote(titleText, lines, i + 1)) {
        blocks.push({ kind: 'ability-note', text: titleText });
        i++;
        continue;
      }
      blocks.push({ kind: 'banner', text: titleText });
      i++;
      continue;
    }
    const si = t.match(BANNER_INLINE_RE);
    if (si) {
      flushPara();
      const inlineText = si[2].trim().replace(/^[:：]\s*/, '');
      blocks.push({ kind: 'banner', text: `${si[1].trim()}：${inlineText}` });
      i++;
      continue;
    }

    pendingPara.push(t);
    i++;
  }
  flushPara();

  return { title, blocks };
}
