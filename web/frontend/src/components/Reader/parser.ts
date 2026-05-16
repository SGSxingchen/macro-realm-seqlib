/** 阅读区结构化解析器：泛用、不折叠 section（避免误切）。
 *
 * 设计原则：
 * - 凡是包在 `【...】` 里的整行内容（小节标题、修订标签、更新条目），统一作为 'banner' 渲染——
 *   视觉上突出但不可折叠。这样无论是「【职业称号】」、「【加强伤害和烂钱】」、还是
 *   「【25.10.4：为所有技能明确词条】」都不会误切结构。
 * - 真正结构化只识别 `[能力名称]:`、`[字段]:值` 形式的能力卡，把它们渲染成属性表。
 * - 其余一律 paragraph 平铺。
 *
 * 等用户内容侧统一格式后，再决定是否启用 section 折叠。
 */

export type Block =
  | { kind: 'title'; text: string }
  | { kind: 'meta'; text: string }
  | { kind: 'banner'; text: string }                                   // 【...】 行
  | { kind: 'ability'; name: string; level?: string; tags: string[]; fields: Array<{ key: string; value: string }>; tail?: string }
  | { kind: 'paragraph'; text: string };

const BANNER_BARE_RE = /^【([^】]+)】\s*$/;
const BANNER_INLINE_RE = /^【([^】]+)】(.+)$/;
const ABILITY_NAME_RE = /^\[能力名称\]\s*[:：]\s*(.+)$/;
const ABILITY_FIELD_RE = /^\[([^\]]+)\]\s*[:：]\s*(.*)$/;
const META_LINE_RE = /^[（(]\s*(?:制作人|作者|原作者|审核人|修改人|调整人|重置人|复查人|策划)\s*[:：][^)）]*[)）]\s*$/;

const LEVEL_TAGS = ['EX', 'S', 'A', 'B', 'C', 'D', 'E', 'F'];
const LEVEL_RE = new RegExp(`[（(]\\s*(${LEVEL_TAGS.join('|')})级?\\s*[)）]`);

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

    const am = t.match(ABILITY_NAME_RE);
    if (am) {
      flushPara();
      const head = extractAbilityHeader(am[1].trim());
      const fields: Array<{ key: string; value: string }> = [];
      const tail: string[] = [];
      i++;
      while (i < lines.length) {
        const t2 = lines[i].trim();
        if (!t2) break;
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
        } else {
          tail.push(t2);
          i++;
        }
      }
      blocks.push({ kind: 'ability', ...head, fields, tail: tail.join('\n') || undefined });
      continue;
    }

    const sb = t.match(BANNER_BARE_RE);
    if (sb) {
      flushPara();
      blocks.push({ kind: 'banner', text: sb[1].trim() });
      i++;
      continue;
    }
    const si = t.match(BANNER_INLINE_RE);
    if (si) {
      flushPara();
      blocks.push({ kind: 'banner', text: `${si[1].trim()}：${si[2].trim()}` });
      i++;
      continue;
    }

    pendingPara.push(t);
    i++;
  }
  flushPara();

  return { title, blocks };
}
