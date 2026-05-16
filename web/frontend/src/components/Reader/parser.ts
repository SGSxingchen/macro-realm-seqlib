/** 阅读区结构化解析器：把 TXT 拆成 Section / Ability / Paragraph 三类节点。
 * 设计目标是 fail-safe — 任何识别失败都退回为 paragraph，永远能渲染原文。
 */

export type Block =
  | { kind: 'title'; text: string }
  | { kind: 'meta'; text: string }       // (制作人:xxx) (审核人：xxx)
  | { kind: 'section'; text: string; raw: string } // 【职业称号】
  | { kind: 'ability'; name: string; level?: string; tags: string[]; fields: Array<{ key: string; value: string }>; tail?: string }
  | { kind: 'paragraph'; text: string }
  | { kind: 'note'; text: string };       // 编者注、更新日期等

const SECTION_RE = /^【([^】]+)】\s*$/;
const SECTION_INLINE_RE = /^【([^】]+)】(.+)$/;
const ABILITY_NAME_RE = /^\[能力名称\]\s*[:：]\s*(.+)$/;
const ABILITY_FIELD_RE = /^\[([^\]]+)\]\s*[:：]\s*(.*)$/;
const META_LINE_RE = /^[（(]\s*(?:制作人|作者|原作者|审核人|修改人|调整人|重置人|复查人|策划)\s*[:：][^)）]*[)）]\s*$/;
const NOTE_LINE_RE = /^(?:【)?(\d{2,4}[.·]\d{1,2}[.·]\d{1,2}|\d{4}\/\d{1,2}\/\d{1,2})/;

const LEVEL_TAGS = ['EX', 'S', 'A', 'B', 'C', 'D', 'E', 'F'];
const LEVEL_RE = new RegExp(`[（(]\\s*(${LEVEL_TAGS.join('|')})级?\\s*[)）]`);

function extractAbilityHeader(name: string): { name: string; level?: string; tags: string[] } {
  // [能力名称]：罗马！（B级）【异常状态】
  let lvl: string | undefined;
  const tags: string[] = [];
  let cleaned = name;
  const lvlMatch = cleaned.match(LEVEL_RE);
  if (lvlMatch) {
    lvl = lvlMatch[1];
    cleaned = cleaned.replace(LEVEL_RE, '').trim();
  }
  cleaned = cleaned.replace(/【([^】]+)】/g, (_m, t) => {
    tags.push(t.trim());
    return '';
  }).trim();
  return { name: cleaned, level: lvl, tags };
}

export function parseDocument(content: string): { title: string; blocks: Block[] } {
  if (!content) return { title: '', blocks: [] };
  const lines = content.split(/\r?\n/);
  const blocks: Block[] = [];

  let i = 0;
  // 标题行
  while (i < lines.length && !lines[i].trim()) i++;
  const title = (lines[i] || '').trim();
  if (title) {
    blocks.push({ kind: 'title', text: title });
    i++;
  }

  // 元数据行（连续的 (制作人/审核人/...) ）
  while (i < lines.length) {
    const ln = lines[i].trim();
    if (!ln) { i++; continue; }
    if (META_LINE_RE.test(ln)) {
      blocks.push({ kind: 'meta', text: ln });
      i++;
    } else {
      break;
    }
  }

  let pendingPara: string[] = [];

  const flushPara = () => {
    if (pendingPara.length) {
      const text = pendingPara.join('\n').trim();
      if (text) {
        if (NOTE_LINE_RE.test(text)) blocks.push({ kind: 'note', text });
        else blocks.push({ kind: 'paragraph', text });
      }
      pendingPara = [];
    }
  };

  while (i < lines.length) {
    const ln = lines[i];
    const t = ln.trim();
    if (!t) { flushPara(); i++; continue; }

    const inline = t.match(SECTION_INLINE_RE);
    const sec = t.match(SECTION_RE) || (inline && inline[2].trim() === '' ? [t, inline[1]] as RegExpMatchArray : null);
    if (sec) {
      flushPara();
      blocks.push({ kind: 'section', text: sec[1].trim(), raw: t });
      i++;
      continue;
    }
    if (inline) {
      // 【小节】尾部还有内容（如「【终极能力】消耗9000+B兑换」），section + 后续段落
      flushPara();
      blocks.push({ kind: 'section', text: inline[1].trim(), raw: t });
      pendingPara.push(inline[2].trim());
      i++;
      continue;
    }

    const am = t.match(ABILITY_NAME_RE);
    if (am) {
      flushPara();
      const head = extractAbilityHeader(am[1].trim());
      const fields: Array<{ key: string; value: string }> = [];
      i++;
      const tail: string[] = [];
      while (i < lines.length) {
        const ln2 = lines[i];
        const t2 = ln2.trim();
        if (!t2) {
          // 空行：能力字段段落结束
          break;
        }
        if (ABILITY_NAME_RE.test(t2) || SECTION_RE.test(t2) || SECTION_INLINE_RE.test(t2)) break;
        const fm = t2.match(ABILITY_FIELD_RE);
        if (fm) {
          fields.push({ key: fm[1].trim(), value: fm[2].trim() });
          i++;
          // 可能跨多行：下一行不是 [字段]: 也不是空行也不是新能力/小节，则视作上一字段续行
          while (i < lines.length) {
            const next = lines[i];
            const nt = next.trim();
            if (!nt) break;
            if (ABILITY_NAME_RE.test(nt) || SECTION_RE.test(nt) || SECTION_INLINE_RE.test(nt) || ABILITY_FIELD_RE.test(nt)) break;
            const last = fields[fields.length - 1];
            last.value = (last.value ? last.value + '\n' : '') + nt;
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

    pendingPara.push(t);
    i++;
  }
  flushPara();

  return { title, blocks };
}
