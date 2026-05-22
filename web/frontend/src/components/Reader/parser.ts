/** 阅读区结构化解析器：泛用、不折叠 section（避免误切）。
 *
 * 设计原则：
 * - 凡是包在 `【...】` 里的整行内容（小节标题、修订标签、更新条目），默认作为 'banner' 渲染。
 * - 紧贴在单个能力卡前、描述最近修改的 `【...】` 行渲染成 ability-note，不当作首部审核记录。
 * - 明确属于规则说明的小节渲染成 rule-section，称号/道具/技艺等专用字段渲染成 special-card。
 * - 真正结构化优先识别标准 `[字段名]:`，并兼容历史资源里常见的 `【字段名】:`
 *   与少量裸字段名写法，把它们渲染成属性表。
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
const META_LINE_RE = /^[（(]\s*(?:制作人|制作者|作者|原作者|投稿人|审核|审核人|审核者|修改人|修改内容|调整人|重置人|重置内容|复查人|策划|文本优化)\s*[:：].*[)）]\s*$/;
const LEVEL_TREE_PREFIX_RE = /^(?:【\s*特质等级\s*】\s*[:：]\s*)?/;
const LEVEL_TREE_UNLOCK_WORDS = '解锁|升级|获得|开启|提升|开放';

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
const STRUCTURED_FIELD_KEYS = new Set([
  ...Object.keys(CARD_STARTERS),
  '能力名称',
  '能力简介',
  '能力效果',
  '能力形容',
  '能力消耗',
  '释放类型',
  '打击类型',
  '伤害类型',
  '段位/等级',
  '消耗',
  '消耗能量',
  '消耗规则',
  '冷却',
  '冷却时间',
  '技能冷却',
  '持续',
  '持续时间',
  '恢复方式',
  '使用要求',
  '开放条件',
  '解锁条件',
  '兑换/消耗条件',
  '获取资格',
  '所需仪式',
  '所需材料',
  '晋升条件',
  '基础限制',
  '补充说明',
  '冥想规则',
  '技艺栏限制',
  '风险/副作用',
  '初级效果',
  '中级效果',
  '高级效果',
  '终级效果',
  '一级效果',
  '二级效果',
  '三级效果',
  '四级效果',
  '五级效果',
  '效果',
  '效果1',
  '效果2',
  '效果3',
]);

function cleanFieldKey(key: string): string {
  return key.trim().replace(/\s+/g, '');
}

function parseStructuredField(line: string): { key: string; value: string } | undefined {
  const t = line.trim();
  const square = t.match(/^\[([^\]]+)\]\s*[:：]\s*(.*)$/);
  if (square) return { key: cleanFieldKey(square[1]), value: square[2].trim() };

  const full = t.match(/^【([^】]+)】\s*[:：]\s*(.*)$/);
  if (full) {
    const key = cleanFieldKey(full[1]);
    if (STRUCTURED_FIELD_KEYS.has(key)) return { key, value: full[2].trim() };
  }

  const bare = t.match(/^([^【】\[\]\s:：]{1,12})\s*[:：]\s*(.*)$/);
  if (bare) {
    const key = cleanFieldKey(bare[1]);
    if (STRUCTURED_FIELD_KEYS.has(key)) return { key, value: bare[2].trim() };
  }

  return undefined;
}

function parseAbilityNameLine(line: string): string | undefined {
  const field = parseStructuredField(line);
  if (field?.key !== '能力名称') return undefined;
  return field.value;
}

function isAbilityNameLine(line: string): boolean {
  return parseAbilityNameLine(line) !== undefined;
}

function isStructuredFieldLine(line: string): boolean {
  return parseStructuredField(line) !== undefined;
}

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
  return isAbilityNameLine(nextMeaningfulLine(lines, nextIndex));
}

function parseLevelTreeLine(line: string): { level: string; text: string } | undefined {
  const prefixed = line.trim().replace(LEVEL_TREE_PREFIX_RE, '');
  const patterns: Array<RegExp> = [
    new RegExp(`^[【\\[]\\s*(EX|S|A|B|C|D|E|F)\\s*级\\s*[】\\]]\\s*[:：]?\\s*(.+)$`, 'i'),
    new RegExp(`^[【\\[]\\s*(EX|S|A|B|C|D|E|F)\\s*[】\\]]\\s*(?:级)?\\s*[:：]?\\s*(.+)$`, 'i'),
    new RegExp(`^(EX|S|A|B|C|D|E|F)\\s*级\\s*[:：]\\s*(.+)$`, 'i'),
    new RegExp(`^(EX|S|A|B|C|D|E|F)\\s*级\\s*((?:${LEVEL_TREE_UNLOCK_WORDS}).+)$`, 'i'),
    new RegExp(`^(EX|S|A|B|C|D|E|F)\\s*[:：]\\s*(.+)$`, 'i'),
  ];
  for (const pattern of patterns) {
    const m = prefixed.match(pattern);
    if (!m) continue;
    const text = m[2].trim();
    if (!text || /^\[[^\]]+\]\s*[:：]/.test(text)) continue;
    return { level: m[1].toUpperCase(), text };
  }
  return undefined;
}

function parseLevelTreeHeading(line: string): { level: string } | undefined {
  const prefixed = line.trim().replace(LEVEL_TREE_PREFIX_RE, '');
  const m = prefixed.match(/^[【\[]\s*(EX|S|A|B|C|D|E|F)\s*级\s*[】\]]\s*$/i);
  if (!m) return undefined;
  return { level: m[1].toUpperCase() };
}

function isHardBlockBoundary(line: string): boolean {
  if (isAbilityNameLine(line) || isStructuredFieldLine(line)) return true;
  const bare = line.match(BANNER_BARE_RE);
  if (bare && !parseLevelTreeHeading(line)) return true;
  const inline = line.match(BANNER_INLINE_RE);
  return Boolean(inline && !parseLevelTreeLine(line));
}

function shouldAttachLevelContinuation(currentText: string, line: string): boolean {
  if (/^(?:注|备注|PS|P\.S)\s*[:：]/i.test(line)) return false;
  if (isHardBlockBoundary(line) || parseLevelTreeLine(line) || parseLevelTreeHeading(line)) return false;
  if (/^(?:解锁条件|效果)\s*[:：]/.test(line)) return true;
  if (/^[^:：\s]{2,14}\s*[:：]/.test(line) && /(?:不同|选择|道路|分支|称号|效果|如下|以下)/.test(currentText)) return true;
  return false;
}

function readLevelTreeItem(lines: string[], start: number): { item: { level: string; text: string }; next: number } | undefined {
  const current = lines[start].trim();
  const inline = parseLevelTreeLine(current);
  if (inline) {
    const textParts = [inline.text];
    let i = start + 1;
    while (i < lines.length) {
      const next = lines[i].trim();
      if (!next || !shouldAttachLevelContinuation(textParts.join('\n'), next)) break;
      textParts.push(next);
      i++;
    }
    return { item: { level: inline.level, text: textParts.join('\n') }, next: i };
  }

  const heading = parseLevelTreeHeading(current);
  if (!heading) return undefined;
  const body: string[] = [];
  let i = start + 1;
  while (i < lines.length) {
    const next = lines[i].trim();
    if (!next || parseLevelTreeLine(next) || parseLevelTreeHeading(next) || isHardBlockBoundary(next)) break;
    body.push(next);
    i++;
  }
  const text = body.join('\n').trim();
  if (!text) return undefined;
  return { item: { level: heading.level, text }, next: i };
}

function readLevelTree(lines: string[], start: number): { items: Array<{ level: string; text: string }>; next: number } | undefined {
  const items: Array<{ level: string; text: string }> = [];
  let i = start;
  while (i < lines.length) {
    const current = lines[i].trim();
    if (!current) {
      const next = nextMeaningfulLine(lines, i + 1);
      if (parseLevelTreeLine(next) || parseLevelTreeHeading(next)) {
        i++;
        continue;
      }
      break;
    }
    const result = readLevelTreeItem(lines, i);
    if (!result) break;
    items.push(result.item);
    i = result.next;
  }
  if (new Set(items.map((item) => item.level)).size < 3) return undefined;
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
    const field = parseStructuredField(t);
    if (field) {
      if (field.key === '能力名称') break;
      const key = field.key;
      if (key === firstKey) break;
      fields.push({ key, value: field.value });
      i++;
      continue;
    }
    if (isAbilityNameLine(t) || BANNER_BARE_RE.test(t) || BANNER_INLINE_RE.test(t)) break;
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

  const flushPlainPara = (parts: string[]) => {
    const text = parts.join('\n').trim();
    if (text) blocks.push({ kind: 'paragraph', text });
  };

  const flushPara = () => {
    if (!pendingPara.length) return;
    let plain: string[] = [];
    let i = 0;
    while (i < pendingPara.length) {
      const item = parseLevelTreeLine(pendingPara[i]);
      if (!item) {
        plain.push(pendingPara[i]);
        i++;
        continue;
      }
      const treeItems: Array<{ level: string; text: string }> = [item];
      i++;
      while (i < pendingPara.length) {
        const nextItem = parseLevelTreeLine(pendingPara[i]);
        if (!nextItem) break;
        treeItems.push(nextItem);
        i++;
      }
      if (treeItems.length >= 3) {
        flushPlainPara(plain);
        plain = [];
        blocks.push({ kind: 'level-tree', items: treeItems });
      } else {
        plain.push(...treeItems.map((treeItem) => `${treeItem.level}级：${treeItem.text}`));
      }
    }
    flushPlainPara(plain);
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

    const abilityName = parseAbilityNameLine(t);
    if (abilityName !== undefined) {
      flushPara();
      const head = extractAbilityHeader(abilityName.trim());
      const fields: Array<{ key: string; value: string }> = [];
      const tail: string[] = [];
      i++;
      while (i < lines.length) {
        const t2 = lines[i].trim();
        if (!t2) {
          const next = nextMeaningfulLine(lines, i + 1);
          if (next && isStructuredFieldLine(next) && !isAbilityNameLine(next)) {
            i++;
            continue;
          }
          if (next && (isAbilityNameLine(next) || (BANNER_BARE_RE.test(next) && !parseLevelTreeHeading(next)) || (BANNER_INLINE_RE.test(next) && !parseLevelTreeLine(next)))) {
            break;
          }
          if (next && fields.length) {
            i++;
            continue;
          }
          break;
        }
        if (isAbilityNameLine(t2)) break;
        if (BANNER_BARE_RE.test(t2) && !parseLevelTreeHeading(t2)) break;
        const field = parseStructuredField(t2);
        if (field && field.key !== '能力名称') {
          fields.push({ key: field.key, value: field.value });
          i++;
          // 续行
          while (i < lines.length) {
            const next = lines[i].trim();
            if (!next) break;
            if (
              isAbilityNameLine(next) ||
              isStructuredFieldLine(next) ||
              (BANNER_BARE_RE.test(next) && !parseLevelTreeHeading(next)) ||
              (BANNER_INLINE_RE.test(next) && !parseLevelTreeLine(next))
            ) break;
            const last = fields[fields.length - 1];
            last.value = (last.value ? last.value + '\n' : '') + next;
            i++;
          }
        } else if ((parseLevelTreeLine(t2) || parseLevelTreeHeading(t2)) && fields.length) {
          const last = fields[fields.length - 1];
          last.value = (last.value ? last.value + '\n' : '') + t2;
          i++;
        } else {
          if (fields.length) {
            const last = fields[fields.length - 1];
            last.value = (last.value ? last.value + '\n' : '') + t2;
          } else {
            tail.push(t2);
          }
          i++;
        }
      }
      blocks.push({ kind: 'ability', ...head, fields: filledFields(fields), tail: tail.join('\n') || undefined });
      continue;
    }

    const cardMatch = parseStructuredField(t);
    if (cardMatch) {
      const firstKey = cardMatch.key;
      const cardType = CARD_STARTERS[firstKey];
      if (cardType) {
        flushPara();
        const { fields, next } = readFieldCard(lines, i, firstKey, cardMatch.value);
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
          if (isAbilityNameLine(next) || BANNER_BARE_RE.test(next) || BANNER_INLINE_RE.test(next)) break;
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
