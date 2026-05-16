export const kb = (n: number) => n < 1024 ? `${n} B` : `${(n / 1024).toFixed(n > 1024 * 100 ? 0 : 1)} KB`;

export const words = (s: string) => Array.from(s.replace(/\s+/g, '')).length;

export const stamp = (path: string) =>
  (Math.abs(Array.from(path).reduce((a, c) => ((a << 5) - a + c.charCodeAt(0)) | 0, 0)) % 9000 + 1000).toString();

export const suggestHonorCategory = (path: string, category: string) => {
  const top = path.split('/')[1] || category.split('/')[0] || '';
  const map: Record<string, string> = {
    '特质改造': '特质',
    '职业': '职业',
    '技能表': '技能表',
    '能量池': '能量池',
    '魔药列表': '魔药列表',
    '成就': '成就',
  };
  return map[top] || '其他';
};

/** 在 text 里把 needles 的命中位置高亮成 <mark>。所有 needle 大小写不敏感。 */
export const highlight = (text: string, needles: string[]): Array<{ text: string; mark: boolean }> => {
  if (!text) return [];
  const valid = needles.filter(n => n && n.length > 0);
  if (valid.length === 0) return [{ text, mark: false }];
  // 用 lower 找位置；大小写不敏感
  const lower = text.toLowerCase();
  const ranges: Array<[number, number]> = [];
  for (const n of valid) {
    const lc = n.toLowerCase();
    let i = 0;
    while (i < lower.length) {
      const idx = lower.indexOf(lc, i);
      if (idx < 0) break;
      ranges.push([idx, idx + lc.length]);
      i = idx + lc.length;
    }
  }
  if (ranges.length === 0) return [{ text, mark: false }];
  ranges.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const merged: Array<[number, number]> = [];
  for (const r of ranges) {
    const last = merged[merged.length - 1];
    if (last && r[0] <= last[1]) last[1] = Math.max(last[1], r[1]);
    else merged.push([...r] as [number, number]);
  }
  const out: Array<{ text: string; mark: boolean }> = [];
  let cur = 0;
  for (const [s, e] of merged) {
    if (cur < s) out.push({ text: text.slice(cur, s), mark: false });
    out.push({ text: text.slice(s, e), mark: true });
    cur = e;
  }
  if (cur < text.length) out.push({ text: text.slice(cur), mark: false });
  return out;
};
