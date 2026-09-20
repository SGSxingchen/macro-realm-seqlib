const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');

function loadTS(relative) {
  const filename = path.resolve(__dirname, relative);
  const source = fs.readFileSync(filename, 'utf8');
  const output = ts.transpileModule(source, { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2021 }, reportDiagnostics: true });
  assert.equal((output.diagnostics || []).filter(d => d.category === ts.DiagnosticCategory.Error).length, 0);
  const mod = new Module(filename, module);
  mod.filename = filename;
  mod.paths = Module._nodeModulePaths(path.dirname(filename));
  mod._compile(output.outputText, filename);
  return mod.exports;
}
const { buildDocument, readField, resourceLink } = loadTS('../src/components/Reader/adaptive.ts');
const { readLocation, locationUrl, loadSaved, saveResources } = loadTS('../src/library-state.ts');

function assertLossless(text) {
  const sections = buildDocument(text);
  assert.equal(sections.map(section => section.raw).join(''), text);
  assert.equal(new Set(sections.map(section => section.id)).size, sections.length);
  let end = 0;
  for (const section of sections) {
    assert.equal(section.start, end);
    assert.equal(section.raw, text.slice(section.start, section.end));
    assert.equal(section.parts.map(part => part.raw).join(''), section.raw);
    for (const part of section.parts) if (part.field) assert.equal(part.field.prefix + part.field.value, part.raw);
    end = section.end;
  }
  assert.equal(end, text.length);
  return sections;
}

test('standard fields retain costs, limits, notes and indentation verbatim', () => {
  const text = '标题\n\n[能力名称]：锚定（A级）\n[消耗能量]：10点，不可减免。\n[能力效果]：第一行\n  缩进\n\n注：离开范围立即失效。\n';
  const sections = assertLossless(text);
  assert.equal(sections[1].kind, 'entry');
  assert.equal(sections[1].raw, text.slice(text.indexOf('[能力名称]')));
  assert.ok(sections[1].raw.includes('不可减免'));
  assert.ok(sections[1].raw.includes('离开范围立即失效'));
});
test('legacy full-width, bare and colonless keys remain readable', () => {
  for (const text of ['【能力名称】：旧版\r\n【能力效果】：效果\r\n续行\r\n', '能力名称: 旧版\n能力效果: 保留\n', '[能力名称]\n没有冒号的名称\n[能力效果]\n完整规则\n']) assertLossless(text);
  assert.equal(readField('【能力效果】\n').key, '能力效果');
});
test('unknown explicit fields are retained, ordinary prose is not guessed', () => {
  assert.equal(readField('[自定义代价]：支付未来的记忆。').key, '自定义代价');
  assert.equal(readField('注意：这是一段叙事。'), undefined);
  assert.equal(buildDocument('自由格式\n注意：先决条件不能删。\n自定义 ▲ ◎\n')[0].kind, 'intro');
});
test('mixed formats and unfamiliar headings preserve exact source order', () => {
  assertLossless('\ufeff初始标题\r\n\r\n【更新记录】\r\n2026/9/20 保留\n# 独立规则\n[技能名称]：异界（EX）【被动】\n[不存在于模板的字段]：仍应保留。\r未识别的条目\n');
});
test('duplicate names get distinct stable anchors', () => {
  const text = '[能力名称]：同名\n[能力效果]：A\n[能力名称]：同名\n[能力效果]：B';
  const sections = assertLossless(text);
  assert.notEqual(sections[0].id, sections[1].id);
  assert.deepEqual(buildDocument('新增序言\n' + text).slice(1).map(s => s.id), sections.map(s => s.id));
});
test('blank content, all newline forms, emoji and literal HTML are lossless', () => {
  for (const text of ['', '\n\n', '\r', '\r\n', '终行无换行', '😀\t\t零宽\u200b字\n<script>alert(1)</script>\n', '【能力名称】：😀\n[能力效果]: <img src=x onerror=alert(1)>']) assertLossless(text);
});
test('seeded mixed-format corpus is lossless', () => {
  const atoms = ['\n', '\r\n', '说明。\n', '【标题】\n', '[能力名称]: A（S级）\n', '【自定义字段】：未知规则\n', '能力效果：不改写\n', '  1. 缩进\n', '[能力效果]\n', '注意: 限制\n', '😀', '\t', '# 章节\n'];
  let seed = 19491001;
  for (let trial = 0; trial < 400; trial++) {
    let text = '';
    for (let i = 0; i < 100; i++) { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; text += atoms[seed % atoms.length]; }
    assertLossless(text);
  }
});
test('real sequence-library TXT corpus is lossless when available', t => {
  const root = path.resolve(__dirname, '../../../序列库');
  if (!fs.existsSync(root)) { assert.ok(!process.env.CI, 'Full repository corpus must be available in CI'); t.skip('Repository corpus is not mounted in this local runtime; CI checks the full checkout.'); return; }
  let count = 0;
  const walk = directory => {
    for (const entry of fs.readdirSync(directory, { withFileTypes: true })) {
      const filename = path.join(directory, entry.name);
      if (entry.isDirectory()) walk(filename);
      else if (entry.name.toLowerCase().endsWith('.txt')) {
        const bytes = fs.readFileSync(filename);
        let text;
        try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); }
        catch { text = new TextDecoder('gb18030').decode(bytes); }
        try { assertLossless(text); } catch (error) { error.message = `${filename}: ${error.message}`; throw error; }
        count++;
      }
    }
  };
  walk(root);
  assert.ok(count > 0);
  t.diagnostic(`Validated source ranges and field round-trips for ${count} real TXT resources.`);
});
test('URL round-trip retains multi-facets, Unicode path and section anchor', () => {
  global.window = { location: { href: 'https://example.test/library', pathname: '/library' } };
  const state = { tab: 'read', filters: { q: '锚定 空间', category: '技能表/科技侧', kinds: ['技能表', '职业'], sides: ['科技侧'], authors: ['甲', '乙'] }, openPath: '序列库/001】空 格#号?.txt', anchor: 'entry-a-1' };
  const url = new URL(locationUrl(state), window.location.href).href;
  assert.deepEqual(readLocation(url), state);
  window.location.href = 'https://example.test/library?tab=updates&q=x&open=old#old';
  const share = new URL(resourceLink(state.openPath, 'entry-b-1'));
  assert.equal(share.searchParams.get('open'), state.openPath);
  assert.equal(share.searchParams.has('tab'), false);
  assert.equal(share.hash, '#entry-b-1');
});
test('damaged or unavailable bookmark storage does not crash the reader', () => {
  global.localStorage = { getItem: () => '{broken', setItem: () => { throw new Error('blocked'); } };
  assert.deepEqual(loadSaved(), []);
  assert.equal(saveResources([]), false);
  global.localStorage.getItem = () => JSON.stringify([null, 1, {}, { path: '荣誉室/a.txt', title: 'x' }, { path: '序列库/a.txt', title: 'A' }]);
  assert.deepEqual(loadSaved(), [{ path: '序列库/a.txt', title: 'A' }]);
});

test('free paragraphs are not swallowed into a name or preceding field', () => {
  const sections = assertLossless('[能力名称]：异界旅人\n不规则描述段落。\n\n[能力效果]：效果\n\n注意：仍是独立备注。\n');
  assert.equal(sections[0].parts[0].field.value.trim(), '异界旅人');
  assert.ok(sections[0].parts.some(part => !part.field && part.raw.includes('独立备注')));
  assert.equal(buildDocument('[能力名称]\n无冒号能力\n[能力效果]\n保留。')[0].kind, 'entry');
});
