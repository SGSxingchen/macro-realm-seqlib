const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const Module = require('node:module');
const ts = require('typescript');
const filename = path.resolve(__dirname, '../src/workbench-state.ts');
const mod = new Module(filename, module);
mod.filename = filename;
mod.paths = Module._nodeModulePaths(path.dirname(filename));
mod._compile(ts.transpileModule(fs.readFileSync(filename, 'utf8'), { compilerOptions: { module: ts.ModuleKind.CommonJS, target: ts.ScriptTarget.ES2021 } }).outputText, filename);
const { rememberTab, closeTab, readOpenTabs, storeOpenTabs, readPanePreferences, storePanePreferences } = mod.exports;
const a = { path: '序列库/A.txt', title: 'A' };
const b = { path: '序列库/B.txt', title: 'B' };
const c = { path: '序列库/C.txt', title: 'C' };

test('open tabs deduplicate by full path, not potentially duplicate titles', () => {
  let tabs = rememberTab([a], a.path, 'A更新');
  assert.equal(tabs.length, 1);
  assert.equal(tabs[0].title, 'A更新');
  tabs = rememberTab(tabs, b.path, 'A更新');
  assert.equal(tabs.length, 2);
  assert.deepEqual(rememberTab(tabs, ''), tabs);
});
test('closing inactive tab preserves active; closing active selects adjacent', () => {
  assert.deepEqual(closeTab([a, b, c], b.path, a.path), { tabs: [a, c], active: a.path });
  assert.deepEqual(closeTab([a, b, c], b.path, b.path), { tabs: [a, c], active: a.path });
  assert.deepEqual(closeTab([a], a.path, a.path), { tabs: [], active: '' });
});
test('pane defaults and malformed local storage are safe', () => {
  global.localStorage = { getItem: () => '{broken', setItem: () => { throw Error('blocked'); } };
  assert.deepEqual(readPanePreferences(), { index: true, results: true });
  assert.doesNotThrow(() => storePanePreferences({ index: false, results: true }));
  localStorage.getItem = () => JSON.stringify({ index: false, results: true });
  assert.deepEqual(readPanePreferences(), { index: false, results: true });
  localStorage.getItem = () => 'null';
  assert.deepEqual(readPanePreferences(), { index: true, results: true });
});
test('session restoration ignores corrupt entries and duplicates', () => {
  global.sessionStorage = { getItem: () => JSON.stringify([null, a, a, b, { path: 'secret', title: 'x' }]), setItem: () => { throw Error('blocked'); } };
  assert.deepEqual(readOpenTabs(), [a, b]);
  assert.doesNotThrow(() => storeOpenTabs([a]));
  sessionStorage.getItem = () => '{}';
  assert.deepEqual(readOpenTabs(), []);
});
