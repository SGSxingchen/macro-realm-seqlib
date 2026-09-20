const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });
const resources = ['时空锚定', '跨界旅人', '能量核心'].map((name, index) => ({
  path: `序列库/技能表/科技侧/00${index + 1}】${name}.txt`, title: name, filename: `${name}.txt`, root: '序列库', category: '技能表/科技侧', top_kind: '技能表', side: '科技侧', size: 5000, mtime: 1, encoding: 'utf-8',
  content: `${name}\n\n不规则开头保留。\n\n[能力名称]：共同名称\n[能力效果]：规则不可省略。\n\n` + Array.from({ length: 24 }, (_, i) => `[能力名称]：${name} ${i}\n[自定义代价]：不属于固定模板。\n[能力效果]：每一条效果均保留完整条件与限制。\n\n`).join(''), snippet: '不同格式完整保留，打开进行对照。',
}));
const oldPath = '序列库/技能表/科技侧/099】旧名称.txt';
const delay = ms => new Promise(resolve => setTimeout(resolve, ms));

async function setup(browser, width) {
  const context = await browser.newContext({ viewport: { width, height: 1000 }, colorScheme: 'dark' });
  const errors = [];
  const calls = [];
  context.on('page', page => page.on('pageerror', error => errors.push(error.message)));
  await context.route('**/api/**', async route => {
    const url = new URL(route.request().url());
    calls.push(url);
    let data, status = 200;
    if (url.pathname === '/api/tree') data = { items: [{ path: '序列库', name: '序列库', count: resources.length, children: [] }] };
    else if (url.pathname === '/api/resources') data = { items: resources, count: resources.length, total: resources.length, offset: 0, limit: 100, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
    else if (url.pathname.startsWith('/api/resources/')) {
      data = resources.find(item => item.path === decodeURIComponent(url.pathname.slice('/api/resources/'.length)));
      if (!data) { status = 404; data = { detail: 'missing' }; }
    } else if (url.pathname === '/api/git/changes') {
      data = { from_ref: url.searchParams.get('from_ref') || 'v-default', to: 'working-tree/latest', readable: { added: [], modified: [], deleted: [], renamed: resources.map(item => ({ path: item.path, old_path: oldPath })) }, raw: { committed: { returncode: 0 } } };
    } else if (url.pathname.startsWith('/api/git/change-detail/')) {
      const baseline = url.searchParams.get('from_ref') || 'v-default';
      if (baseline === 'slow') await delay(400);
      data = { from_ref: baseline, to: 'working-tree/latest', path: decodeURIComponent(url.pathname.slice('/api/git/change-detail/'.length)), old_path: url.searchParams.get('old_path'), kind: 'renamed', old_exists: !url.searchParams.get('old_path').includes('missing'), new_exists: true, old_encoding: 'utf-8', new_encoding: 'utf-8', old_line_count: 20, new_line_count: 21, additions: 1, deletions: 1, truncated: baseline === 'truncated', rows: [
        { type: 'context', old_no: 1, new_no: 1, text: '[能力名称]：共同名称' },
        { type: 'removed', old_no: 2, text: '[消耗能量]：旧版 10' },
        { type: 'added', new_no: 2, text: '[消耗能量]：新版 12' },
        { type: 'gap' },
      ] };
    } else { status = 404; data = { detail: 'unsupported mock' }; }
    try { await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(data) }); } catch { /* cancelled request */ }
  });
  const page = await context.newPage();
  page.setDefaultTimeout(12000);
  await page.goto(base);
  await page.locator('[data-resource-index="0"]').waitFor();
  return { page, context, calls, errors };
}
const primary = page => page.locator('.workbench-primary');
const reference = page => page.locator('.workbench-reference');
const tab = (page, i) => page.locator('.workbench-tab-title').filter({ hasText: resources[i].title });
async function titleIs(page, index, pane = '.workbench-primary') {
  await page.waitForFunction(({ pane, title }) => document.querySelector(pane + ' .realm-document-heading h1')?.textContent === title, { pane, title: resources[index].title });
}
async function shot(page, name) { await page.screenshot({ path: path.join(out, 'workbench-' + name + '.png'), fullPage: true, animations: 'disabled' }); }
async function noOverflow(page) {
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true);
  const values = await page.locator('.realm-reader-scroll').evaluateAll(nodes => nodes.filter(node => node.clientWidth).map(node => node.scrollWidth <= node.clientWidth + 1));
  assert.ok(values.every(Boolean));
}
module.exports = { setup, resources, oldPath, titleIs, noOverflow, primary, reference, tab };

if (require.main === module) (async () => {
  const browser = await chromium.launch();
  let active;
  try {
    const desk = await setup(browser, 1440); const { page } = desk; active = page;
    await page.locator('[data-resource-index="0"]').click(); await titleIs(page, 0);
    await page.getByRole('button', { name: '收起索引', exact: true }).click();
    assert.equal(await page.locator('.realm-navigation').isVisible(), false);
    await page.getByRole('button', { name: '收起结果', exact: true }).click();
    assert.equal(await page.locator('.realm-results').isVisible(), false);
    await noOverflow(page); await shot(page, 'both-collapsed');
    await page.getByRole('button', { name: '展开结果', exact: true }).click();
    await page.getByRole('button', { name: '专注阅读', exact: true }).click();
    assert.equal(await page.locator('.realm-results').isVisible(), false);
    await page.getByRole('button', { name: '退出专注', exact: true }).click();
    assert.equal(await page.locator('.realm-results').isVisible(), true);
    assert.equal(await page.locator('.realm-navigation').isVisible(), false, 'Exit focus must restore the independently collapsed index');
    await page.getByRole('button', { name: '专注阅读', exact: true }).click();
    await page.keyboard.press('Control+k');
    await page.waitForFunction(() => document.activeElement.id === 'realm-query');
    await primary(page).getByRole('button', { name: '原文', exact: true }).click();
    await primary(page).locator('.realm-reader-scroll').evaluate(node => { node.scrollTop = 650; });
    await page.locator('[data-resource-index="1"]').click(); await titleIs(page, 1);
    await page.locator('[data-resource-index="2"]').click(); await titleIs(page, 2);
    assert.equal(await page.locator('.workbench-tab').count(), 3);
    await tab(page, 0).click(); await titleIs(page, 0);
    assert.equal(await primary(page).getByRole('button', { name: '原文', exact: true }).getAttribute('aria-pressed'), 'true');
    assert.ok(await primary(page).locator('.realm-reader-scroll').evaluate(node => node.scrollTop >= 640), 'Each tab keeps its own reading position');
    await tab(page, 0).click(); assert.equal(await page.locator('.workbench-tab').count(), 3);
    await primary(page).getByRole('button', { name: '自适应', exact: true }).click();
    await page.getByRole('button', { name: '并排对照', exact: true }).click();
    await titleIs(page, 2, '.workbench-reference');
    assert.equal(await page.locator('.realm-results').isVisible(), false, 'Comparison should free reading space');
    await reference(page).locator('.realm-reader-scroll').evaluate(node => { node.scrollTop = 500; });
    await tab(page, 1).click(); await titleIs(page, 1);
    await titleIs(page, 2, '.workbench-reference');
    assert.ok(await reference(page).locator('.realm-reader-scroll').evaluate(node => node.scrollTop >= 490), 'Pinned pane must not jump when primary tab changes');
    assert.equal(await page.evaluate(() => { const ids = Array.from(document.querySelectorAll('[id]')).map(node => node.id); return new Set(ids).size === ids.length; }), true, 'Comparison must not duplicate anchor IDs');
    await page.getByRole('button', { name: '对照设置', exact: true }).click();
    await page.getByRole('button', { name: '交换主 / 参照', exact: true }).click();
    await titleIs(page, 2); await titleIs(page, 1, '.workbench-reference');
    await noOverflow(page); await shot(page, 'desktop-compare');
    await page.getByRole('button', { name: '退出并排', exact: true }).click();
    assert.equal(await page.locator('.realm-results').isVisible(), true);
    assert.equal(await page.locator('.realm-navigation').isVisible(), false);
    await page.getByRole('button', { name: '关闭档案：跨界旅人', exact: true }).click();
    await titleIs(page, 2); assert.equal(await page.locator('.workbench-tab').count(), 2);
    await page.reload(); await titleIs(page, 2);
    assert.equal(await page.locator('.workbench-tab').count(), 2, 'Tabs survive reload in this browser session');
    assert.equal(await page.locator('.realm-navigation').isVisible(), false);

    await page.getByRole('button', { name: '历史版本', exact: true }).click();
    const history = page.getByRole('region', { name: '历史版本对照' });
    await history.getByRole('textbox', { name: '历史基准版本' }).fill('v-old');
    await history.getByRole('button', { name: '查找改名候选' }).click();
    await history.getByRole('button', { name: '使用此候选比较' }).waitFor();
    assert.equal(desk.calls.filter(url => url.pathname.includes('/change-detail/')).length, 0, 'A heuristic rename candidate is not silently accepted');
    await history.getByRole('button', { name: '使用此候选比较' }).click();
    await history.locator('.history-line.added').waitFor();
    assert.equal(await history.getByRole('textbox', { name: '旧版文件路径' }).inputValue(), oldPath);
    assert.ok((await history.locator('.history-sources').textContent()).includes('v-old'));
    await shot(page, 'history-diff');
    await history.getByRole('textbox', { name: '旧版文件路径' }).fill('序列库/missing.txt');
    await history.getByRole('button', { name: '按此路径比较' }).click();
    await history.getByRole('alert').filter({ hasText: '不能把缺失记录' }).waitFor();
    assert.equal(await history.locator('.history-line').count(), 0);
    await history.getByRole('textbox', { name: '旧版文件路径' }).fill(oldPath);
    await history.getByRole('textbox', { name: '历史基准版本' }).fill('truncated');
    await history.getByRole('button', { name: '按此路径比较' }).click();
    await history.getByRole('alert').filter({ hasText: '不是完整差异' }).waitFor();
    await history.getByRole('textbox', { name: '历史基准版本' }).fill('slow');
    await history.getByRole('button', { name: '按此路径比较' }).click();
    await history.getByRole('textbox', { name: '历史基准版本' }).fill('v-new');
    await history.getByRole('button', { name: '按此路径比较' }).click();
    await history.waitFor({ state: 'visible' });
    await page.waitForFunction(() => document.querySelector('.history-sources')?.textContent.includes('v-new'));
    await delay(500);
    assert.ok(!(await history.locator('.history-sources').textContent()).includes('slow'));
    await noOverflow(page);
    assert.deepEqual(desk.errors, []);
    await desk.context.close();

    for (const width of [1024, 390, 320]) {
      const small = await setup(browser, width); const p = small.page; active = p;
      await p.locator('[data-resource-index="0"]').click(); await titleIs(p, 0);
      if (width <= 820) await p.locator('.workbench-primary .realm-back').click();
      await p.locator('[data-resource-index="1"]').click(); await titleIs(p, 1);
      await p.getByRole('button', { name: '并排对照', exact: true }).click();
      await titleIs(p, 0, '.workbench-reference');
      if (width <= 820) {
        assert.equal(await primary(p).isVisible(), true);
        assert.equal(await reference(p).isVisible(), false);
        await p.getByRole('button', { name: '参照档案', exact: true }).click();
        assert.equal(await reference(p).isVisible(), true);
        assert.equal(await primary(p).isVisible(), false);
        await p.getByRole('button', { name: '主档案', exact: true }).click();
      }
      await noOverflow(p); await shot(p, width + '-compare');
      await p.getByRole('button', { name: '历史版本', exact: true }).click();
      await noOverflow(p); await shot(p, width + '-history');
      await p.getByRole('button', { name: '返回正文', exact: true }).click();
      await p.getByRole('button', { name: '关闭档案：跨界旅人', exact: true }).click();
      await titleIs(p, 0);
      assert.deepEqual(small.errors, []);
      await small.context.close();
    }
    fs.writeFileSync(path.join(out, 'workbench-summary.txt'), 'PASS: independently collapsible panels and restoration; focus/search recovery; deduplicated persistent tabs; per-resource view and scroll state; independent pinned comparison and unique anchors; close/swap; explicit history rename confirmation; missing and truncated history guards; stale diff cancellation; responsive 1440/1024/390/320px layouts. API fixtures are simulated, not production history.\n');
    console.log('Workbench browser regressions passed.');
  } catch (error) {
    fs.writeFileSync(path.join(out, 'workbench-failure.txt'), error.stack || String(error));
    if (active && !active.isClosed()) await shot(active, 'failure').catch(() => {});
    throw error;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
