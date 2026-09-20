// Layout regressions use the same deterministic six-document fixture on the
// base and proposed revisions. No production endpoint is contacted.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const baseline = process.argv.includes('--baseline');
const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });
const resources = ['通用仪式', '猎人仪式', '密教仪式', '仪式学技能组（注解）', '团长技能表：奥尔加·伊兹卡', '菜狗种散装技能'].map((name, index) => ({
  path: `序列库/技能表/测试侧/00${index + 1}】${name}.txt`, title: `00${index + 1}】${name}`, filename: `${name}.txt`, root: '序列库', category: '技能表/测试侧', size: 5000, mtime: 1, encoding: 'utf-8',
  snippet: '布局回归测试资料，不是线上游戏规则。',
  content: `${name}\n\n布局回归测试资料，不是线上游戏规则。\n（制作人：测试）\n\n【自由格式说明】\n不符合标准模板的段落完整保留。\n不隐藏规则、数值、限制和更新记录。\n\n` + Array.from({ length: 24 }, (_, i) => `[能力名称]：${name} ${i}\n[自定义代价]：不是标准模板中的字段。\n[能力效果]：这一段用于验证视口、排版与原文一致性。\n\n注意：空行后自由文字仍须保留。\n`).join(''),
}));

async function setup(browser, width, height = 930) {
  const context = await browser.newContext({ viewport: { width, height }, colorScheme: 'light', permissions: ['clipboard-read', 'clipboard-write'] });
  const errors = [];
  context.on('page', page => page.on('pageerror', error => errors.push(error.message)));
  await context.route('**/api/**', async route => {
    const url = new URL(route.request().url());
    let data, status = 200;
    if (url.pathname === '/api/tree') data = { items: [{ path: '序列库', name: '序列库', count: 6, children: [] }] };
    else if (url.pathname === '/api/resources') data = { items: resources.map(({ content, ...item }) => item), count: 6, total: 6, offset: 0, limit: 100, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
    else if (url.pathname.startsWith('/api/resources/')) {
      data = resources.find(item => item.path === decodeURIComponent(url.pathname.slice('/api/resources/'.length)));
      if (!data) { status = 404; data = { detail: 'missing' }; }
    } else { status = 404; data = { detail: 'Unsupported mock API' }; }
    try { await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(data) }); } catch { /* cancelled */ }
  });
  const page = await context.newPage();
  page.setDefaultTimeout(12000);
  await page.goto(base);
  await page.locator('[data-resource-index="0"]').waitFor();
  return { context, page, errors };
}
async function ready(page, i, pane = '.workbench-primary') {
  await page.waitForFunction(({ pane, title }) => document.querySelector(pane + ' .realm-document-heading h1')?.textContent === title, { pane, title: resources[i].title });
}
async function openSix(page, width) {
  for (let i = 0; i < resources.length; i++) {
    if (width <= 820 && i) {
      await page.locator('.workbench-primary .realm-back').click();
      await page.locator('.realm-results').waitFor({ state: 'visible' });
    }
    await page.locator(`[data-resource-index="${i}"]`).click();
    await ready(page, i);
  }
  assert.equal(await page.locator('.workbench-tab').count(), 6);
  await page.getByRole('button', { name: '并排对照', exact: true }).click();
  await ready(page, 4, '.workbench-reference');
}
async function metrics(page) {
  return page.evaluate(() => {
    const panes = Array.from(document.querySelectorAll('.realm-reader-scroll')).filter(node => node.clientWidth && node.clientHeight);
    const rect = panes[0].getBoundingClientRect();
    return { width: innerWidth, height: innerHeight, controlsHeight: Math.round(rect.top), readingHeight: Math.round(rect.height), readingShare: Number((rect.height / innerHeight).toFixed(4)), commandBar: !!document.querySelector('.workbench-command-bar') };
  });
}
async function geometry(page) {
  assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true, 'Document overflows the viewport');
  assert.ok((await page.locator('.realm-reader-scroll').evaluateAll(nodes => nodes.filter(node => node.clientWidth).map(node => node.scrollWidth <= node.clientWidth + 1))).every(Boolean), 'Reader overflows horizontally');
  const clipped = await page.evaluate(() => Array.from(document.querySelectorAll('.workbench-controls button, .realm-layout-controls button, .realm-reader-toolbar button, .realm-reader-toolbar summary')).filter(node => {
    const rect = node.getBoundingClientRect();
    if (!rect.width || !rect.height || !node.checkVisibility()) return false;
    // Closed disclosure descendants have no rendered box. Open menus are tested separately.
    return rect.left < -1 || rect.right > innerWidth + 1 || rect.top < -1 || rect.bottom > innerHeight + 1;
  }).map(node => node.textContent));
  assert.deepEqual(clipped, [], 'A persistent operation is clipped');
}
async function shot(page, name) { await page.screenshot({ path: path.join(out, name + '.png'), fullPage: true, animations: 'disabled' }); }

(async () => {
  const browser = await chromium.launch();
  const results = [];
  let active;
  try {
    for (const width of baseline ? [1600] : [1600, 1024, 821, 820, 390, 320]) {
      const fixture = await setup(browser, width, width <= 820 ? 844 : 930);
      const { page } = fixture; active = page;
      await openSix(page, width);
      if (baseline) {
        await page.getByRole('combobox', { name: '选择参照档案' }).selectOption(resources[3].path);
        await ready(page, 3, '.workbench-reference');
        results.push(await metrics(page));
        await shot(page, 'focus-before-1600');
        await fixture.context.close();
        continue;
      }
      await page.locator('.realm-topbar').waitFor({ state: 'hidden' });
      if (width <= 820) await page.getByRole('button', { name: '参照档案', exact: true }).click();
      const beforeMenu = await metrics(page);
      const chooser = page.locator('.reader-reference-picker > summary');
      await chooser.click();
      await page.getByRole('combobox', { name: '选择参照档案' }).waitFor({ state: 'visible' });
      assert.equal((await metrics(page)).controlsHeight, beforeMenu.controlsHeight, 'Reference picker must overlay, not allocate another row');
      await page.getByRole('combobox', { name: '选择参照档案' }).selectOption(resources[3].path);
      await ready(page, 3, '.workbench-reference');
      await page.locator('.reader-reference-popover').waitFor({ state: 'hidden' });
      await chooser.click();
      await page.keyboard.press('Escape');
      assert.equal(await page.evaluate(() => document.activeElement?.className === '' && document.activeElement?.parentElement?.classList.contains('reader-reference-picker')), true, 'Escape restores focus to the disclosure');
      await chooser.click();
      await page.locator('.realm-layout-controls').click({ position: { x: 2, y: 2 } });
      await page.locator('.reader-reference-popover').waitFor({ state: 'hidden' });
      if (width <= 820) await page.getByRole('button', { name: '主档案', exact: true }).click();
      await geometry(page);
      const m = await metrics(page); results.push(m);
      assert.ok(m.controlsHeight <= (width > 820 ? 165 : 235), `Reading chrome height budget exceeded: ${JSON.stringify(m)}`);
      await shot(page, `focus-after-${width}`);

      // Native DOM nodes survive enter/exit focus; views are not unmounted to
      // obtain smaller chrome, and the reference remains pinned.
      if (width === 1600) {
        await page.evaluate(() => { window.__readingNode = document.querySelector('.workbench-primary .realm-reader-scroll'); window.__referenceNode = document.querySelector('.workbench-reference .realm-reader-scroll'); });
        const primary = page.locator('.workbench-primary');
        await primary.getByRole('button', { name: '原文', exact: true }).click();
        assert.equal(await primary.locator('.realm-document-content').textContent(), resources[5].content);
        await primary.locator('.realm-reader-scroll').evaluate(node => { node.scrollTop = 650; });
        await page.getByRole('button', { name: '退出专注', exact: true }).click();
        await page.locator('.realm-topbar').waitFor({ state: 'visible' });
        await page.getByRole('button', { name: '专注阅读', exact: true }).click();
        await page.locator('.realm-topbar').waitFor({ state: 'hidden' });
        assert.equal(await page.evaluate(() => window.__readingNode === document.querySelector('.workbench-primary .realm-reader-scroll') && window.__referenceNode === document.querySelector('.workbench-reference .realm-reader-scroll')), true);
        assert.equal(await primary.getByRole('button', { name: '原文', exact: true }).getAttribute('aria-pressed'), 'true');
        assert.ok(await primary.locator('.realm-reader-scroll').evaluate(node => node.scrollTop > 500));
        await page.keyboard.press('Control+k');
        await page.waitForFunction(() => document.activeElement.id === 'realm-query');
        await page.locator('.realm-topbar').waitFor({ state: 'visible' });
        await page.getByRole('button', { name: '专注阅读', exact: true }).click();
        await primary.getByRole('button', { name: '自适应', exact: true }).click();
        await primary.locator('.realm-reader-scroll').evaluate(node => { node.scrollTop = 0; });
        // Theme control deliberately lives in full-site navigation: exit focus,
        // switch theme, re-enter, and keep the same active/reference documents.
        await page.getByRole('button', { name: '退出专注', exact: true }).click();
        await page.getByRole('button', { name: '切换主题' }).click();
        await page.waitForFunction(() => document.documentElement.dataset.theme === 'dark');
        await page.getByRole('button', { name: '专注阅读', exact: true }).click();
        await shot(page, 'focus-after-1600-dark');
        await page.getByRole('button', { name: '退出并排', exact: true }).click();
        await shot(page, 'focus-single-1600-dark');
        assert.ok((await metrics(page)).controlsHeight <= 165);
        await primary.getByRole('button', { name: '复制全文', exact: true }).click();
        assert.equal(await page.evaluate(() => navigator.clipboard.readText()), resources[5].content);
      }
      assert.deepEqual(fixture.errors, []);
      await fixture.context.close();
    }
    const filename = baseline ? 'focus-baseline-metrics.json' : 'focus-metrics.json';
    fs.writeFileSync(path.join(out, filename), JSON.stringify(results, null, 2));
    if (!baseline && fs.existsSync(path.join(out, 'focus-baseline-metrics.json'))) {
      const old = JSON.parse(fs.readFileSync(path.join(out, 'focus-baseline-metrics.json'), 'utf8'))[0];
      const current = results[0];
      if (!old.commandBar) assert.ok(current.controlsHeight < old.controlsHeight * 0.6, 'Must eliminate stacked rows, not only reduce font sizes');
      fs.writeFileSync(path.join(out, 'focus-comparison.json'), JSON.stringify({ before: old, after: current, reclaimedPixels: old.controlsHeight - current.controlsHeight }, null, 2));
    }
    console.log(baseline ? 'Baseline reading geometry captured.' : 'Focus layout, reference disclosure, restored state and viewport budget passed.');
  } catch (error) {
    fs.writeFileSync(path.join(out, 'focus-failure.txt'), error.stack || String(error));
    if (active && !active.isClosed()) await shot(active, 'focus-failure').catch(() => {});
    throw error;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
