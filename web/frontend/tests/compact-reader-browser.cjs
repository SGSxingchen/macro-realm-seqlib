const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });
const titles = ['通用仪式', '猎人仪式', '密教仪式', '仪式学技能组（注解）', '团长技能表：奥尔加·伊兹卡', '菜狗种散装技能'];
const resources = titles.map((title, i) => ({
  path: `序列库/技能表/其他及特殊/${String(i + 1).padStart(3, '0')}】${title}.txt`,
  title: `${String(i + 1).padStart(3, '0')}】${title}`, category: '技能表/其他及特殊', filename: title + '.txt',
  root: '序列库', size: 10000, mtime: 1, encoding: 'utf-8', snippet: '排版测试数据：条件与限制完整展示。',
  content: title + '\n\n此为界面回归测试数据，不是正式技能规则。\n\n' + Array.from({ length: 40 }, (_, n) => `[能力名称]：测试条目 ${n}\n[能力效果]：这是较长的测试正文，用于验证专注与并排阅读时，操作条不会占据过多纵向空间。\n[使用要求]：未知字段和自由备注仍然保留。\n\n`).join(''),
}));
const measurements = [];
async function ready(page, item, pane = '.workbench-primary') {
  await page.waitForFunction(({ title, pane }) => document.querySelector(pane + ' h1')?.textContent === title, { title: item.title, pane });
}
async function measure(page, label) {
  const values = await page.evaluate(() => {
    const visible = node => node && node.getBoundingClientRect().width > 0 && node.getBoundingClientRect().height > 0;
    const reader = Array.from(document.querySelectorAll('.realm-reader-scroll')).find(visible);
    const rect = reader.getBoundingClientRect();
    const tabs = document.querySelector('.workbench-tab-strip').getBoundingClientRect();
    const actions = document.querySelector('.workbench-controls').getBoundingClientRect();
    const fonts = Array.from(document.querySelectorAll('.archive-field-value')).filter(visible).map(node => parseFloat(getComputedStyle(node).fontSize));
    return { width: innerWidth, height: innerHeight, top: rect.top, readingHeight: rect.height, horizontalOverflow: document.documentElement.scrollWidth > innerWidth + 1,
      sameHeaderRow: Math.abs(tabs.top - actions.top) < 10, minimumBodyFont: Math.min(...fonts),
      controlsHeight: actions.height, maximumReaderOverflow: Math.max(...Array.from(document.querySelectorAll('.realm-reader-scroll')).filter(visible).map(node => node.scrollWidth - node.clientWidth)) };
  });
  measurements.push({ label, ...values });
  assert.equal(values.horizontalOverflow, false, JSON.stringify(values));
  assert.ok(values.maximumReaderOverflow <= 1, JSON.stringify(values));
  assert.ok(values.minimumBodyFont >= 15, 'Do not save space by shrinking body text');
  if (values.width > 820) {
    assert.ok(values.top <= 205, `Desktop focus chrome too tall: ${JSON.stringify(values)}`);
    assert.equal(values.sameHeaderRow, true, 'Tabs and shared controls should occupy one row');
  } else {
    assert.ok(values.top <= 330, `Phone focus chrome too tall: ${JSON.stringify(values)}`);
    assert.ok(values.readingHeight >= values.height * 0.6, 'At least 60% of phone viewport remains scrollable');
  }
  return values;
}
async function shot(page, label) { await page.screenshot({ path: path.join(out, label + '.png'), fullPage: true, animations: 'disabled' }); }

(async () => {
  const browser = await chromium.launch();
  let active;
  try {
    for (const width of [1440, 1024, 390, 320]) {
      const context = await browser.newContext({ viewport: { width, height: 928 }, colorScheme: 'light' });
      const errors = [];
      context.on('page', page => page.on('pageerror', error => errors.push(error.message)));
      await context.addInitScript(tabs => sessionStorage.setItem('seqlib-open-tabs-v1', JSON.stringify(tabs)), resources.map(({ path, title }) => ({ path, title })));
      await context.route('**/api/**', async route => {
        const url = new URL(route.request().url());
        let data, status = 200;
        if (url.pathname === '/api/tree') data = { items: [{ name: '序列库', path: '序列库', count: 6, children: [] }] };
        else if (url.pathname === '/api/resources') data = { items: resources, count: 6, total: 6, limit: 100, offset: 0, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
        else if (url.pathname.startsWith('/api/resources/')) { data = resources.find(item => item.path === decodeURIComponent(url.pathname.slice('/api/resources/'.length))); if (!data) { status = 404; data = { detail: 'missing' }; } }
        else { status = 404; data = { detail: 'unsupported test route' }; }
        await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(data) });
      });
      const page = await context.newPage(); active = page; page.setDefaultTimeout(12000);
      await page.goto(base + '/?open=' + encodeURIComponent(resources[0].path));
      await ready(page, resources[0]);
      // Open all resources through real navigation, independent of storage key.
      for (let i = 1; i < resources.length; i++) {
        if (width <= 820) await page.locator('.workbench-primary .realm-back').click();
        await page.locator(`[data-resource-index="${i}"]`).click();
        await ready(page, resources[i]);
      }
      assert.equal(await page.locator('.workbench-tab').count(), 6);
      await page.getByRole('button', { name: '并排对照', exact: true }).click();
      await ready(page, resources[4], '.workbench-reference');
      await measure(page, `${width}-compare`);
      await shot(page, `compact-${width}-compare`);
      const summary = page.locator('.workbench-reference-settings > summary');
      const before = await page.locator('.workbench-primary .realm-reader-scroll').evaluate(node => node.getBoundingClientRect().top);
      await summary.click();
      assert.equal(await page.locator('.workbench-reference-settings').getAttribute('open') !== null, true);
      const after = await page.locator('.workbench-primary .realm-reader-scroll').evaluate(node => node.getBoundingClientRect().top);
      assert.equal(after, before, 'Reference settings must overlay rather than push the reader down');
      const popover = await page.locator('.workbench-reference-popover').boundingBox();
      assert.ok(popover.x >= 0 && popover.x + popover.width <= width + 1, 'Reference settings must fit viewport');
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('.workbench-reference-settings').getAttribute('open'), null);
      assert.equal(await summary.evaluate(node => document.activeElement === node), true);
      await summary.click();
      await page.getByRole('combobox', { name: '选择参照档案' }).selectOption(resources[2].path);
      await ready(page, resources[2], '.workbench-reference');
      assert.equal(await page.locator('.workbench-reference-settings').getAttribute('open'), null);
      await summary.click();
      await page.locator('.workbench-primary .realm-document-heading').click();
      assert.equal(await page.locator('.workbench-reference-settings').getAttribute('open'), null);
      if (width <= 820) {
        await page.getByRole('button', { name: '参照档案', exact: true }).click();
        assert.equal(await page.locator('.workbench-reference').isVisible(), true);
        assert.equal(await page.locator('.workbench-primary').isVisible(), false);
        await measure(page, `${width}-reference`);
        await page.getByRole('button', { name: '主档案', exact: true }).click();
      }
      await page.getByRole('button', { name: '交换主 / 参照', exact: true }).click();
      await ready(page, resources[2]);
      await ready(page, resources[5], '.workbench-reference');
      if (width > 820) {
        await page.getByRole('button', { name: '退出并排', exact: true }).click();
        await page.getByRole('button', { name: '专注阅读', exact: true }).click();
        await measure(page, `${width}-solo-focus`);
        await shot(page, `compact-${width}-solo-focus`);
        await page.getByRole('button', { name: '退出专注', exact: true }).click();
        assert.equal(await page.locator('.realm-results').isVisible(), true);
      }
      assert.deepEqual(errors, []);
      await context.close();
    }
    fs.writeFileSync(path.join(out, 'compact-layout-metrics.json'), JSON.stringify(measurements, null, 2));
    console.log('Compact reader regressions passed.', JSON.stringify(measurements));
  } catch (error) {
    fs.writeFileSync(path.join(out, 'compact-failure.txt'), error.stack || String(error));
    fs.writeFileSync(path.join(out, 'compact-layout-metrics.json'), JSON.stringify(measurements, null, 2));
    if (active && !active.isClosed()) await shot(active, 'compact-failure').catch(() => {});
    throw error;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
