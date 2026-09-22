// Deterministic UI tests. No production data, credentials or write APIs.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });
const resources = ['时空锚定', '跨界旅人', '超长标题用于验证手机端打开多个档案后仍可关闭和切换'].map((title, i) => ({
  path: `序列库/技能表/科技侧/00${i + 1}】${title}.txt`, title, filename: title + '.txt', root: '序列库',
  category: '技能表/科技侧', top_kind: '技能表', side: '科技侧', size: 8000, mtime: 1, encoding: 'utf-8', authors: ['测试作者'],
  snippet: '完整规则 · 所有条件和限制保留。',
  content: `${title}\n\n这段开头保留原有顺序。\n\n` + Array.from({ length: 30 }, (_, n) => `[能力名称]：${title} · ${n + 1}\n[消耗能量]：10 点\n[施展时间]：一个标准动作\n[能力效果]：在指定范围内建立稳定锚点。所有目标、持续时间与限制条件必须完整保留，不得因屏幕变窄而省略。\n[特殊限制]：${n === 0 ? 'Unbroken_identifier_'.repeat(12) : '此效果不可叠加。'}\n\n`).join(''),
}));
const scroll = page => page.locator('.workbench-primary .realm-reader-scroll');
async function opened(page, index) {
  await page.waitForFunction(title => document.querySelector('.workbench-primary h1')?.textContent === title, resources[index].title);
}
async function screenshot(page, name) {
  await page.screenshot({ path: path.join(out, `reading-${name}.png`), fullPage: true, animations: 'disabled' });
}
async function noOverflow(page) {
  assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), 'No horizontal page overflow');
  assert.ok(await page.locator('.realm-reader-scroll').evaluateAll(nodes => nodes.filter(n => n.clientWidth).every(n => n.scrollWidth <= n.clientWidth + 1)), 'No horizontal document overflow');
}
async function setup(browser, width, height = 844, storedPanes) {
  const context = await browser.newContext({ viewport: { width, height }, colorScheme: 'dark', ...(width <= 820 ? { isMobile: true, hasTouch: true } : {}) });
  const errors = [];
  context.on('page', page => page.on('pageerror', error => errors.push(error.message)));
  if (storedPanes) await context.addInitScript(value => localStorage.setItem('seqlib-panes-v1', JSON.stringify(value)), storedPanes);
  await context.route('**/api/**', async route => {
    const url = new URL(route.request().url());
    let data, status = 200;
    if (url.pathname === '/api/tree') data = { items: [{ name: '序列库', path: '序列库', count: 3, children: [] }] };
    else if (url.pathname === '/api/resources') {
      const items = resources.filter(item => !url.searchParams.get('q') || item.title.includes(url.searchParams.get('q')));
      data = { items, count: items.length, total: items.length, offset: 0, limit: 100, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
    } else if (url.pathname.startsWith('/api/resources/')) {
      data = resources.find(item => item.path === decodeURIComponent(url.pathname.slice('/api/resources/'.length)));
      if (!data) { status = 404; data = { detail: 'not found' }; }
    } else if (url.pathname === '/api/git/changes') data = { from_ref: 'v-test', to: 'working-tree/latest', readable: { added: [], modified: [], deleted: [], renamed: [] } };
    else { status = 404; data = { detail: 'unsupported fixture' }; }
    await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(data) }).catch(() => {});
  });
  const page = await context.newPage();
  page.setDefaultTimeout(12000);
  await page.goto(base);
  await page.locator('[data-resource-index="0"]').waitFor();
  return { page, context, errors };
}

(async () => {
  const browser = await chromium.launch();
  const measurements = [];
  let active;
  try {
    for (const width of [320, 390, 768, 1024, 1440]) {
      const { page, context, errors } = await setup(browser, width); active = page;
      assert.ok(await page.getByRole('button', { name: '专注阅读', exact: true }).isDisabled());
      assert.equal(Math.round((await page.locator('.realm-resource-row').first().boundingBox()).height), 104);
      await noOverflow(page);
      await screenshot(page, `${width}-results`);
      await page.locator('[data-resource-index="0"]').click(); await opened(page, 0);
      await noOverflow(page);
      const normal = await scroll(page).boundingBox();
      if (width <= 820) {
        assert.ok(normal.width >= width - 2, 'The reader fills the phone width');
        assert.ok(normal.height >= 844 * .72, 'Normal mode must not require focus to be usable');
        const textWidth = await page.locator('.archive-field.is-rule .archive-field-value').first().evaluate(n => n.getBoundingClientRect().width);
        assert.ok(textWidth >= width - 36, 'Rules should not lose width to nested cards');
        assert.ok(await page.locator('.realm-reader-toolbar').evaluate(n => n.getBoundingClientRect().height <= 50), 'One-row mobile reading toolbar');
        // Navigation opens without moving the reader or consuming a permanent row.
        await page.getByRole('button', { name: '导航', exact: true }).click();
        await page.getByRole('button', { name: '切换主题' }).click();
        await page.waitForFunction(() => document.documentElement.dataset.theme === 'light');
        assert.equal(await page.getByRole('button', { name: '导航', exact: true }).getAttribute('aria-expanded'), 'false');
      } else if (width === 1440) {
        assert.ok(normal.width >= width * .68, 'Normal desktop reading receives most of the workspace');
      }
      await screenshot(page, `${width}-normal`);
      // Explicit focus keeps the same mounted document and reading position.
      const handle = await scroll(page).elementHandle();
      await scroll(page).evaluate(n => { n.scrollTop = 500; });
      await page.getByRole('button', { name: '专注阅读', exact: true }).click();
      await page.locator('.realm-focus-reading').waitFor();
      const focused = await scroll(page).boundingBox();
      assert.ok(focused.height >= normal.height + 70, 'Focus actually frees vertical space');
      assert.equal(await page.locator('.realm-topbar').isVisible(), false);
      assert.ok(await handle.evaluate(n => n === document.querySelector('.workbench-primary .realm-reader-scroll')));
      if (width <= 820) assert.ok(Math.abs(await scroll(page).evaluate(n => n.scrollTop) - 500) < 3);
      await noOverflow(page);
      await screenshot(page, `${width}-focus`);
      measurements.push({ width, height: 844, normal, focused });
      if (width <= 820) {
        const settings = page.getByRole('button', { name: '阅读设置', exact: true });
        await settings.click();
        await page.getByRole('button', { name: '原文', exact: true }).click();
        assert.equal(await page.locator('.realm-document-content').textContent(), resources[0].content);
        await page.getByRole('button', { name: '☆ 收藏', exact: true }).click();
        assert.ok((await page.evaluate(() => localStorage.getItem('seqlib-saved-v1'))).includes(resources[0].path));
        await page.getByRole('button', { name: '自适应', exact: true }).click();
        await page.keyboard.press('Escape');
        assert.equal(await settings.getAttribute('aria-expanded'), 'false');
        assert.ok(await page.locator('.realm-focus-reading').count(), 'Escape in settings closes only the disclosure');
        assert.ok(await settings.evaluate(n => document.activeElement === n));
        await settings.click();
        // The title's centre is covered by the menu on 320px screens.
        // Click its left edge to exercise a genuine outside interaction.
        await page.locator('.realm-document-heading h1').click({ position: { x: 2, y: 2 } });
        assert.equal(await settings.getAttribute('aria-expanded'), 'false', 'Outside click dismisses settings');
      }
      // The TOC dismisses locally and leaves focus mode active.
      await page.locator('.workbench-primary .realm-toc summary').click();
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('.workbench-primary .realm-toc').getAttribute('open'), null);
      assert.ok(await page.locator('.realm-focus-reading').count());
      await page.getByRole('button', { name: '退出专注', exact: true }).focus();
      await page.keyboard.press('Escape');
      assert.equal(await page.locator('.realm-focus-reading').count(), 0);
      assert.ok(await page.locator('.realm-topbar').isVisible());
      await page.getByRole('button', { name: '专注阅读', exact: true }).click();
      if (width <= 820) {
        await page.getByRole('button', { name: '索引 / 收藏', exact: true }).click();
        await page.locator('dialog[open]').waitFor();
        await page.getByRole('button', { name: /^查看结果/ }).click();
      } else await page.keyboard.press('Control+k');
      await page.waitForFunction(() => document.activeElement.id === 'realm-query');
      assert.ok(await page.locator('.realm-results').isVisible());
      await page.locator('[data-resource-index="2"]').click(); await opened(page, 2);
      await noOverflow(page);
      await page.getByRole('button', { name: '并排对照', exact: true }).click();
      await page.locator('.workbench-reference h1').waitFor({ state: 'attached' });
      await page.getByRole('button', { name: '专注阅读', exact: true }).click();
      assert.ok(await page.getByRole('button', { name: '退出并排', exact: true }).isVisible(), 'Comparison remains escapable in focus');
      await page.getByRole('button', { name: '退出并排', exact: true }).click();
      await page.getByRole('button', { name: '退出专注', exact: true }).click();
      await page.getByRole('button', { name: `关闭档案：${resources[2].title}`, exact: true }).click(); await opened(page, 0);
      await page.getByRole('button', { name: `关闭档案：${resources[0].title}`, exact: true }).click();
      assert.ok(await page.locator('.realm-results').isVisible(), 'Closing the final tab returns to usable results');
      assert.deepEqual(errors, []);
      await context.close();
    }
    // Restored collapsed panes must never produce a blank no-document view.
    const restored = await setup(browser, 1440, 844, { index: false, results: false }); active = restored.page;
    assert.ok(await active.locator('.realm-results').isVisible());
    await active.locator('[data-resource-index="0"]').click(); await opened(active, 0);
    for (const width of [1181, 1180, 821, 820, 390, 844]) {
      await active.setViewportSize({ width, height: 600 });
      await noOverflow(active);
      assert.ok(await scroll(active).isVisible());
      assert.ok((await scroll(active).boundingBox()).height >= 300, 'Resizing keeps room for the reader');
    }
    await restored.context.close();
    fs.writeFileSync(path.join(out, 'reading-layout-measurements.json'), JSON.stringify(measurements, null, 2));
    fs.writeFileSync(path.join(out, 'reading-layout-summary.txt'), 'PASS: 320/390/768/1024/1440px viewports; normal/focus space budgets; borderless full-width mobile text; compact virtual rows; mobile navigation/settings; raw text preservation; bookmark; local Escape/outside dismissal; focus restoration; search/filter escape routes; long tabs/comparison; closing final tab; restored hidden panes; responsive boundary resizing. Chromium with deterministic mock API, not real iOS hardware.\n');
    console.log('Reading layout browser regressions passed.');
  } catch (error) {
    fs.writeFileSync(path.join(out, 'reading-layout-failure.txt'), error.stack || String(error));
    if (active && !active.isClosed()) await screenshot(active, 'failure').catch(() => {});
    throw error;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
