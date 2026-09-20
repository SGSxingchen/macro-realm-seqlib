// Geometry and interaction regressions for reading density. Fixtures are
// shared with the existing workbench suite; no production API is contacted.
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const { setup, resources, titleIs, noOverflow, primary, reference } = require('./workbench-browser.cjs');
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });
for (const [index, name] of ['通用仪式', '仪式学技能组（注解）', '跨界技能与特殊规则：很长的档案名也不能挤出新一层工具栏'].entries()) {
  resources.push({ ...resources[index], path: `序列库/技能表/科技侧/00${index + 4}】${name}.txt`, title: name });
}
const metrics = [];
const button = (page, name) => page.getByRole('button', { name, exact: true });
const tab = (page, index) => page.locator('.workbench-tab-title').filter({ hasText: resources[index].title });
const snapshot = async (page, name) => page.screenshot({ path: path.join(out, `focus-${name}.png`), fullPage: true, animations: 'disabled' });

async function measure(page, mode) {
  const result = await page.evaluate(() => ({
    width: innerWidth, height: innerHeight,
    chrome: document.querySelector('.workbench-chrome').getBoundingClientRect().height,
    panes: Array.from(document.querySelectorAll('.realm-reader-scroll')).filter(node => node.clientHeight && node.clientWidth).map(node => ({ top: node.getBoundingClientRect().top, height: node.clientHeight, width: node.clientWidth })),
    // Ability names intentionally remain larger; measure actual rule text.
    bodyFont: getComputedStyle(document.querySelector('.archive-field:not(.is-name) .archive-field-value') || document.querySelector('.archive-prose')).fontSize,
    visibleGeneratedHeadings: Array.from(document.querySelectorAll('.realm-document-heading')).filter(node => node.getClientRects().length).length,
  }));
  metrics.push({ mode, ...result });
  assert.equal(result.visibleGeneratedHeadings, 0, 'Compact reading must not repeat the decorative document header');
  assert.equal(result.bodyFont, '15px', 'Do not buy vertical space by shrinking rule text');
  const ceiling = result.width <= 820 ? 170 : 120;
  assert.ok(result.panes.length > 0);
  for (const pane of result.panes) assert.ok(pane.top <= ceiling, `Fixed chrome exceeded ${ceiling}px: ${JSON.stringify(result)}`);
  await noOverflow(page);
  return result;
}
async function currentAnchor(page) {
  return primary(page).locator('.realm-reader-scroll').evaluate(node => {
    const top = node.getBoundingClientRect().top;
    const section = Array.from(node.querySelectorAll('.archive-section')).find(item => item.getBoundingClientRect().bottom > top + 20);
    return { id: section.id, offset: section.getBoundingClientRect().top - top };
  });
}
async function sameAnchor(page, anchor) {
  const offset = await primary(page).locator('.realm-reader-scroll').evaluate((node, target) => node.querySelector('#' + CSS.escape(target.id)).getBoundingClientRect().top - node.getBoundingClientRect().top, anchor);
  assert.ok(Math.abs(offset - anchor.offset) <= 5, `Visible source jumped on density toggle: ${anchor.offset} -> ${offset}`);
}
async function fits(page, dialog, width, height) {
  const rect = await dialog.boundingBox();
  assert.ok(rect.x >= 0 && rect.y >= 0 && rect.x + rect.width <= width + 1 && rect.y + rect.height <= height + 1, 'Popover must fit within the viewport');
}

(async () => {
  const browser = await chromium.launch();
  let active;
  try {
    for (const [width, height] of [[1611, 928], [1440, 900], [1024, 768], [390, 844], [320, 640]]) {
      const run = await setup(browser, width);
      const { page, context } = run; active = page;
      await page.setViewportSize({ width, height });
      await context.grantPermissions(['clipboard-read', 'clipboard-write']);
      for (let index = 0; index < resources.length; index++) {
        if (index && width <= 820) {
          await page.keyboard.press('Control+k');
          await page.waitForFunction(() => document.activeElement.id === 'realm-query');
        }
        await page.locator(`[data-resource-index="${index}"]`).click();
        await titleIs(page, index);
      }
      assert.equal(await page.locator('.workbench-tab').count(), 6);
      if (width > 820) {
        if (width > 1180) await button(page, '收起索引').click();
        await primary(page).locator('.realm-reader-scroll').evaluate(node => { node.scrollTop = 900; });
        await page.waitForTimeout(50); // allow the scroll event to capture the source anchor
        const before = await currentAnchor(page);
        await button(page, '专注阅读').click();
        await sameAnchor(page, before);
        assert.equal(await page.locator('.realm-topbar').isVisible(), false);
        await button(page, '退出专注').click();
        await sameAnchor(page, before);
        assert.equal(await page.locator('.realm-results').isVisible(), true);
        if (width > 1180) assert.equal(await page.locator('.realm-navigation').isVisible(), false);
        await button(page, '专注阅读').click();
      }
      await measure(page, 'single');
      await primary(page).locator('.realm-reader-scroll').evaluate(node => { node.scrollTop = 0; });
      await snapshot(page, `${width}-single`);
      await button(page, '并排对照').click();
      await titleIs(page, 4, '.workbench-reference');
      const comparison = await measure(page, 'compare');
      assert.equal(await page.locator('.workbench-controls').count(), 1);
      await button(page, '对照设置').click();
      const config = page.getByRole('dialog', { name: '对照设置', exact: true });
      await config.waitFor();
      const during = await primary(page).locator('.realm-reader-scroll').boundingBox();
      assert.ok(Math.abs(during.y - comparison.panes[0].top) < 1, 'Settings cannot push the reader down');
      await fits(page, config, width, height);
      await page.keyboard.press('Tab');
      assert.equal(await page.evaluate(() => !!document.activeElement.closest('dialog[open]')), true);
      await page.keyboard.press('Escape');
      assert.equal(await button(page, '对照设置').evaluate(node => node === document.activeElement), true);
      await button(page, '对照设置').click();
      await page.getByRole('combobox', { name: '选择参照档案' }).selectOption(resources[0].path);
      await button(page, '交换主 / 参照').click();
      await titleIs(page, 0); await titleIs(page, 5, '.workbench-reference');
      assert.equal(await config.isVisible(), false);
      await tab(page, 2).click(); await titleIs(page, 2);
      await titleIs(page, 5, '.workbench-reference');
      if (width <= 820) {
        await button(page, '参照档案').click();
        assert.equal(await reference(page).isVisible(), true);
        await measure(page, 'reference-mobile');
        await button(page, '主档案').click();
      }
      // Directory must also escape the narrow pane's overflow clipping.
      await primary(page).getByRole('button', { name: '文档目录', exact: true }).click();
      const directory = page.getByRole('dialog', { name: '文档目录', exact: true });
      await directory.waitFor();
      await fits(page, directory, width, height);
      await directory.locator('.reader-directory button').nth(4).click();
      assert.equal(await directory.isVisible(), false);
      assert.ok(await primary(page).locator('.realm-reader-scroll').evaluate(node => node.scrollTop > 0));
      await primary(page).getByRole('button', { name: '阅读工具', exact: true }).click();
      const tools = page.getByRole('dialog', { name: '阅读工具', exact: true });
      await tools.getByRole('button', { name: '原文', exact: true }).click();
      await page.keyboard.press('Escape');
      assert.equal(await primary(page).locator('.realm-document-content').textContent(), resources[2].content);
      await primary(page).getByRole('button', { name: '阅读工具', exact: true }).click();
      await tools.getByRole('button', { name: '复制全文', exact: true }).click();
      assert.equal(await page.evaluate(() => navigator.clipboard.readText()), resources[2].content);
      await primary(page).getByRole('button', { name: '阅读工具', exact: true }).click();
      await tools.getByRole('button', { name: '自适应', exact: true }).click();
      await page.keyboard.press('Escape');
      await button(page, '阅读布局').click();
      const layout = page.getByRole('dialog', { name: '阅读布局', exact: true });
      await layout.getByRole('button', { name: '切换主题' }).click();
      await page.keyboard.press('Escape');
      await page.waitForFunction(() => document.documentElement.dataset.theme === 'light');
      await snapshot(page, `${width}-compare-light`);
      await measure(page, 'compare-light');
      if (width > 820) {
        await button(page, '退出专注').click();
        assert.equal(await page.locator('.realm-topbar').isVisible(), true);
        const theme = page.locator('.realm-topbar').getByRole('button', { name: '切换主题' });
        assert.equal(await theme.getAttribute('title'), '切到暗色', 'Compact and header theme controls must stay synchronized');
        await theme.click();
        await button(page, '专注阅读').click();
      } else {
        await button(page, '阅读布局').click();
        await layout.getByRole('button', { name: '切换主题' }).click();
        await page.keyboard.press('Escape');
      }
      await page.waitForFunction(() => document.documentElement.dataset.theme === 'dark');
      await snapshot(page, `${width}-compare-dark`);
      await page.keyboard.press('Control+k');
      await page.waitForFunction(() => document.activeElement.id === 'realm-query');
      assert.equal(await page.locator('.realm-topbar').isVisible(), true);
      assert.equal(await page.locator('.realm-results').isVisible(), true);
      assert.deepEqual(run.errors, []);
      await context.close();
    }
    fs.writeFileSync(path.join(out, 'focus-summary.txt'), 'PASS: compact fixed-chrome geometry, six/long tabs, unchanged rule text size, source anchor restoration, top-layer directory/settings without layout shift, focus return, raw/copy tools, reference selection/swap, responsive pane switching, synchronized theme and search recovery. Fixtures use simulated read-only APIs.\n');
    console.log('Focused-reading browser regressions passed.');
  } catch (error) {
    fs.writeFileSync(path.join(out, 'focus-failure.txt'), error.stack || String(error));
    if (active && !active.isClosed()) await snapshot(active, 'failure').catch(() => {});
    throw error;
  } finally {
    fs.writeFileSync(path.join(out, 'focus-metrics.json'), JSON.stringify(metrics, null, 2));
    await browser.close();
  }
})().catch(error => { console.error(error); process.exitCode = 1; });
