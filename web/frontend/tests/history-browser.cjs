// 使用真实后端与正式档案核验历史记录，保留截图和原文复制证据。
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const before = process.env.HISTORY_BEFORE === '1';
const out = path.resolve(__dirname, '../test-results/history');
const files = ['序列库/特质改造/生化改造类/040】巨龙.txt', '序列库/特质改造/异化改造类/005】从者化《Fate》.txt', '序列库/技能表/其他及特殊/030】魂5技能表.txt', '序列库/技能表/其他及特殊/035】言灵《龙族》.txt'];
fs.mkdirSync(out, { recursive: true });
(async () => {
  const browser = await chromium.launch();
  try {
    for (const width of [390, 1440]) {
      for (const [i, file] of files.entries()) {
        const content = fs.readFileSync(path.resolve(__dirname, '../../..', file), 'utf8');
        const logs = content.split(/\r?\n/).filter(line => /^【(?:历史更新：|\d{2,4}[./年-])/.test(line));
        const context = await browser.newContext({ viewport: { width, height: 1000 }, permissions: ['clipboard-read', 'clipboard-write'], colorScheme: 'light' });
        const page = await context.newPage();
        page.setDefaultTimeout(15000);
        const errors = [];
        page.on('pageerror', e => errors.push(e.message));
        page.on('console', m => { if (m.type() === 'error') errors.push(m.text()); });
        await page.goto(`${base}/?open=${encodeURIComponent(file)}`);
        await page.locator('.workbench-primary h1').waitFor();
        await page.getByRole('button', { name: '复制全文', exact: true }).click();
        assert.equal(await page.evaluate(() => navigator.clipboard.readText()), content);
        if (!before) {
          const rows = page.locator('.workbench-primary .archive-section-changelog');
          if (i < 2) assert.equal(await rows.count(), logs.length, file);
          for (let n = 0; i < 2 && n < logs.length; n++) {
            const row = rows.nth(n);
            assert.equal(await row.locator('.changelog-date').count(), 1);
            await row.getByRole('button', { name: '复制原文', exact: true }).click();
            assert.equal((await page.evaluate(() => navigator.clipboard.readText())).trim(), logs[n]);
          }
        }
        if (!before && i >= 2) {
          const raw = i === 2
            ? content.slice(content.indexOf('【25.8.26：'), content.indexOf('\n【26.3.28：')).trim()
            : content.split(/\r?\n/).find(line => line.startsWith('【26.4.20：'));
          const row = page.locator('.workbench-primary .archive-section-changelog').filter({ hasText: i === 2 ? '25.8.26' : '26.4.20' });
          assert.equal(await row.count(), 1);
          await row.getByRole('button', { name: '复制原文', exact: true }).click();
          assert.equal((await page.evaluate(() => navigator.clipboard.readText())).trim(), raw);
          assert.ok((await row.textContent()).includes(i === 2 ? '涂抹油脂' : '无打折'));
        }
        await page.locator('.realm-reader-scroll').first().evaluate(n => { n.scrollTop = 500; });
        await page.screenshot({ path: path.join(out, `${before ? 'before' : 'after'}-${i}-${width}.png`) });
        await page.getByRole('button', { name: '文内查找', exact: true }).click();
        await page.getByRole('searchbox', { name: '文内精确查找' }).fill(i === 2 ? '25.8.26' : i === 3 ? '26.4.20' : '历史更新');
        await page.locator('mark[data-reader-match]').first().waitFor();
        assert.ok(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth));
        assert.deepEqual(errors, []);
        await page.screenshot({ path: path.join(out, `${before ? 'before' : 'after'}-${i}-${width}-find.png`) });
        await context.close();
      }
    }
    console.log('真实档案历史记录、全文/单条复制、查找浮窗和双端截图通过');
  } finally { await browser.close(); }
})().catch(e => { console.error(e); process.exitCode = 1; });
