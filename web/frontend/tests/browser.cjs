/** Integration tests use mock read-only APIs, not the production backend. */
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });
const first = '001】时空锚定\n\n（制作人：测试）\n保留序言与修改记录。\n\n[能力名称]：时空锚定（A级）\n[消耗能量]：10点；不可减免。\n[能力效果]：第一行效果。\n  缩进不能丢。\n\n注：离开范围立即失效。\n\n【非标准章节】\n不符合模板的自由规则 ▲ ◎\n未知限制也必须展示。\n\n' + Array.from({ length: 18 }, (_, i) => `[能力名称]：备用能力 ${i}\n[自定义代价]：保留未知字段 ${i}\n[能力效果]：逐字保留，不推断缺失规则。\n\n`).join('');
const second = '002】自由格式：跨界行者\n\n没有标准能力卡的长文。\n<script>window.__unsafe = true</script>\n注意：不可删除这一条限制。\n\n' + '无模板段落原样保留。\n'.repeat(80);
const resources = Array.from({ length: 620 }, (_, i) => ({
  path: `序列库/技能表/科技侧/${String(i + 1).padStart(3, '0')}】档案${i}.txt`,
  title: i === 0 ? '001】时空锚定' : i === 1 ? '002】自由格式：跨界行者' : `档案${i} · 足够长的标题用于检查两行排版与虚拟列表边界`,
  filename: `档案${i}.txt`, root: '序列库', category: '技能表/科技侧', top_kind: '技能表', side: '科技侧',
  authors: ['测试'], mtime: 1, size: 1000, snippet: '规则、条件和限制均来自原始文本；这里展示搜索片段，不替代完整正文。',
  content: i === 0 ? first : i === 1 ? second : `[能力名称]：能力${i}\n[能力效果]：第${i}份档案的完整效果。`, encoding: 'utf-8',
}));
const sleep = ms => new Promise(resolve => setTimeout(resolve, ms));
const failures = [];

async function setup(browser, viewport, colorScheme = 'dark') {
  const context = await browser.newContext({ viewport, colorScheme, permissions: ['clipboard-read', 'clipboard-write'] });
  context.on('page', page => { page.on('pageerror', error => failures.push(error.message)); page.on('dialog', async dialog => { failures.push(`Unexpected dialog: ${dialog.message()}`); await dialog.dismiss(); }); });
  const control = { delayed: false, detailError: false, queries: [] };
  await context.route('**/api/**', async route => {
    const url = new URL(route.request().url());
    let status = 200;
    let body;
    if (url.pathname === '/api/tree') body = { items: [{ name: '序列库', path: '序列库', count: 620, children: [{ name: '技能表', path: '序列库/技能表', count: 620, children: [{ name: '科技侧', path: '序列库/技能表/科技侧', count: 620, children: [] }] }] }] };
    else if (url.pathname === '/api/resources') {
      const q = url.searchParams.get('q') || '';
      control.queries.push(q);
      if (q === '__error__') { status = 503; body = { detail: 'mock list failure' }; }
      else {
        const data = resources.filter(item => (!q || `${item.title} ${item.content}`.includes(q)) && (!url.searchParams.get('category') || item.path.startsWith('序列库/' + url.searchParams.get('category'))));
        const offset = Number(url.searchParams.get('offset') || 0);
        const limit = Number(url.searchParams.get('limit') || 100);
        const items = data.slice(offset, offset + limit).map(({ content, ...item }) => item);
        body = { items, count: items.length, total: data.length, offset, limit, tokens: q ? [q] : [], facets: { kinds: [{ name: '技能表', count: data.length }], sides: [{ name: '科技侧', count: data.length }], authors: [{ name: '测试', count: data.length }] }, engine: { pinyin: false, opencc: false } };
      }
    } else if (url.pathname.startsWith('/api/resources/')) {
      const resourcePath = decodeURIComponent(url.pathname.slice('/api/resources/'.length));
      body = resources.find(item => item.path === resourcePath);
      if (control.detailError || !body) { status = 404; body = { detail: 'missing resource' }; }
      if (control.delayed) await sleep(resourcePath === resources[3].path ? 450 : 20);
    } else { status = 404; body = { detail: 'Unknown mock API' }; }
    try { await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) }); }
    catch { /* An aborted stale request is expected in the race test. */ }
  });
  const page = await context.newPage();
  page.setDefaultTimeout(12000);
  return { context, page, control };
}
async function titleIs(page, text) {
  await page.waitForFunction(expected => document.querySelector('.realm-document-heading h1')?.textContent === expected, text);
}
async function noOverflow(page) {
  const size = await page.evaluate(() => {
    const reader = document.querySelector('.realm-reader-scroll');
    const cards = Array.from(document.querySelectorAll('.realm-resource-card')).map(node => node.getBoundingClientRect()).filter(rect => rect.width > 0).sort((a, b) => a.top - b.top);
    return { documentWidth: document.documentElement.scrollWidth, viewport: innerWidth, readerWidth: reader?.clientWidth, readerScroll: reader?.scrollWidth, overlap: cards.some((rect, i) => i > 0 && cards[i - 1].bottom > rect.top + 1) };
  });
  assert.ok(size.documentWidth <= size.viewport + 1, JSON.stringify(size));
  if (size.readerWidth) assert.ok(size.readerScroll <= size.readerWidth + 1, JSON.stringify(size));
  assert.equal(size.overlap, false, 'Virtualized cards overlap');
}
async function shot(page, name) { await page.screenshot({ path: path.join(out, name + '.png'), fullPage: true, animations: 'disabled' }); }

(async () => {
  const browser = await chromium.launch({ headless: true });
  let activePage;
  try {
    const { context, page, control } = await setup(browser, { width: 1440, height: 1000 });
    activePage = page;
    await page.goto(base, { waitUntil: 'domcontentloaded' });
    await page.locator('[data-resource-index="0"]').waitFor();
    await noOverflow(page);
    await shot(page, 'desktop-dark-overview');
    await page.locator('[data-resource-index="0"]').click();
    await titleIs(page, resources[0].title);
    await page.getByRole('button', { name: '复制全文', exact: true }).click();
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), first);
    const entry = page.locator('.archive-entry').first();
    await entry.getByRole('button', { name: '复制原文', exact: true }).click();
    assert.equal(await page.evaluate(() => navigator.clipboard.readText()), first.slice(first.indexOf('[能力名称]'), first.indexOf('【非标准章节】')));
    await entry.getByRole('button', { name: /分享定位/ }).click();
    const link = await page.evaluate(() => navigator.clipboard.readText());
    assert.ok(new URL(link).hash.startsWith('#entry-'));
    const shared = await context.newPage();
    await shared.goto(link);
    await titleIs(shared, resources[0].title);
    await shared.waitForFunction(() => {
      const section = document.getElementById(location.hash.slice(1));
      const reader = document.querySelector('.realm-reader-scroll');
      return section && reader && Math.abs(section.getBoundingClientRect().top - reader.getBoundingClientRect().top - 20) < 4;
    });
    await shared.close();
    await page.getByRole('button', { name: '文内查找', exact: true }).click();
    await page.getByRole('searchbox', { name: '文内精确查找' }).fill('不可减免');
    await page.waitForSelector('[data-reader-match][data-current="true"]');
    assert.equal(await page.locator('[data-reader-match]').count(), 1);
    await page.getByRole('button', { name: '关闭文内查找' }).click();
    await page.getByRole('button', { name: '☆ 收藏', exact: true }).click();
    assert.ok((await page.evaluate(() => localStorage.getItem('seqlib-saved-v1'))).includes(resources[0].path));
    await page.getByRole('button', { name: '原文', exact: true }).click();
    assert.equal(await page.locator('.realm-document-content').textContent(), first);
    await page.getByRole('button', { name: '自适应', exact: true }).click();
    await noOverflow(page);
    await shot(page, 'desktop-dark-reader');
    await page.evaluate(() => { document.querySelector('.realm-reader-scroll').scrollTop = 0; });
    await shot(page, 'desktop-dark-reader-top');
    await page.getByRole('button', { name: '切换主题' }).click();
    await page.waitForFunction(() => document.documentElement.dataset.theme === 'light');
    await shot(page, 'desktop-light-reader');

    control.delayed = true;
    // Use uncached resources so both competing requests really execute.
    await page.locator('[data-resource-index="2"]').click();
    await titleIs(page, resources[2].title);
    await page.locator('[data-resource-index="3"]').click();
    await page.locator('[data-resource-index="4"]').click();
    await titleIs(page, resources[4].title);
    await sleep(550);
    assert.equal(await page.locator('.realm-document-heading h1').textContent(), resources[4].title);
    await page.goBack();
    await titleIs(page, resources[3].title);
    await page.goForward();
    await titleIs(page, resources[4].title);
    await page.locator('[data-resource-index="1"]').click();
    await titleIs(page, resources[1].title);
    await page.reload();
    await titleIs(page, resources[1].title);
    assert.equal(await page.evaluate(() => window.__unsafe), undefined);
    assert.ok((await page.locator('.realm-document-content').textContent()).includes('<script>'));
    control.delayed = false;

    // Reach a real row beyond the old 500-item limit.
    for (const total of [200, 300, 400, 500, 600, 620]) {
      await page.getByRole('button', { name: /加载更多/ }).click();
      await page.waitForFunction(n => { const button = document.querySelector('.realm-load-more'); return n === 620 ? !button : button?.textContent.includes(`已显示 ${n} / 620`); }, total);
    }
    await page.locator('[data-resource-index]').first().focus();
    await page.keyboard.press('End');
    await page.locator('[data-resource-index="619"]').click();
    await titleIs(page, resources[619].title);
    await noOverflow(page);

    const search = page.getByRole('searchbox', { name: '搜索档案标题或正文' });
    await search.fill('__error__');
    await page.getByRole('alert').filter({ hasText: '检索失败' }).waitFor();
    assert.equal(await page.locator('.realm-no-results').count(), 0);
    await search.dispatchEvent('compositionstart');
    await search.fill('shikong');
    await sleep(300);
    assert.equal(control.queries.includes('shikong'), false, 'IME intermediate text must not trigger a search');
    await search.fill('时空锚定');
    await search.dispatchEvent('compositionend');
    await page.waitForFunction(() => document.querySelector('.realm-results-meta')?.textContent.includes('找到 1 份'));
    await page.locator('[data-resource-index="0"]').click();
    await titleIs(page, resources[0].title);
    // Rejected clipboard promises must be surfaced instead of reporting success.
    await page.evaluate(() => { navigator.clipboard.writeText = async () => { throw new Error('复制权限被拒绝'); }; });
    await page.getByRole('button', { name: '复制全文', exact: true }).click();
    await page.locator('.realm-toast').filter({ hasText: '复制权限被拒绝' }).waitFor();
    const missing = await context.newPage();
    await missing.goto(base + '/?open=' + encodeURIComponent('序列库/不存在.txt'));
    await missing.getByRole('heading', { name: '暂时无法打开这份档案' }).waitFor();
    assert.equal(await missing.locator('.realm-document-heading').count(), 0);
    await missing.close();
    await context.close();

    for (const width of [390, 320, 1024]) {
      const mobile = await setup(browser, { width, height: width === 1024 ? 768 : 844 });
      activePage = mobile.page;
      await mobile.page.goto(base);
      await mobile.page.locator('[data-resource-index="0"]').waitFor();
      await noOverflow(mobile.page);
      await shot(mobile.page, `${width}-results`);
      await mobile.page.getByRole('button', { name: /^筛选/ }).click();
      await mobile.page.waitForSelector('dialog[open]');
      await mobile.page.keyboard.press('Tab');
      assert.equal(await mobile.page.evaluate(() => !!document.activeElement.closest('dialog')), true);
      await mobile.page.keyboard.press('Escape');
      await mobile.page.waitForSelector('dialog[open]', { state: 'hidden' });
      assert.equal(await mobile.page.evaluate(() => document.activeElement.classList.contains('realm-filter-trigger')), true);
      await mobile.page.getByRole('searchbox', { name: '搜索档案标题或正文' }).fill('时空锚定');
      await mobile.page.waitForFunction(() => document.querySelector('.realm-results-meta')?.textContent.includes('找到 1 份'));
      await mobile.page.locator('[data-resource-index="0"]').click();
      await titleIs(mobile.page, resources[0].title);
      await noOverflow(mobile.page);
      await shot(mobile.page, `${width}-reader`);
      if (width <= 820) {
        assert.equal(await mobile.page.locator('.realm-results').isVisible(), false);
        await mobile.page.locator('.realm-reader-toolbar .realm-back').click();
        await mobile.page.waitForSelector('.realm-results', { state: 'visible' });
        assert.equal(await mobile.page.locator('#realm-query').inputValue(), '时空锚定');
        await mobile.page.goForward();
        await titleIs(mobile.page, resources[0].title);
        await mobile.page.keyboard.press('Control+k');
        await mobile.page.waitForSelector('.realm-results', { state: 'visible' });
        assert.equal(await mobile.page.evaluate(() => document.activeElement.id), 'realm-query');
      }
      await mobile.context.close();
    }
    assert.deepEqual(failures, [], 'Browser runtime errors');
    fs.writeFileSync(path.join(out, 'summary.txt'), 'PASS: exact source copy, raw rendering, stable section links, text safety, find, bookmarks, browser history, pagination >500, errors, clipboard rejection, responsive layouts and modal focus. Browser tests use deterministic mock APIs.\n');
    console.log('All archive browser regressions passed. Screenshots: test-results/');
  } catch (error) {
    fs.writeFileSync(path.join(out, 'failure.txt'), error.stack || String(error));
    if (activePage && !activePage.isClosed()) await shot(activePage, 'failure').catch(() => {});
    throw error;
  } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
