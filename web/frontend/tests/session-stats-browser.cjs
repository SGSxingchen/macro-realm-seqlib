/** E2E regression: discovered group months, async group sync and full session edits. */
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');

async function assertNoOverflow(page, width, stage) {
  const dimensions = await page.evaluate(() => ({
    scrollWidth: document.documentElement.scrollWidth,
    innerWidth,
  }));
  assert.ok(dimensions.scrollWidth <= dimensions.innerWidth + 1,
    `No horizontal overflow at ${width}px (${stage}): ${JSON.stringify(dimensions)}`);
}

(async () => {
  fs.mkdirSync(out, { recursive: true });
  const browser = await chromium.launch({ headless: true });
  try {
    for (const width of [1440, 390]) {
      const context = await browser.newContext({
        viewport: { width, height: 920 },
        colorScheme: 'light',
        timezoneId: 'Asia/Shanghai',
        serviceWorkers: 'block',
        ...(width <= 820 ? { isMobile: true, hasTouch: true } : {}),
      });
      try {
        // Keep November as the initial month, even when this runs in a later year.
        await context.addInitScript(() => {
          const RealDate = Date;
          window.Date = class extends RealDate {
            constructor(...args) { super(...(args.length ? args : ['2026-11-17T04:00:00Z'])); }
            static now() { return new RealDate('2026-11-17T04:00:00Z').getTime(); }
          };
        });
        const pageErrors = [];
        const unexpectedRequests = [];
        context.on('page', page => page.on('pageerror', error => pageErrors.push(error.message)));
        let pollCount = 0;
        let synced = false;
        let syncPayload = null;
        let patchPayload = null;
        let postSyncListReads = 0;
        let postPatchListReads = 0;
        const listMonths = [];
        let detail = {
          id: 101, month: '2026-12', title: '十二月群文件新团', duration_hours: 4,
          source_filename: '十二月结团记录.txt', model_name: 'mock-parser', confidence: 0.98,
          created_at: '2026-12-01T12:00:00Z', raw_payload: { source: 'deterministic mock' },
          kp: { id: 1, name: '原 KP', qq: '100000001' },
          participants: [
            { id: 1, name: '原 KP', qq: '100000001', is_host: true, role: 'KP', duration_hours: 4 },
            { id: 2, name: '原 PL', qq: '100000002', is_host: false, role: 'PL', duration_hours: 4 },
          ],
        };
        const job = completed => ({
          job_id: 'mock-group-job', month: '2026-12', status: completed ? 'completed' : 'running',
          total_count: 1, processed_count: completed ? 1 : 0,
          success_count: completed ? 1 : 0, failure_count: 0, skip_count: 0, concurrency: 2,
          items: completed ? [{ filename: detail.source_filename, success: true, session_id: 101 }] : [],
        });

        // Catch every API request: no route can fall through to a real backend.
        await context.route('**/api/**', async route => {
          const req = route.request();
          const url = new URL(req.url());
          const method = req.method();
          const month = url.searchParams.get('month');
          let status = 200;
          let body;
          if (url.pathname === '/api/tree' && method === 'GET') {
            body = { items: [] };
          } else if (url.pathname === '/api/resources' && method === 'GET') {
            body = { items: [], count: 0, total: 0, offset: 0, limit: 100, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
          } else if (url.pathname === '/api/auth/me' && method === 'GET') {
            body = { authenticated: false, auth_configured: true, admin_authenticated: false, is_moderator: false, user: null };
          } else if (url.pathname === '/api/session-stats/group-source' && method === 'GET') {
            body = {
              configured: true, can_manage: true, stale: false,
              auto_sync: { enabled: true, interval_seconds: 300 },
              items: ['2026-12', '2026-11'].map((itemMonth, index) => ({
                month: itemMonth, file_count: index ? 2 : 1, text_count: index ? 2 : 1,
                folders: [{ folder_id: `folder-${itemMonth}`, folder_name: `${index ? 11 : 12}月结团文件` }],
              })),
            };
          } else if (url.pathname === '/api/session-stats/overview' && method === 'GET') {
            const present = month === '2026-12' && synced;
            body = { session_count: present ? 1 : 0, participant_count: present ? detail.participants.length - 1 : 0,
              total_game_hours: present ? detail.duration_hours * (detail.participants.length - 1) : 0,
              total_host_hours: present ? detail.duration_hours : 0 };
          } else if (url.pathname === '/api/session-stats/players' && method === 'GET') {
            const items = month === '2026-12' && synced ? detail.participants.map(p => ({
              id: p.id, name: p.name, qq: p.qq, game_count: p.is_host ? 0 : 1,
              game_hours: p.is_host ? 0 : detail.duration_hours, reincarnation_count: 0,
              host_count: p.is_host ? 1 : 0, host_hours: p.is_host ? detail.duration_hours : 0,
            })) : [];
            body = { items, count: items.length, month };
          } else if (url.pathname === '/api/session-stats/sessions' && method === 'GET') {
            listMonths.push(month);
            if (synced) postSyncListReads++;
            if (patchPayload) postPatchListReads++;
            const items = month === '2026-12' && synced ? [{
              id: detail.id, title: detail.title, duration_hours: detail.duration_hours,
              source_filename: detail.source_filename, kp_name: detail.kp.name, kp_qq: detail.kp.qq,
              pl_count: detail.participants.length - 1, confidence: detail.confidence,
            }] : [];
            body = { items, count: items.length, month };
          } else if (url.pathname === '/api/session-stats/errors' && method === 'GET') {
            body = { items: [], count: 0, month };
          } else if (url.pathname === '/api/session-stats/sync-group' && method === 'POST') {
            syncPayload = req.postDataJSON();
            body = job(false);
          } else if (url.pathname === '/api/session-stats/import-jobs/mock-group-job' && method === 'GET') {
            pollCount++;
            synced = pollCount >= 2;
            body = job(synced);
          } else if (url.pathname === '/api/session-stats/sessions/101' && method === 'GET') {
            body = detail;
          } else if (url.pathname === '/api/session-stats/sessions/101' && method === 'PATCH') {
            patchPayload = req.postDataJSON();
            detail = {
              ...detail, title: patchPayload.title, duration_hours: patchPayload.duration_hours,
              kp: { id: 1, ...patchPayload.kp },
              participants: [
                { id: 1, ...patchPayload.kp, role: 'KP', is_host: true, duration_hours: patchPayload.duration_hours },
                ...patchPayload.players.map((p, i) => ({ id: i + 2, ...p, role: 'PL', is_host: false, duration_hours: patchPayload.duration_hours })),
              ],
            };
            body = { ok: true, session: detail };
          } else {
            unexpectedRequests.push(`${method} ${url.pathname}`);
            status = 404;
            body = { detail: 'not mocked' };
          }
          await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
        });

        const page = await context.newPage();
        page.setDefaultTimeout(12000);
        await page.goto(`${base}/?tab=stats`, { waitUntil: 'domcontentloaded' });
        await page.getByRole('heading', { name: '结团统计', exact: true }).waitFor();
        await page.getByText('群文件自动同步中', { exact: true }).waitFor();
        const november = page.getByRole('button', { name: '2026-11（群文件 2 份）', exact: true });
        const december = page.getByRole('button', { name: '2026-12（群文件 1 份）', exact: true });
        await november.waitFor();
        await december.waitFor();
        assert.equal(await page.locator('input[type="month"]').inputValue(), '2026-11');
        await december.click();
        await page.waitForFunction(() => document.querySelector('input[type="month"]')?.value === '2026-12' && !document.querySelector('[aria-label="刷新"]')?.disabled);
        assert.equal(await december.getAttribute('class'), 'on');
        await november.click();
        await page.waitForFunction(() => document.querySelector('input[type="month"]')?.value === '2026-11' && !document.querySelector('[aria-label="刷新"]')?.disabled);
        assert.equal(await november.getAttribute('class'), 'on');
        await december.click();
        await page.waitForFunction(() => document.querySelector('input[type="month"]')?.value === '2026-12' && !document.querySelector('[aria-label="刷新"]')?.disabled);
        assert.ok(listMonths.includes('2026-11') && listMonths.includes('2026-12'));
        await assertNoOverflow(page, width, 'month directories');

        await page.getByRole('button', { name: '⟳ 群文件拉取 / 导入', exact: true }).click();
        await page.getByRole('heading', { name: '自动从 QQ 群拉取结团文件 / 手动上传 TXT' }).waitFor();
        await assertNoOverflow(page, width, 'expanded import panel');
        await page.getByRole('button', { name: '拉取并解析 2026-12 群结团文件', exact: true }).click();
        await page.getByText('导入完成', { exact: true }).waitFor();
        await page.locator('.session-row').filter({ hasText: '十二月群文件新团' }).waitFor();
        assert.deepEqual(syncPayload, { month: '2026-12', concurrency: 2, force_retry_errors: false });
        assert.equal(pollCount, 2, 'Running job must be polled again until completed');
        assert.ok(postSyncListReads > 0, 'Completed sync must reload the session list');

        await page.locator('.session-row').filter({ hasText: '十二月群文件新团' }).getByRole('button', { name: '查看', exact: true }).click();
        const panel = page.locator('.stats-detail-panel');
        await panel.getByLabel('标题', { exact: true }).fill('十二月结团记录（已补录玩家）');
        await panel.getByLabel('时长（小时）', { exact: true }).fill('6.5');
        await panel.getByLabel('KP 昵称', { exact: true }).fill('补录 KP');
        await panel.getByLabel('KP QQ（选填）', { exact: true }).fill('200000001');
        await panel.getByRole('button', { name: '+ 添加玩家', exact: true }).click();
        assert.equal(await panel.getByPlaceholder('玩家 / 卡名', { exact: true }).count(), 2);
        await panel.getByPlaceholder('玩家 / 卡名', { exact: true }).last().fill('补录 PL');
        await panel.getByPlaceholder('QQ 号（可空）', { exact: true }).last().fill('200000003');
        await assertNoOverflow(page, width, 'session editor with added player');
        await panel.getByRole('button', { name: '保存全部修改', exact: true }).click();
        await page.locator('.session-row').filter({ hasText: '十二月结团记录（已补录玩家）' }).waitFor();
        await panel.getByRole('button', { name: '保存全部修改', exact: true }).waitFor();
        assert.deepEqual(patchPayload, {
          title: '十二月结团记录（已补录玩家）', duration_hours: 6.5,
          kp: { name: '补录 KP', qq: '200000001' },
          players: [{ name: '原 PL', qq: '100000002' }, { name: '补录 PL', qq: '200000003' }],
        });
        assert.ok(postPatchListReads > 0, 'Saving all edits must reload the session list');
        assert.equal(await panel.getByLabel('标题', { exact: true }).inputValue(), patchPayload.title);
        await panel.locator('.stats-participant-list b').filter({ hasText: '补录 PL' }).waitFor();
        await assertNoOverflow(page, width, 'saved session details');
        await page.screenshot({ path: path.join(out, `session-stats-${width}.png`), fullPage: true, animations: 'disabled' });
        assert.deepEqual(unexpectedRequests, [], 'All API requests must be explicitly mocked');
        assert.deepEqual(pageErrors, [], `No pageerror at ${width}px`);
        console.log(`Session stats browser regression passed at ${width}px (sync polling, KP/PL PATCH, zero overflow).`);
      } finally {
        await context.close();
      }
    }
  } finally {
    await browser.close();
  }
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
