/** E2E regression: QQ private code login, independent home message board, resource anchor comments, likes & moderation. */
const { chromium } = require(process.env.PLAYWRIGHT_MODULE || 'playwright');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const base = process.env.TEST_BASE_URL || 'http://127.0.0.1:4173';
const out = path.resolve(__dirname, '../test-results');
fs.mkdirSync(out, { recursive: true });

const sampleDoc = '001】时空锚定\n\n（制作人：测试）\n\n[能力名称]：时空锚定（A级）\n[消耗能量]：10点\n[能力效果]：锁定一片空间。\n\n[能力名称]：相位折跃（B级）\n[消耗能量]：5点\n[能力效果]：短距瞬移。\n';
const resources = [{
  path: '序列库/技能表/科技侧/001】时空锚定.txt',
  title: '001】时空锚定',
  filename: '001】时空锚定.txt',
  root: '序列库',
  category: '技能表/科技侧',
  top_kind: '技能表',
  side: '科技侧',
  authors: ['测试'],
  mtime: 1,
  size: sampleDoc.length,
  snippet: '锁定一片空间。',
  content: sampleDoc,
  encoding: 'utf-8',
}];

(async () => {
  const browser = await chromium.launch({ headless: true });
  try {
    for (const width of [1440, 390]) {
      const context = await browser.newContext({
        viewport: { width, height: 900 },
        colorScheme: 'dark',
        ...(width <= 820 ? { isMobile: true, hasTouch: true } : {}),
      });
      const pageErrors = [];
      context.on('page', p => p.on('pageerror', e => pageErrors.push(e.message)));

      let loggedIn = false;
      let challengeVerified = false;
      const boardPosts = [
        {
          id: 1,
          scope: 'board',
          resource_path: '',
          resource_title: '',
          anchor_id: '',
          anchor_label: '',
          parent_id: null,
          reply_to_id: null,
          reply_to_nickname: '',
          author: { qq_masked: '134****7221', qq: '1348697221', nickname: '无语', avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', badge: '维护者' },
          content: '欢迎来到界域留言板！',
          is_pinned: true,
          like_count: 2,
          liked_by_me: false,
          can_delete: true,
          can_pin: true,
          status: 'active',
          created_at: new Date().toISOString(),
          created_ts: Math.floor(Date.now() / 1000) - 120,
          replies: [],
          reply_count: 0,
        },
      ];
      const resourceComments = [];
      let nextId = 2;

      await context.route('**/api/**', async route => {
        const req = route.request();
        const url = new URL(req.url());
        let status = 200;
        let body = {};

        const currentUser = loggedIn
          ? {
              qq: '1348697221',
              qq_masked: '134****7221',
              nickname: '调查员星尘',
              avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
              card_count: 2,
              badge: '维护者',
              is_banned: false,
            }
          : null;

        if (url.pathname === '/api/tree') {
          body = { items: [{ name: '序列库', path: '序列库', count: 1, children: [{ name: '技能表', path: '序列库/技能表', count: 1, children: [] }] }] };
        } else if (url.pathname === '/api/resources') {
          body = { items: resources, count: 1, total: 1, offset: 0, limit: 100, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
        } else if (url.pathname.startsWith('/api/resources/')) {
          body = resources[0];
        } else if (url.pathname === '/api/auth/me') {
          body = { authenticated: loggedIn, auth_configured: true, bot_qq: '3239195029', group_id: 672204828, admin_authenticated: false, is_moderator: loggedIn, user: currentUser };
        } else if (url.pathname === '/api/auth/qq/challenges' && req.method() === 'POST') {
          // Simulates user immediately sending private message right after challenge creation
          setTimeout(() => { challengeVerified = true; }, 300);
          body = {
            ok: true,
            qq: '1348697221',
            group_id: 672204828,
            bot_qq: '3239195029',
            verification_code: 'TRPG-TESTCODE8888',
            message_to_send: 'TRPG-TESTCODE8888',
            expires_at: new Date(Date.now() + 300000).toISOString(),
            status: 'pending',
          };
        } else if (url.pathname === '/api/auth/qq/challenges/current') {
          body = {
            qq: '1348697221',
            bot_qq: '3239195029',
            status: challengeVerified ? 'verified' : 'pending',
            expires_at: new Date(Date.now() + 300000).toISOString(),
          };
        } else if (url.pathname === '/api/auth/qq/challenges/consume' && req.method() === 'POST') {
          loggedIn = true;
          body = {
            ok: true,
            authenticated: true,
            auth_configured: true,
            bot_qq: '3239195029',
            group_id: 672204828,
            admin_authenticated: false,
            is_moderator: true,
            user: {
              qq: '1348697221',
              qq_masked: '134****7221',
              nickname: '调查员星尘',
              avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
              card_count: 2,
              badge: '维护者',
              is_banned: false,
            },
          };
        } else if (url.pathname === '/api/community/home') {
          body = {
            board: { items: boardPosts, count: boardPosts.length, total_threads: boardPosts.length, total_posts: boardPosts.length, anchor_counts: {}, limit: 35, offset: 0 },
            recent_resource_comments: resourceComments,
            stats: { board_post_count: boardPosts.length, resource_comment_count: resourceComments.length, member_count: 1, banned_count: 0 },
          };
        } else if (url.pathname === '/api/community/board' && req.method() === 'POST') {
          const payload = JSON.parse(req.postData() || '{}');
          const post = {
            id: nextId++,
            scope: 'board',
            resource_path: '',
            resource_title: '',
            anchor_id: '',
            anchor_label: '',
            parent_id: null,
            reply_to_id: null,
            reply_to_nickname: '',
            author: currentUser,
            content: payload.content,
            is_pinned: false,
            like_count: 0,
            liked_by_me: false,
            can_delete: true,
            can_pin: true,
            status: 'active',
            created_at: new Date().toISOString(),
            created_ts: Math.floor(Date.now() / 1000),
            replies: [],
            reply_count: 0,
          };
          boardPosts.unshift(post);
          body = { ok: true, post };
        } else if (url.pathname === '/api/community/comments' && req.method() === 'GET') {
          const anchorCounts = {};
          for (const c of resourceComments) {
            if (c.anchor_id) {
              anchorCounts[c.anchor_id] = { count: (anchorCounts[c.anchor_id]?.count || 0) + 1, label: c.anchor_label };
            }
          }
          body = {
            items: resourceComments,
            count: resourceComments.length,
            total_threads: resourceComments.length,
            total_posts: resourceComments.length,
            anchor_counts: anchorCounts,
            limit: 60,
            offset: 0,
          };
        } else if (url.pathname === '/api/community/comments' && req.method() === 'POST') {
          const payload = JSON.parse(req.postData() || '{}');
          const post = {
            id: nextId++,
            scope: 'resource',
            resource_path: payload.resource_path,
            resource_title: resources[0].title,
            anchor_id: payload.anchor_id || '',
            anchor_label: payload.anchor_label || '',
            parent_id: null,
            reply_to_id: null,
            reply_to_nickname: '',
            author: currentUser,
            content: payload.content,
            is_pinned: false,
            like_count: 0,
            liked_by_me: false,
            can_delete: true,
            can_pin: true,
            status: 'active',
            created_at: new Date().toISOString(),
            created_ts: Math.floor(Date.now() / 1000),
            replies: [],
            reply_count: 0,
          };
          resourceComments.unshift(post);
          body = { ok: true, post };
        } else if (url.pathname.endsWith('/like') && req.method() === 'POST') {
          const id = Number(url.pathname.split('/').slice(-2)[0]);
          const found = [...boardPosts, ...resourceComments].find(x => x.id === id);
          if (found) {
            found.liked_by_me = !found.liked_by_me;
            found.like_count += found.liked_by_me ? 1 : -1;
          }
          body = { ok: true, post_id: id, liked: true, like_count: found ? found.like_count : 1 };
        } else {
          status = 404;
          body = { detail: 'not mocked' };
        }
        await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
      });

      const page = await context.newPage();
      page.setDefaultTimeout(12000);
      await page.goto(base, { waitUntil: 'domcontentloaded' });
      await page.getByRole('heading', { name: '界域公共留言板' }).waitFor();

      // Check zero horizontal overflow on home board
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true, `No overflow on home board at ${width}px`);

      // Perform QQ private verification code login flow
      await page.getByRole('button', { name: 'QQ 验证登录' }).click();
      await page.locator('#comm-qq-input').fill('1348697221');
      await page.getByRole('button', { name: '获取验证码' }).click();
      await page.getByTestId('qq-verification-code').waitFor();
      assert.equal(await page.getByTestId('qq-verification-code').textContent(), 'TRPG-TESTCODE8888');

      // Wait for automatic poll -> consume -> modal close -> logged-in state
      await page.locator('dialog.comm-auth-dialog[open]').waitFor({ state: 'hidden', timeout: 8000 });

      // Post on home message board
      await page.locator('.comm-board-main textarea').fill('端到端测试：跑团招募今晚八点发车！');
      await page.getByRole('button', { name: '发布留言' }).click();
      await page.locator('.comm-post-body').filter({ hasText: '端到端测试：跑团招募今晚八点发车！' }).waitFor();

      // Enter sequence library & open resource, then comment on a specific ability card anchor
      await page.getByRole('button', { name: '进入序列库 →' }).click();
      await page.locator('[data-resource-index="0"]').click();
      await page.waitForFunction(() => document.querySelector('.realm-document-heading h1')?.textContent === '001】时空锚定');
      await page.locator('.archive-entry').nth(1).getByRole('button', { name: '评论此卡' }).click();
      await page.locator('#resource-comments-panel textarea').fill('相位折跃点满后机动性极强。');
      await page.getByRole('button', { name: '发表评论' }).click();
      await page.locator('#resource-comments-panel .comm-post-body').filter({ hasText: '相位折跃点满后机动性极强。' }).waitFor();
      await page.locator('#resource-comments-panel .comm-anchor-pill').filter({ hasText: '#相位折跃' }).waitFor();

      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true, `No overflow on reader comments at ${width}px`);
      await page.screenshot({ path: path.join(out, `community-${width}.png`), fullPage: true, animations: 'disabled' });

      assert.deepEqual(pageErrors, []);
      await context.close();
    }
    console.log('Community & QQ Auth browser regressions passed.');
  } finally {
    await browser.close();
  }
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
