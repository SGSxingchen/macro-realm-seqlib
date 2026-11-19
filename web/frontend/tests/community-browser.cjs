/** E2E regression: QQ group verification login (incl. WebView reload recovery), classic forum topic list + 1F/2F thread detail, resource anchor comments & visual captures. */
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
        viewport: { width, height: 920 },
        colorScheme: 'light',
        ...(width <= 820 ? { isMobile: true, hasTouch: true } : {}),
      });
      const pageErrors = [];
      context.on('page', p => p.on('pageerror', e => pageErrors.push(e.message)));

      let loggedIn = false;
      let pendingChallenge = null;
      let challengeVerified = false;
      const nowSec = Math.floor(Date.now() / 1000);
      const boardThreads = [
        {
          id: 1,
          scope: 'board',
          title: '【公告】宏观界域轮回者论坛启用与交易群验证说明',
          excerpt: '欢迎各位轮回者，在交易群（672204828）登记角色卡后即可发帖交流。',
          resource_path: '',
          resource_title: '',
          anchor_id: '',
          anchor_label: '',
          parent_id: null,
          reply_to_id: null,
          reply_to_nickname: '',
          author: { qq_masked: '853****398', qq: '853304398', nickname: '沧羽/明羽', avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', badge: '维护者' },
          content: '欢迎各位轮回者，在交易群（672204828）登记角色卡后即可发帖交流。',
          is_pinned: true,
          like_count: 5,
          liked_by_me: false,
          can_delete: true,
          can_pin: true,
          status: 'active',
          created_at: new Date((nowSec - 3600) * 1000).toISOString(),
          created_ts: nowSec - 3600,
          last_reply_ts: nowSec - 600,
          last_reply_nickname: '爱莉希雅/符华/希儿',
          replies: [
            {
              id: 2,
              floor: 2,
              scope: 'board',
              resource_path: '',
              resource_title: '',
              anchor_id: '',
              anchor_label: '',
              parent_id: 1,
              reply_to_id: 1,
              reply_to_nickname: '沧羽/明羽',
              author: { qq_masked: '320****101', qq: null, nickname: '爱莉希雅/符华/希儿', avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', badge: '' },
              content: '前排支持！终于有地方集中讨论序列构筑和发车招募了。',
              is_pinned: false,
              like_count: 2,
              liked_by_me: false,
              can_delete: false,
              can_pin: false,
              status: 'active',
              created_at: new Date((nowSec - 600) * 1000).toISOString(),
              created_ts: nowSec - 600,
            },
          ],
          reply_count: 1,
        },
        {
          id: 3,
          scope: 'board',
          title: '求问科技侧动力装甲配哪套能量池续航最稳？',
          excerpt: '目前准备进二阶团，队伍里没有专职辅助，想问问老轮回者们的经验。',
          resource_path: '',
          resource_title: '',
          anchor_id: '',
          anchor_label: '',
          parent_id: null,
          reply_to_id: null,
          reply_to_nickname: '',
          author: { qq_masked: '135****867', qq: null, nickname: '可爱纳垢虾/罗睿', avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', badge: '' },
          content: '目前准备进二阶团，队伍里没有专职辅助，想问问老轮回者们的经验。',
          is_pinned: false,
          like_count: 3,
          liked_by_me: false,
          can_delete: false,
          can_pin: true,
          status: 'active',
          created_at: new Date((nowSec - 1800) * 1000).toISOString(),
          created_ts: nowSec - 1800,
          last_reply_ts: nowSec - 1800,
          last_reply_nickname: '可爱纳垢虾/罗睿',
          replies: [],
          reply_count: 0,
        },
      ];
      const resourceComments = [
        {
          id: 10,
          scope: 'resource',
          resource_path: resources[0].path,
          resource_title: resources[0].title,
          anchor_id: 'entry-old',
          anchor_label: '时空锚定',
          parent_id: null,
          reply_to_id: null,
          reply_to_nickname: '',
          author: { qq_masked: '134****7221', qq: '1348697221', nickname: '无语', avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7', badge: '维护者' },
          content: '控场神技，注意10点消耗不可减免。',
          is_pinned: false,
          like_count: 4,
          liked_by_me: false,
          can_delete: true,
          can_pin: true,
          status: 'active',
          created_at: new Date((nowSec - 900) * 1000).toISOString(),
          created_ts: nowSec - 900,
          replies: [],
          reply_count: 0,
        },
      ];
      let nextId = 20;

      await context.route('**/api/**', async route => {
        const req = route.request();
        const url = new URL(req.url());
        let status = 200;
        let body = {};

        const currentUser = loggedIn
          ? {
              qq: '1348697221',
              qq_masked: '134****7221',
              nickname: '轮回者星尘',
              avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
              card_count: 2,
              badge: '维护者',
              is_banned: false,
            }
          : null;

        if (url.pathname === '/api/tree') {
          body = {
            items: [{
              name: '序列库',
              path: '序列库',
              count: 692,
              children: [
                { name: '特质改造', path: '序列库/特质改造', count: 268, children: [] },
                { name: '技能表', path: '序列库/技能表', count: 210, children: [] },
                { name: '职业', path: '序列库/职业', count: 182, children: [] },
                { name: '能量池', path: '序列库/能量池', count: 18, children: [] },
                { name: '公共建筑', path: '序列库/公共建筑', count: 13, children: [] },
              ],
            }],
          };
        } else if (url.pathname === '/api/resources') {
          body = { items: resources, count: 1, total: 692, offset: 0, limit: 100, tokens: [], facets: { kinds: [], sides: [], authors: [] } };
        } else if (url.pathname.startsWith('/api/resources/')) {
          body = resources[0];
        } else if (url.pathname === '/api/auth/me') {
          body = {
            authenticated: loggedIn,
            auth_configured: true,
            auth_mode: 'official_bot_code',
            bot_name: '阿瓦娜',
            bot_app_id: '1903175013',
            bot_qq: '412318056',
            group_id: 672204828,
            admin_authenticated: false,
            is_moderator: loggedIn,
            pending_challenge: null,
            user: currentUser,
          };
        } else if (url.pathname === '/api/auth/web-challenges' && req.method() === 'POST') {
          body = { status: 'pending', command: '/登录 839204', bot_name: '阿瓦娜', expires_at: new Date(Date.now() + 300000).toISOString() };
        } else if (url.pathname === '/api/auth/web-challenges/current' && challengeVerified) {
          body = {
            status: 'verified',
            confirmation: 'browser-confirmation',
            user_key_masked: 'ID:FF83…0D3A',
            is_new_user: true,
            suggested_nickname: '牛奶泡LanStarD',
            sealdice_nickname: '牛奶泡LanStarD',
            avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
            expires_at: new Date(Date.now() + 300000).toISOString(),
          };
        } else if (url.pathname === '/api/auth/web-challenges/consume' && req.method() === 'POST') {
          assert.equal(req.headers()['x-login-confirmation'], 'browser-confirmation');
          const payload = JSON.parse(req.postData() || '{}');
          assert.deepEqual(Object.keys(payload), ['nickname']);
          loggedIn = true;
          const finalNick = (payload.nickname || '').trim() || '轮回者星尘';
          body = {
            ok: true,
            authenticated: true,
            auth_configured: true,
            auth_mode: 'official_bot_code',
            bot_name: '阿瓦娜',
            bot_app_id: '1903175013',
            bot_qq: '412318056',
            group_id: 672204828,
            admin_authenticated: false,
            is_moderator: true,
            user: {
              qq: 'oid:FF834522288AB2D2993970BE14970D3A',
              qq_masked: 'ID:FF83…0D3A',
              nickname: finalNick,
              avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
              card_count: 0,
              badge: '维护者',
              is_banned: false,
            },
          };
        } else if (url.pathname === '/api/auth/profile' && req.method() === 'PATCH') {
          const payload = JSON.parse(req.postData() || '{}');
          body = {
            ok: true,
            authenticated: true,
            auth_configured: true,
            auth_mode: 'official_bot_code',
            bot_name: '阿瓦娜',
            bot_app_id: '1903175013',
            bot_qq: '412318056',
            group_id: 672204828,
            admin_authenticated: false,
            is_moderator: true,
            user: {
              qq: 'oid:FF834522288AB2D2993970BE14970D3A',
              qq_masked: 'ID:FF83…0D3A',
              nickname: payload.nickname || '轮回者星尘',
              avatar_url: 'data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7',
              card_count: 0,
              badge: '维护者',
              is_banned: false,
            },
          };
        } else if (url.pathname === '/api/community/home') {
          body = {
            board: {
              items: boardThreads.map(({ replies, ...rest }) => rest),
              count: boardThreads.length,
              total_threads: boardThreads.length,
              total_posts: boardThreads.reduce((s, t) => s + 1 + (t.replies?.length || 0), 0),
              anchor_counts: {},
              limit: 45,
              offset: 0,
            },
            recent_resource_comments: resourceComments,
            stats: {
              board_thread_count: boardThreads.length,
              board_post_count: boardThreads.length + 1,
              resource_comment_count: resourceComments.length,
              member_count: 12,
              banned_count: 0,
            },
          };
        } else if (url.pathname.startsWith('/api/community/threads/')) {
          const tid = Number(url.pathname.split('/').pop());
          const found = boardThreads.find(t => t.id === tid);
          body = { thread: { ...found, floor: 1 } };
        } else if (url.pathname === '/api/community/board' && req.method() === 'POST') {
          const payload = JSON.parse(req.postData() || '{}');
          if (payload.reply_to_id) {
            const targetThread = boardThreads.find(t => t.id === payload.reply_to_id) || boardThreads[0];
            const rep = {
              id: nextId++,
              floor: (targetThread.replies?.length || 0) + 2,
              scope: 'board',
              resource_path: '',
              resource_title: '',
              anchor_id: '',
              anchor_label: '',
              parent_id: targetThread.id,
              reply_to_id: payload.reply_to_id,
              reply_to_nickname: targetThread.author.nickname,
              author: currentUser,
              content: payload.content,
              is_pinned: false,
              like_count: 0,
              liked_by_me: false,
              can_delete: true,
              can_pin: false,
              status: 'active',
              created_at: new Date().toISOString(),
              created_ts: Math.floor(Date.now() / 1000),
            };
            targetThread.replies.push(rep);
            targetThread.reply_count = targetThread.replies.length;
            body = { ok: true, post: rep };
          } else {
            const thread = {
              id: nextId++,
              scope: 'board',
              title: payload.title,
              excerpt: payload.content.slice(0, 80),
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
              last_reply_ts: Math.floor(Date.now() / 1000),
              last_reply_nickname: currentUser.nickname,
              replies: [],
              reply_count: 0,
            };
            boardThreads.splice(1, 0, thread);
            body = { ok: true, post: thread };
          }
        } else if (url.pathname === '/api/community/comments' && req.method() === 'GET') {
          const anchorCounts = {};
          const labelCounts = {};
          for (const c of resourceComments) {
            if (c.anchor_label) {
              anchorCounts[c.anchor_id || c.anchor_label] = { count: (anchorCounts[c.anchor_id || c.anchor_label]?.count || 0) + 1, label: c.anchor_label, anchor_id: c.anchor_id };
              labelCounts[c.anchor_label] = (labelCounts[c.anchor_label] || 0) + 1;
            }
          }
          body = {
            items: resourceComments,
            count: resourceComments.length,
            total_threads: resourceComments.length,
            total_posts: resourceComments.length,
            anchor_counts: anchorCounts,
            label_counts: labelCounts,
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
        } else {
          status = 404;
          body = { detail: 'not mocked' };
        }
        await route.fulfill({ status, contentType: 'application/json', body: JSON.stringify(body) });
      });

      const page = await context.newPage();
      page.setDefaultTimeout(12000);
      await page.goto(base, { waitUntil: 'domcontentloaded' });
      await page.getByRole('heading', { name: '轮回者终端论坛' }).waitFor();
      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true, `No overflow on forum list at ${width}px`);

      // Official Bot 6-digit short code login + default SealDice nickname editable flow
      if (width <= 820) {
        await page.getByRole('button', { name: '导航', exact: true }).click();
      }
      await page.getByRole('button', { name: '轮回者登录' }).click();
      assert.equal(await page.getByTestId('web-login-command').textContent(), '/登录 839204');
      challengeVerified = true;
      const nickInput = page.getByTestId('bot-login-nickname-input');
      await nickInput.waitFor();
      assert.equal(await nickInput.inputValue(), '牛奶泡LanStarD');
      await nickInput.fill('轮回者星尘');
      await page.getByRole('button', { name: '确认并登录' }).click();
      await page.waitForFunction(() => document.body.textContent.includes('轮回者星尘'));

      await page.screenshot({ path: path.join(out, `forum-light-${width}.png`), fullPage: true, animations: 'disabled' });
      if (width === 1440) {
        await page.getByRole('button', { name: '切换主题' }).click();
        await page.waitForFunction(() => document.documentElement.dataset.theme === 'dark');
        await page.screenshot({ path: path.join(out, `forum-dark-${width}.png`), fullPage: true, animations: 'disabled' });
      }

      // Create new forum topic via modal & view 1F / 2F thread detail
      await page.getByRole('button', { name: '+ 发布新主题', exact: true }).click();
      await page.locator('#forum-input-title').fill('周六晚八点二阶无限流模组招募（3=1）');
      await page.locator('#forum-input-body').fill('已有前排武僧和神秘侧控场，来一位远程火力或医疗轮回者。');
      await page.getByRole('button', { name: '发布主题' }).click();

      // Automatically navigated into thread detail 1F view
      await page.getByRole('heading', { name: '周六晚八点二阶无限流模组招募（3=1）' }).waitFor();
      assert.ok((await page.locator('#floor-1').textContent()).includes('1F'));

      // Reply to create 2F
      await page.locator('.forum-reply-box textarea').fill('科技侧狙击手报名，卡已在交易群上传！');
      await page.getByRole('button', { name: '发送回帖' }).click();
      await page.locator('#floor-2').filter({ hasText: '科技侧狙击手报名' }).waitFor();
      if (width === 1440) {
        await page.screenshot({ path: path.join(out, `forum-thread-detail-${width}.png`), fullPage: true, animations: 'disabled' });
      }

      await page.getByRole('button', { name: '← 返回主题列表' }).click();
      await page.getByRole('heading', { name: '轮回者终端论坛' }).waitFor();

      // Open sequence library & verify label-based resilient anchor badge & comment
      if (width <= 820) {
        await page.getByRole('button', { name: '导航', exact: true }).click();
      }
      await page.getByRole('button', { name: '序列检索', exact: true }).click();
      await page.locator('[data-resource-index="0"]').click();
      await page.waitForFunction(() => document.querySelector('.realm-document-heading h1')?.textContent === '001】时空锚定');
      // Even though mock comment had old anchor_id 'entry-old', label_counts['时空锚定']=1 matches the first ability card!
      await page.locator('.archive-entry').first().getByRole('button', { name: '评论 (1)' }).waitFor();

      await page.locator('.archive-entry').nth(1).getByRole('button', { name: '评论此卡' }).click();
      await page.locator('#resource-comments-panel textarea').fill('相位折跃点满后机动性极强。');
      await page.getByRole('button', { name: '发表评论' }).click();
      await page.locator('#resource-comments-panel .comm-post-body').filter({ hasText: '相位折跃点满后机动性极强。' }).waitFor();

      assert.equal(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1), true, `No overflow on reader comments at ${width}px`);
      assert.deepEqual(pageErrors, []);
      await context.close();
    }
    console.log('Community Forum, QQ WebView reload recovery & Resilient Anchor browser regressions passed.');
  } finally {
    await browser.close();
  }
})().catch(error => {
  console.error(error);
  process.exitCode = 1;
});
