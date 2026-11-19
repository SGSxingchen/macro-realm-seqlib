// ==UserScript==
// @name         宏观界域论坛认证
// @author       LanStarD
// @version      1.1.0
// @description  为主群官方 QQ Bot 提供 /登录 6位一次性论坛登录短码签发（严格仅通过私聊发码，绝不在群公屏暴露验证码）
// @timestamp    1795073600
// @license      MIT
// ==/UserScript==

(() => {
  // 挂载到全群与私聊默认开启的 hg（宏观界域）扩展上
  let ext = seal.ext.find('hg') || seal.ext.find('宏观界域');
  if (!ext) {
    ext = seal.ext.new('seqlib_forum_auth', 'LanStarD', '1.1.0');
    seal.ext.register(ext);
  }

  seal.ext.registerStringConfig(ext, 'forum_auth_api_base', 'http://127.0.0.1:39510');
  seal.ext.registerStringConfig(ext, 'forum_auth_token', '');
  seal.ext.registerStringConfig(
    ext,
    'forum_allowed_group',
    'OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91'
  );
  seal.ext.registerStringConfig(ext, 'forum_bot_app_id', '1903175013');

  function normGroupToken(raw) {
    const s = String(raw || '').trim();
    const idx = s.lastIndexOf('-');
    if (s.toLowerCase().startsWith('openqq-group:') && idx !== -1) {
      return s.slice(idx + 1).trim().toUpperCase();
    }
    return s.toUpperCase();
  }

  function sendPrivateToUser(ctx, userId, text) {
    const ep =
      ctx.endPoint ||
      (seal.getEndPoints() || []).find(e => String(e.userId || '').startsWith('OpenQQ:')) ||
      (seal.getEndPoints() || [])[0];
    if (!ep) return false;
    const pm = seal.newMessage();
    pm.messageType = 'private';
    pm.sender.userId = userId;
    const tmpCtx = seal.createTempCtx(ep, pm);
    seal.replyPerson(tmpCtx, pm, text);
    return true;
  }

  // 被动监听主群日常发言：群友只要在主群说过话，自动静默登记主群成员资格，之后可直接私聊发 /登录
  const prevOnNotCmd = ext.onNotCommandReceived;
  ext.onNotCommandReceived = (ctx, msg) => {
    if (typeof prevOnNotCmd === 'function') {
      try {
        prevOnNotCmd(ctx, msg);
      } catch (_e) {}
    }
    try {
      if (ctx.isPrivate) return;
      const groupId = String(msg.groupId || ctx.group?.groupId || '').trim();
      const allowedRaw =
        seal.ext.getStringConfig(ext, 'forum_allowed_group') ||
        'OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91';
      const allowedSet = allowedRaw.split(/[,;\s]+/).map(normGroupToken).filter(Boolean);
      if (!allowedSet.includes('*') && !allowedSet.includes(normGroupToken(groupId))) return;

      const userId = String(msg.sender?.userId || ctx.player?.userId || '').trim();
      if (!userId) return;
      const seenKey = `bound_v1:${userId}`;
      if (ext.storageGet(seenKey)) return;

      const baseUrl = (seal.ext.getStringConfig(ext, 'forum_auth_api_base') || 'http://127.0.0.1:39510').replace(
        /\/+$/,
        ''
      );
      const token = (seal.ext.getStringConfig(ext, 'forum_auth_token') || '').trim();
      if (!token) return;
      const rawNick = String(ctx.player?.name || msg.sender?.nickname || '').trim();
      const cleanNick = rawNick.includes('|') ? rawNick.split('|')[0].trim() : rawNick;

      ext.storageSet(seenKey, String(Date.now()));
      fetch(`${baseUrl}/api/auth/bot/issue-code`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${token}`,
        },
        body: JSON.stringify({
          open_user_id: userId,
          open_group_id: groupId,
          sealdice_nickname: cleanNick,
          source: 'group_bind_only',
        }),
      }).catch(() => {});
    } catch (_err) {}
  };

  const cmd = seal.ext.newCmdItemInfo();
  cmd.name = '登录';
  cmd.help = '私聊机器人或在宏观界域主群发送 /登录 获取轮回者论坛 6 位一次性私密登录码';
  cmd.solve = (ctx, msg, _cmdArgs) => {
    const baseUrl = (seal.ext.getStringConfig(ext, 'forum_auth_api_base') || 'http://127.0.0.1:39510').replace(
      /\/+$/,
      ''
    );
    const token = (seal.ext.getStringConfig(ext, 'forum_auth_token') || '').trim();
    const botAppId = (seal.ext.getStringConfig(ext, 'forum_bot_app_id') || '1903175013').trim();
    if (!token) {
      seal.replyToSender(ctx, msg, '论坛登录密钥尚未配置，请联系维护者检查插件设置。');
      return seal.ext.newCmdExecuteResult(true);
    }

    const userId = String(msg.sender?.userId || ctx.player?.userId || '').trim();
    const rawNick = String(ctx.player?.name || msg.sender?.nickname || '').trim();
    const cleanNick = rawNick.includes('|') ? rawNick.split('|')[0].trim() : rawNick;

    // 1. 私聊触发：直接在私聊窗口返回 6 位登录码（最安全，绝无群公屏泄露风险）
    if (ctx.isPrivate) {
      fetch(`${baseUrl}/api/auth/bot/issue-code`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Authorization: `Bearer ${token}`,
        },
        body: JSON.stringify({
          open_user_id: userId,
          open_group_id: '',
          sealdice_nickname: cleanNick,
          bot_app_id: botAppId,
          source: 'private',
        }),
      })
        .then(async resp => {
          let data = {};
          try {
            data = await resp.json();
          } catch (_e) {
            data = {};
          }
          if (!resp.ok || !data.code) {
            const detail = data.detail || `服务暂不可用（HTTP ${resp.status}）`;
            seal.replyToSender(ctx, msg, `${detail}`);
            return;
          }
          const lines = [
            `【轮回者论坛认证 · 私密验证码】`,
            `您的 6 位一次性登录码：${data.code}`,
            `有效期 5 分钟，请勿转发他人。前往在线序列库点击右上角「轮回者登录」填入即可。`,
          ];
          seal.replyToSender(ctx, msg, lines.join('\n'));
        })
        .catch(err => {
          seal.replyToSender(ctx, msg, `获取论坛登录码失败，请稍后重试（${err}）`);
        });

      return seal.ext.newCmdExecuteResult(true);
    }

    // 2. 群聊内触发：先核验主群白名单，登记成员身份并通过私聊投递登录码，群公屏绝不输出任何数字验证码！
    const groupId = String(msg.groupId || ctx.group?.groupId || '').trim();
    const allowedRaw =
      seal.ext.getStringConfig(ext, 'forum_allowed_group') ||
      'OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91';
    const allowedSet = allowedRaw.split(/[,;\s]+/).map(normGroupToken).filter(Boolean);

    if (allowedSet.length > 0 && !allowedSet.includes('*') && !allowedSet.includes(normGroupToken(groupId))) {
      seal.replyToSender(ctx, msg, '仅限在宏观界域主群内登记身份后私聊机器人获取登录码。');
      return seal.ext.newCmdExecuteResult(true);
    }

    ext.storageSet(`bound_v1:${userId}`, String(Date.now()));
    fetch(`${baseUrl}/api/auth/bot/issue-code`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${token}`,
      },
      body: JSON.stringify({
        open_user_id: userId,
        open_group_id: groupId,
        sealdice_nickname: cleanNick,
        bot_app_id: botAppId,
        source: 'group',
      }),
    })
      .then(async resp => {
        let data = {};
        try {
          data = await resp.json();
        } catch (_e) {
          data = {};
        }
        if (!resp.ok || !data.code) {
          const detail = data.detail || `服务暂不可用（HTTP ${resp.status}）`;
          seal.replyToSender(ctx, msg, `认证请求未通过：${detail}`);
          return;
        }
        const dmLines = [
          `【轮回者论坛认证 · 私密验证码】`,
          `您的 6 位一次性登录码：${data.code}`,
          `有效期 5 分钟，请勿转发他人。前往在线序列库点击右上角「轮回者登录」填入即可。`,
        ];
        sendPrivateToUser(ctx, userId, dmLines.join('\n'));
        seal.replyToSender(
          ctx,
          msg,
          '已在主群完成身份核验，6 位登录码已通过【私聊】发送给您！（若因首次未开启机器人私聊未收到，请点击我的头像打开私聊窗口发送 /登录 获取）'
        );
      })
      .catch(err => {
        seal.replyToSender(ctx, msg, `认证服务连接失败，请稍后重试（${err}）`);
      });

    return seal.ext.newCmdExecuteResult(true);
  };

  ext.cmdMap['登录'] = cmd;
  ext.cmdMap['论坛登录'] = cmd;
  ext.cmdMap['login'] = cmd;
})();
