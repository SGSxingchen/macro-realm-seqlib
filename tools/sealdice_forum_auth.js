// ==UserScript==
// @name         宏观界域论坛认证
// @author       LanStarD
// @version      1.0.0
// @description  为主群官方 QQ Bot 提供 /登录 6位一次性论坛登录短码签发（严格不返回外部链接）
// @timestamp    1795070400
// @license      MIT
// ==/UserScript==

(() => {
  // 挂载到全群默认开启的 hg（宏观界域）扩展上，避免在已有群内额外手动 .ext 开启新扩展
  let ext = seal.ext.find('hg') || seal.ext.find('宏观界域');
  if (!ext) {
    ext = seal.ext.new('seqlib_forum_auth', 'LanStarD', '1.0.0');
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

  const cmd = seal.ext.newCmdItemInfo();
  cmd.name = '登录';
  cmd.help = '在宏观界域主群发送 /登录 获取轮回者论坛 6 位一次性登录码（5 分钟有效）';
  cmd.solve = (ctx, msg, _cmdArgs) => {
    if (ctx.isPrivate) {
      seal.replyToSender(ctx, msg, '请在宏观界域主群内唤醒机器人发送 /登录 获取论坛登录码。');
      return seal.ext.newCmdExecuteResult(true);
    }

    const groupId = String(msg.groupId || ctx.group?.groupId || '').trim();
    const allowedRaw =
      seal.ext.getStringConfig(ext, 'forum_allowed_group') ||
      'OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91';
    const allowedSet = allowedRaw
      .split(/[,;\s]+/)
      .map(normGroupToken)
      .filter(Boolean);

    if (allowedSet.length > 0 && !allowedSet.includes('*') && !allowedSet.includes(normGroupToken(groupId))) {
      seal.replyToSender(ctx, msg, '仅限在宏观界域主群内发送 /登录 获取论坛登录码。');
      return seal.ext.newCmdExecuteResult(true);
    }

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
    // 若名片带有宏观界域自动模板（如 "角色名 | HP 10/10 | ..."），仅截取竖线前的干净称呼作为默认昵称
    const cleanNick = rawNick.includes('|') ? rawNick.split('|')[0].trim() : rawNick;

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
          seal.replyToSender(ctx, msg, `未能生成登录码：${detail}`);
          return;
        }
        // 注意：官方 QQ Bot 严禁发送外部未备案 URL，回复内容严格保持纯文本无域名链接
        const lines = [
          `【轮回者论坛认证】`,
          `您的 6 位一次性登录码：${data.code}`,
          `有效期 5 分钟，请勿转发他人。前往在线序列库点击右上角「轮回者登录」填入即可。`,
        ];
        seal.replyToSender(ctx, msg, lines.join('\n'));
      })
      .catch(err => {
        seal.replyToSender(ctx, msg, `获取论坛登录码失败，请稍后重试（${err}）`);
      });

    return seal.ext.newCmdExecuteResult(true);
  };

  ext.cmdMap['登录'] = cmd;
  ext.cmdMap['论坛登录'] = cmd;
  ext.cmdMap['login'] = cmd;
})();
