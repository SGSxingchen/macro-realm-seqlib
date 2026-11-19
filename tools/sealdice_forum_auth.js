// ==UserScript==
// @name         宏观界域论坛认证
// @author       LanStarD
// @version      2.0.0
// @description  在宏观界域主群用 /登录 核验网页验证码
// @timestamp    1795075200
// @license      MIT
// ==/UserScript==

(() => {
  let ext = seal.ext.find('hg') || seal.ext.find('宏观界域');
  if (!ext) {
    ext = seal.ext.new('seqlib_forum_auth', 'LanStarD', '2.0.0');
    seal.ext.register(ext);
  }

  const mainGroup = 'OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91';
  seal.ext.registerStringConfig(ext, 'forum_auth_api_base', 'http://127.0.0.1:39510');
  seal.ext.registerStringConfig(ext, 'forum_auth_token', '');
  seal.ext.registerStringConfig(ext, 'forum_allowed_group', mainGroup);
  seal.ext.registerStringConfig(ext, 'forum_bot_app_id', '1903175013');

  function groupToken(raw) {
    const value = raw.trim();
    if (/^openqq-group:/i.test(value)) return value.slice(value.lastIndexOf('-') + 1).toUpperCase();
    return value.toUpperCase();
  }

  const cmd = seal.ext.newCmdItemInfo();
  cmd.name = '登录';
  cmd.help = '在在线序列库点「轮回者登录」获取指令，再到宏观界域主群 @机器人 发送 /登录 6位验证码。';
  cmd.solve = (ctx, msg, args) => {
    const result = seal.ext.newCmdExecuteResult(true);
    const groupId = String(msg.groupId || '').trim();
    const userId = String(msg.sender?.userId || '').trim();
    if (ctx.isPrivate || !/^OpenQQ-Group:[0-9]+-[0-9a-f]{16,64}$/i.test(groupId)
      || !/^OpenQQ:[0-9]+-[0-9a-f]{16,64}$/i.test(userId)) {
      seal.replyToSender(ctx, msg, '请在宏观界域主群使用官方机器人核验论坛身份。');
      return result;
    }
    const allowed = (seal.ext.getStringConfig(ext, 'forum_allowed_group') || mainGroup)
      .split(/[,;\s]+/).filter(Boolean).map(groupToken);
    if (!allowed.includes(groupToken(groupId)) && !allowed.includes('*')) {
      seal.replyToSender(ctx, msg, '仅限宏观界域主群使用论坛登录。');
      return result;
    }
    const code = String(args.getArgN(1) || '').trim();
    if (!/^\d{6}$/.test(code)) {
      seal.replyToSender(ctx, msg, cmd.help);
      return result;
    }
    const token = seal.ext.getStringConfig(ext, 'forum_auth_token').trim();
    if (!token) {
      seal.replyToSender(ctx, msg, '论坛核验服务尚未配置，请联系维护者。');
      return result;
    }
    const base = seal.ext.getStringConfig(ext, 'forum_auth_api_base').trim().replace(/\/+$/, '');
    const appId = seal.ext.getStringConfig(ext, 'forum_bot_app_id').trim();
    const nickname = String(ctx.player?.name || msg.sender?.nickname || '').split('|')[0].trim();
    let timedOut = false;
    const timer = setTimeout(() => {
      timedOut = true;
      seal.replyToSender(ctx, msg, '核验响应超时，请先回网页查看状态。');
    }, 6000);
    fetch(base + '/api/auth/bot/verify-code', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: 'Bearer ' + token },
      body: JSON.stringify({
        code, open_user_id: userId, open_group_id: groupId,
        sealdice_nickname: nickname.slice(0, 80), bot_app_id: appId,
      }),
    }).then(async response => {
      const data = await response.json().catch(() => ({}));
      if (timedOut) return;
      if (response.ok && data.status === 'verified') {
        seal.replyToSender(ctx, msg, '论坛身份已核验，请回到网页完成登录。');
      } else if ([404, 409, 410].includes(response.status) && typeof data.detail === 'string') {
        seal.replyToSender(ctx, msg, data.detail);
      } else {
        seal.replyToSender(ctx, msg, '论坛核验暂未完成，请稍后再试。');
      }
    }).catch(() => {
      if (!timedOut) seal.replyToSender(ctx, msg, '暂时无法连接论坛核验服务，请稍后再试。');
    }).finally(() => clearTimeout(timer));
    return result;
  };

  ext.cmdMap['登录'] = cmd;
  ext.cmdMap['论坛登录'] = cmd;
  ext.cmdMap['login'] = cmd;
})();
