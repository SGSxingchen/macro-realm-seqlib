"""后端单元测试：QQ 私聊验证码三步登录安全契约、会话绑定、留言板、资源评论（锚点/回复/点赞）与管理后台。"""
from __future__ import annotations

import json
import os
import sys
import tempfile
import unittest
from pathlib import Path

import httpx
from fastapi.testclient import TestClient

BACKEND_DIR = Path(__file__).resolve().parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from app import main as app_main
from app.community import CommunityStore, UpstreamAuthClient, compute_resource_key


class CommunityAndQqAuthTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmpdir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.tmpdir.name) / "test_community.sqlite"
        self.prev_env = {
            "ADMIN_PASSWORD": os.environ.get("ADMIN_PASSWORD"),
            "COMMUNITY_SECRET": os.environ.get("COMMUNITY_SECRET"),
            "ONEBOT_SEQUENCE_TOKEN": os.environ.get("ONEBOT_SEQUENCE_TOKEN"),
            "COMMUNITY_ADMIN_QQS": os.environ.get("COMMUNITY_ADMIN_QQS"),
            "COMMUNITY_COOLDOWN_SECONDS": os.environ.get("COMMUNITY_COOLDOWN_SECONDS"),
        }
        os.environ["ADMIN_PASSWORD"] = "test-admin-pw"
        os.environ["COMMUNITY_SECRET"] = "test-community-secret"
        os.environ["ONEBOT_SEQUENCE_TOKEN"] = "mock-secret-bearer-token"
        os.environ["COMMUNITY_ADMIN_QQS"] = "1348697221"
        os.environ["COMMUNITY_COOLDOWN_SECONDS"] = "0"

        self.challenges: dict[str, dict] = {}
        self.upstream_calls: list[tuple[str, str, str | None]] = []

        def handler(request: httpx.Request) -> httpx.Response:
            auth = request.headers.get("Authorization")
            self.upstream_calls.append((request.method, request.url.path, auth))
            if auth != "Bearer mock-secret-bearer-token":
                return httpx.Response(401, json={"error": "unauthorized"})

            if request.method == "POST" and request.url.path == "/onebot-sequences/v1/auth/challenges":
                payload = json.loads(request.content.decode("utf-8"))
                qq = payload.get("qq", "")
                if qq == "99999999":
                    return httpx.Response(403, json={"error": "not_in_group"})
                cid = f"chal_{qq}_123456"
                code = f"TRPG-CODE{qq[-4:]}"
                record = {
                    "challenge_id": cid,
                    "qq": qq,
                    "group_id": 672204828,
                    "bot_qq": "412318056",
                    "verification_code": code,
                    "message_to_send": code,
                    "expires_at": "2026-11-16T12:00:00+00:00",
                    "status": "pending",
                    "verified_at": None,
                }
                self.challenges[cid] = record
                return httpx.Response(201, json=record)

            if request.method == "GET" and request.url.path.startswith("/onebot-sequences/v1/auth/challenges/"):
                cid = request.url.path.rsplit("/", 1)[-1]
                rec = self.challenges.get(cid)
                if not rec:
                    return httpx.Response(404, json={"error": "challenge_not_found"})
                return httpx.Response(
                    200,
                    json={
                        "challenge_id": rec["challenge_id"],
                        "qq": rec["qq"],
                        "group_id": rec["group_id"],
                        "bot_qq": rec["bot_qq"],
                        "status": rec["status"],
                        "expires_at": rec["expires_at"],
                        "verified_at": rec["verified_at"],
                    },
                )

            if request.method == "POST" and request.url.path.endswith("/consume"):
                cid = request.url.path.split("/")[-2]
                rec = self.challenges.get(cid)
                if not rec:
                    return httpx.Response(404, json={"error": "challenge_not_found"})
                if rec["status"] == "pending":
                    return httpx.Response(409, json={"error": "verification_pending"})
                if rec["status"] == "consumed":
                    return httpx.Response(409, json={"error": "already_consumed"})
                if rec["qq"] == "88888888":
                    rec["status"] = "denied"
                    return httpx.Response(403, json={"error": "no_uploaded_card", "passed": False})
                rec["status"] = "consumed"
                return httpx.Response(
                    200,
                    json={
                        "challenge_id": cid,
                        "qq": rec["qq"],
                        "group_id": 672204828,
                        "bot_qq": "412318056",
                        "status": "consumed",
                        "expires_at": rec["expires_at"],
                        "verified_at": rec["verified_at"] or "2026-11-16T11:56:00+00:00",
                        "passed": True,
                        "group_nickname": f"群名片_{rec['qq'][-4:]}",
                        "avatar_url": f"https://q1.qlogo.cn/g?b=qq&nk={rec['qq']}&s=640",
                        "avatar_scope": "qq",
                        "card_count": 3,
                        "last_success": "2026-11-16T11:55:00+00:00",
                        "stale": False,
                        "auth_type": "qq_private_code",
                    },
                )

            return httpx.Response(404, json={"error": "unknown"})

        app_main.COMMUNITY_STORE = CommunityStore(self.db_path)
        app_main.UPSTREAM_AUTH_CLIENT = UpstreamAuthClient(
            base_url="https://trpg.chordvers.org/onebot-sequences",
            transport=httpx.MockTransport(handler),
        )
        app_main.SEARCH_INDEX.refresh()
        self.sample_path = app_main.list_resources(limit=1)["items"][0]["path"]

    def tearDown(self) -> None:
        app_main.COMMUNITY_STORE = None
        app_main.UPSTREAM_AUTH_CLIENT = None
        for k, v in self.prev_env.items():
            if v is None:
                os.environ.pop(k, None)
            else:
                os.environ[k] = v
        self.tmpdir.cleanup()

    def _login_as(self, client: TestClient, qq: str) -> dict:
        r1 = client.post("/api/auth/qq/challenges", json={"qq": qq})
        self.assertEqual(r1.status_code, 200, r1.text)
        body1 = r1.json()
        # 确保响应绝不泄露 Bearer token 或内部 challenge_id
        self.assertNotIn("mock-secret-bearer-token", r1.text)
        self.assertNotIn("challenge_id", body1)
        self.assertEqual(r1.headers.get("cache-control"), "private, no-store")

        # 模拟 BOT 收到私聊把对应 challenge 置为 verified
        self.challenges[f"chal_{qq}_123456"]["status"] = "verified"
        self.challenges[f"chal_{qq}_123456"]["verified_at"] = "2026-11-16T11:56:00+00:00"

        poll = client.get("/api/auth/qq/challenges/current")
        self.assertEqual(poll.status_code, 200)
        self.assertEqual(poll.json()["status"], "verified")
        # verified 状态查询绝不能自动把用户登录
        self.assertFalse(client.get("/api/auth/me").json()["authenticated"])

        consume = client.post("/api/auth/qq/challenges/consume")
        self.assertEqual(consume.status_code, 200, consume.text)
        data = consume.json()
        self.assertTrue(data["authenticated"])
        self.assertEqual(data["user"]["qq"], qq)
        return data

    def test_qq_login_security_boundaries(self) -> None:
        client_a = TestClient(app_main.app)
        client_b = TestClient(app_main.app)

        # 1. 非法 QQ 拦截
        self.assertEqual(client_a.post("/api/auth/qq/challenges", json={"qq": "012345"}).status_code, 400)
        # 2. URL 带查询参数拦截
        self.assertEqual(client_a.post("/api/auth/qq/challenges?foo=1", json={"qq": "12345678"}).status_code, 400)
        # 3. 不在群返回 403
        self.assertEqual(client_a.post("/api/auth/qq/challenges", json={"qq": "99999999"}).status_code, 403)

        # 4. 正常创建后：页面若在 QQ 内置 WebView 被切后台刷新，/api/auth/me 能恢复 pending_challenge 与原验证码
        r = client_a.post("/api/auth/qq/challenges", json={"qq": "12345678"})
        self.assertEqual(r.status_code, 200)
        code_issued = r.json()["verification_code"]
        self.assertTrue(code_issued.startswith("TRPG-"))
        me_pending = client_a.get("/api/auth/me").json()["pending_challenge"]
        self.assertEqual(me_pending["qq"], "12345678")
        self.assertEqual(me_pending["verification_code"], code_issued)
        self.assertEqual(client_a.post("/api/auth/qq/challenges/consume").status_code, 409)

        # 5. 另一个浏览器 client_b 无法替 client_a 轮询或抢先消费
        self.challenges["chal_12345678_123456"]["status"] = "verified"
        self.assertEqual(client_b.get("/api/auth/qq/challenges/current").status_code, 404)
        self.assertEqual(client_b.post("/api/auth/qq/challenges/consume").status_code, 400)
        # 禁止客户端在 consume 正文注入自定义 challenge_id / qq
        self.assertEqual(
            client_a.post("/api/auth/qq/challenges/consume", json={"challenge_id": "chal_12345678_123456"}).status_code,
            400,
        )

        # 6. client_a 正常消费登录；二次重复消费失败
        ok = client_a.post("/api/auth/qq/challenges/consume")
        self.assertEqual(ok.status_code, 200)
        self.assertEqual(ok.json()["user"]["nickname"], "群名片_5678")
        self.assertEqual(client_a.post("/api/auth/qq/challenges/consume").status_code, 400)

        # 7. 无角色卡用户在 consume 阶段被 403 拒绝
        r_nocard = client_b.post("/api/auth/qq/challenges", json={"qq": "88888888"})
        self.assertEqual(r_nocard.status_code, 200)
        self.challenges["chal_88888888_123456"]["status"] = "verified"
        denied = client_b.post("/api/auth/qq/challenges/consume")
        self.assertEqual(denied.status_code, 403)
        self.assertFalse(client_b.get("/api/auth/me").json()["authenticated"])

    def test_board_comments_anchors_likes_and_moderation(self) -> None:
        anon = TestClient(app_main.app)
        member = TestClient(app_main.app)
        mod_qq_client = TestClient(app_main.app)
        admin_client = TestClient(app_main.app)

        # 未登录只读可看，不可发帖或点赞
        self.assertEqual(anon.get("/api/community/home").status_code, 200)
        self.assertEqual(anon.post("/api/community/board", json={"content": "匿名偷发"}).status_code, 401)

        # 普通群员与白名单管理员 QQ 登录
        self._login_as(member, "23456789")
        mod_info = self._login_as(mod_qq_client, "1348697221")
        self.assertTrue(mod_info["is_moderator"])

        # 普通成员在主页留言板发帖
        p1 = member.post("/api/community/board", json={"content": "今天跑团大成功！"}).json()["post"]
        self.assertEqual(p1["content"], "今天跑团大成功！")
        self.assertFalse(p1["can_pin"])

        # 普通成员在具体资源的能力卡锚点下发评，管理员 QQ 进行楼中楼回复 + 点赞
        c1 = member.post(
            "/api/community/comments",
            json={
                "resource_path": self.sample_path,
                "anchor_id": "entry-abc-1",
                "anchor_label": "时空锚定",
                "content": "这张卡的冷却在团战里很关键。",
            },
        ).json()["post"]
        self.assertEqual(c1["anchor_id"], "entry-abc-1")

        reply = mod_qq_client.post(
            "/api/community/comments",
            json={
                "resource_path": self.sample_path,
                "content": "同意，配合队友先手很稳。",
                "reply_to_id": c1["id"],
            },
        ).json()["post"]
        self.assertEqual(reply["parent_id"], c1["id"])
        self.assertEqual(reply["anchor_id"], "entry-abc-1")

        like1 = mod_qq_client.post(f"/api/community/posts/{c1['id']}/like").json()
        self.assertTrue(like1["liked"])
        self.assertEqual(like1["like_count"], 1)

        # 按锚点查询评论计数与列表
        res_comments = anon.get("/api/community/comments", params={"path": self.sample_path}).json()
        self.assertEqual(res_comments["total_posts"], 2)
        self.assertEqual(res_comments["anchor_counts"]["entry-abc-1"]["count"], 2)
        # 未登录访客看不到完整 QQ 号，只能看到脱敏号
        self.assertIsNone(res_comments["items"][0]["author"]["qq"])
        self.assertIn("*", res_comments["items"][0]["author"]["qq_masked"])

        # 白名单管理员 QQ 在前台直接置顶留言
        pinned = mod_qq_client.post(f"/api/community/posts/{p1['id']}/pin", json={"pinned": True})
        self.assertEqual(pinned.status_code, 200)
        self.assertTrue(pinned.json()["post"]["is_pinned"])

        # ADMIN_PASSWORD 维护后台登录：发布置顶公告、封禁违规 QQ、删评
        self.assertEqual(admin_client.post("/api/admin/login", json={"password": "test-admin-pw"}).status_code, 200)
        ann = admin_client.post(
            "/api/community/moderation/announcement",
            json={"content": "欢迎来到界域留言板，请遵守跑团礼仪。", "pinned": True},
        )
        self.assertEqual(ann.status_code, 200)

        ban = admin_client.post(
            "/api/community/moderation/ban",
            json={"qq": "23456789", "duration_hours": 24, "reason": "刷屏测试"},
        )
        self.assertEqual(ban.status_code, 200)

        # 被封禁成员立即无法继续发言
        blocked = member.post("/api/community/board", json={"content": "还能发吗"})
        self.assertEqual(blocked.status_code, 403)
        self.assertIn("刷屏测试", blocked.text)

        # 解除封禁后恢复
        admin_client.post("/api/community/moderation/ban", json={"qq": "23456789", "unban": True})
        # 删除评论
        del_res = admin_client.delete(f"/api/community/posts/{c1['id']}")
        self.assertEqual(del_res.status_code, 200)
        after = anon.get("/api/community/comments", params={"path": self.sample_path}).json()
        self.assertEqual(after["total_posts"], 0)

    def test_forum_threads_and_resource_rename_resilience(self) -> None:
        store = app_main.COMMUNITY_STORE
        assert store is not None

        # 1. 验证 compute_resource_key 剥离编号、侧别子目录与新过审括号
        self.assertEqual(compute_resource_key("序列库/职业/神秘侧/001】天师.txt"), "职业::天师")
        self.assertEqual(compute_resource_key("序列库/职业/战技侧/099】天师.txt"), "职业::天师")
        self.assertEqual(compute_resource_key("序列库/新过审序列/012】剑圣（职业·战技侧）.txt"), "职业::剑圣")

        user = {"qq": "23456789", "nickname": "轮回者甲", "avatar_url": ""}
        user2 = {"qq": "34567890", "nickname": "轮回者乙", "avatar_url": ""}

        # 2. 模拟旧路径 `001】天师.txt` 下的评论（带品阶括号的能力名）
        c_old = store.create_post(
            scope="resource",
            author=user,
            author_badge="",
            content="旧序号 001 时期留下的天师评论",
            resource_path="序列库/职业/神秘侧/001】天师.txt",
            resource_title="001】天师",
            anchor_id="entry-old-hash-1",
            anchor_label="雷法精通（B级）",
        )
        self.assertEqual(c_old["anchor_label"], "雷法精通")

        # 3. 模拟 renumber_resources.py 把文件重排成 `018】天师.txt`，且能力升为 A 级导致前端 hash 变化：
        # 无需手工迁移，直接用新路径查询即可命中并自愈路径与标题
        after_renumber = store.list_threads(
            scope="resource",
            resource_path="序列库/职业/神秘侧/018】天师.txt",
            resource_title="018】天师",
        )
        self.assertEqual(after_renumber["total_posts"], 1)
        self.assertEqual(after_renumber["items"][0]["resource_path"], "序列库/职业/神秘侧/018】天师.txt")
        self.assertEqual(after_renumber["label_counts"]["雷法精通"], 1)

        # 按新能力名过滤（即使 anchor_id 变成 entry-new-hash-1）依然能筛出原卡评论
        filtered_by_label = store.list_threads(
            scope="resource",
            resource_path="序列库/职业/神秘侧/018】天师.txt",
            resource_title="018】天师",
            anchor_id="entry-new-hash-1",
            anchor_label="雷法精通（A级）",
        )
        self.assertEqual(filtered_by_label["total_threads"], 1)

        # 4. 模拟后台彻底改名（如 `018】天师.txt` -> `020】张天师《道门》.txt`），record_resource_move 自动接力
        store.record_resource_move(
            "序列库/职业/神秘侧/018】天师.txt",
            "序列库/职业/神秘侧/020】张天师《道门》.txt",
            "020】张天师《道门》",
        )
        after_rename = store.list_threads(
            scope="resource",
            resource_path="序列库/职业/神秘侧/020】张天师《道门》.txt",
            resource_title="020】张天师《道门》",
        )
        self.assertEqual(after_rename["total_posts"], 1)
        self.assertEqual(after_rename["items"][0]["resource_title"], "020】张天师《道门》")

        # 5. 论坛主题标题、楼层编号与最后回复顶帖排序
        t1 = store.create_post(scope="board", author=user, author_badge="", title="第一帖：求推荐科技侧前排", content="如题，团队缺主坦", bypass_cooldown=True)
        t2 = store.create_post(scope="board", author=user2, author_badge="", title="第二帖：今晚八点模组发车", content="缺1奶妈", bypass_cooldown=True)
        # 回复第一帖后，第一帖应被顶到最前，且详情页楼层为 1F + 2F
        store.create_post(scope="board", author=user2, author_badge="", content="动力装甲或者义体改造都很顶", reply_to_id=t1["id"], bypass_cooldown=True)

        board_list = store.list_threads(scope="board", sort="latest", include_replies=False)
        self.assertEqual(board_list["items"][0]["id"], t1["id"])
        self.assertEqual(board_list["items"][0]["reply_count"], 1)
        self.assertEqual(board_list["items"][0]["last_reply_nickname"], "轮回者乙")
        self.assertEqual(board_list["items"][1]["id"], t2["id"])

        detail = store.get_board_thread_detail(thread_id=t1["id"])["thread"]
        self.assertEqual(detail["floor"], 1)
        self.assertEqual(detail["replies"][0]["floor"], 2)

    def test_official_bot_short_code_login_and_nickname_edit(self) -> None:
        bot_bridge = TestClient(app_main.app)
        browser = TestClient(app_main.app)

        # 1. 未带 Bearer 密钥或非主群请求拒绝签发
        unauth = bot_bridge.post(
            "/api/auth/bot/issue-code",
            json={
                "open_user_id": "OpenQQ:4019653007-49E8BA38D36D9D9EEEA615D657C94E2C",
                "open_group_id": "OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91",
                "sealdice_nickname": "雨霁",
            },
        )
        self.assertEqual(unauth.status_code, 401)

        wrong_group = bot_bridge.post(
            "/api/auth/bot/issue-code",
            headers={"Authorization": "Bearer mock-secret-bearer-token"},
            json={
                "open_user_id": "OpenQQ:4019653007-49E8BA38D36D9D9EEEA615D657C94E2C",
                "open_group_id": "OpenQQ-Group:4019653007-CAB18FD223F3C863DC20A42B6456BA16",
                "sealdice_nickname": "雨霁",
            },
        )
        self.assertEqual(wrong_group.status_code, 403)

        # 2. 主群成员正常签发 6 位短码，预检返回海豹默认昵称与 qlogo 官方 Bot 头像
        issued = bot_bridge.post(
            "/api/auth/bot/issue-code",
            headers={"Authorization": "Bearer mock-secret-bearer-token"},
            json={
                "open_user_id": "OpenQQ:4019653007-49E8BA38D36D9D9EEEA615D657C94E2C",
                "open_group_id": "OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91",
                "sealdice_nickname": "雨霁",
                "bot_app_id": "1903175013",
            },
        )
        self.assertEqual(issued.status_code, 200)
        code = issued.json()["code"]
        self.assertRegex(code, r"^\d{6}$")

        inspected = browser.post("/api/auth/bot/inspect", json={"code": code})
        self.assertEqual(inspected.status_code, 200)
        ins_data = inspected.json()
        self.assertTrue(ins_data["is_new_user"])
        self.assertEqual(ins_data["suggested_nickname"], "雨霁")
        self.assertEqual(
            ins_data["avatar_url"],
            "https://q.qlogo.cn/qqapp/1903175013/49E8BA38D36D9D9EEEA615D657C94E2C/640",
        )

        # 3. 首次登录可自定义修改昵称并建立会话；同一短码二次消费返回 409
        consumed = browser.post("/api/auth/bot/consume", json={"code": code, "nickname": "雨霁·改"})
        self.assertEqual(consumed.status_code, 200)
        viewer = consumed.json()
        self.assertTrue(viewer["authenticated"])
        self.assertEqual(viewer["user"]["qq"], "oid:49E8BA38D36D9D9EEEA615D657C94E2C")
        self.assertEqual(viewer["user"]["nickname"], "雨霁·改")

        replay = TestClient(app_main.app).post("/api/auth/bot/consume", json={"code": code, "nickname": "重放"})
        self.assertEqual(replay.status_code, 409)

        # 4. 登录后发帖与自助修改昵称（历史帖子作者名同步更新）
        post_res = browser.post("/api/community/board", json={"title": "主群短码测试帖", "content": "无需角色卡即可发帖"})
        self.assertEqual(post_res.status_code, 200)
        post_id = post_res.json()["post"]["id"]

        renamed = browser.patch("/api/auth/profile", json={"nickname": "雨霁终章"})
        self.assertEqual(renamed.status_code, 200)
        self.assertEqual(renamed.json()["user"]["nickname"], "雨霁终章")

        thread = browser.get(f"/api/community/threads/{post_id}").json()["thread"]
        self.assertEqual(thread["author"]["nickname"], "雨霁终章")

        # 5. 维护者 OpenID 白名单自动享有管理员身份，且后台支持封禁/解封 oid:... 账号
        admin_issue = bot_bridge.post(
            "/api/auth/bot/issue-code",
            headers={"Authorization": "Bearer mock-secret-bearer-token"},
            json={
                "open_user_id": "OpenQQ:4019653007-FF834522288AB2D2993970BE14970D3A",
                "open_group_id": "OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91",
                "sealdice_nickname": "牛奶泡LanStarD",
            },
        ).json()
        admin_browser = TestClient(app_main.app)
        admin_login = admin_browser.post(
            "/api/auth/bot/consume",
            json={"code": admin_issue["code"], "nickname": "牛奶泡LanStarD"},
        ).json()
        self.assertTrue(admin_login["is_moderator"])
        self.assertEqual(admin_login["user"]["badge"], "维护者")

        ban_oid = admin_browser.post(
            "/api/community/moderation/ban",
            json={"qq": "oid:49E8BA38D36D9D9EEEA615D657C94E2C", "duration_hours": 24, "reason": "OpenID封禁测试"},
        )
        self.assertEqual(ban_oid.status_code, 200)
        blocked = browser.post("/api/community/board", json={"title": "被禁后发帖", "content": "测试拦截"})
        self.assertEqual(blocked.status_code, 403)


if __name__ == "__main__":
    unittest.main()
