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
from app.community import CommunityStore, UpstreamAuthClient


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
                    "bot_qq": "3239195029",
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
                        "bot_qq": "3239195029",
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

        # 4. 正常创建后未私聊直接 consume 返回 409 且不登录
        r = client_a.post("/api/auth/qq/challenges", json={"qq": "12345678"})
        self.assertEqual(r.status_code, 200)
        self.assertTrue(r.json()["verification_code"].startswith("TRPG-"))
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


if __name__ == "__main__":
    unittest.main()
