"""QQ 私聊验证码登录、持久会话与社区留言板 / 资源评论模块。

严格遵循《TRPG 私聊验证码登录接口》安全契约：
1. 网站后端持有 ONEBOT_SEQUENCE_TOKEN，前端永不可见该密钥。
2. challenge_id 与申请绑定的 QQ 仅保存在服务端 SQLite 会话表中（通过 HttpOnly Cookie 绑定当前浏览器），
   轮询和消费时绝不接受浏览器传入任意 challenge_id 或其他 QQ。
3. verified 只代表本人已私聊发码；必须调用一次性 consume 接口且同时满足：
   HTTP 200、passed === True、stale === False、auth_type == "qq_private_code"、
   status == "consumed"、group_id == 672204828、qq 与 challenge_id 与服务端会话一致，才建立登录态。
4. 全部鉴权与写接口响应头设置 Cache-Control: private, no-store。
"""
from __future__ import annotations

import hashlib
import hmac
import json
import os
import re
import secrets
import sqlite3
import time
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Iterator, Literal, Optional

import httpx
from fastapi import APIRouter, Depends, HTTPException, Query, Request, Response
from pydantic import BaseModel, Field

QQ_RE = re.compile(r"^[1-9][0-9]{4,11}$")
EXPECTED_GROUP_ID = 672204828
EXPECTED_BOT_QQ = "3239195029"
DEFAULT_AUTH_BASE = "https://trpg.chordvers.org/onebot-sequences"

BROWSER_COOKIE = "seqlib_community_sid"
BROWSER_TTL_SECONDS = 30 * 86400  # 30 天持久登录会话
PENDING_TTL_SECONDS = 600
COMMENT_COOLDOWN_SECONDS = 3

ERROR_MESSAGES: dict[str, str] = {
    "invalid_request": "QQ 号或请求参数格式不合法。",
    "unauthorized": "网站后端访问密钥错误或未配置，请联系站点维护者检查 ONEBOT_SEQUENCE_TOKEN。",
    "not_in_group": "该 QQ 当前不在宏观界域主群（672204828），无法获取验证码或登录。",
    "no_uploaded_card": "该 QQ 尚未在群文件指定目录上传过 .xlsx 角色卡，暂未通过发言准入核验。",
    "rate_limited": "验证码请求过于频繁（同 QQ 5 分钟最多 5 次），请稍后再试。",
    "member_cache_unavailable": "群成员与角色卡缓存正在同步或暂不可用，请稍后重试。",
    "receiver_unavailable": "BOT 私聊接收端当前离线，暂不发放新验证码，请稍后再试。",
    "verification_not_enabled": "私聊验证码服务当前未启用。",
    "auth_unavailable": "验证服务暂时不可用，请稍后再试。",
    "challenge_not_found": "验证请求不存在或已失效，请重新获取验证码。",
    "verification_pending": "尚未收到本人私聊发送的验证码，请先向 BOT 发送完整验证码。",
    "already_consumed": "该验证码已被使用，请重新申请新验证码。",
    "challenge_expired": "验证码已过期（有效期 5 分钟），请重新获取新验证码。",
    "no_login_attempt": "当前浏览器没有待处理的登录验证，请先输入 QQ 获取验证码。",
    "unable_to_verify": "验证服务响应超时或校验未通过，为安全起见未放行；若验证码已消费请重新申请新码。",
}


def comment_cooldown_seconds() -> int:
    try:
        return max(0, int(os.getenv("COMMUNITY_COOLDOWN_SECONDS", str(COMMENT_COOLDOWN_SECONDS))))
    except ValueError:
        return COMMENT_COOLDOWN_SECONDS


def now_ts() -> int:
    return int(time.time())


def ts_to_iso(ts: int | float | None) -> str | None:
    if ts is None:
        return None
    return datetime.fromtimestamp(float(ts), tz=timezone.utc).isoformat()


def mask_qq(qq: str) -> str:
    if not qq or len(qq) < 5:
        return qq
    if len(qq) <= 7:
        return f"{qq[:2]}***{qq[-2:]}"
    return f"{qq[:3]}****{qq[-3:]}"


def parse_admin_qqs() -> set[str]:
    raw = os.getenv("COMMUNITY_ADMIN_QQS", "")
    result: set[str] = set()
    for part in re.split(r"[,;\s]+", raw):
        cleaned = part.strip()
        if QQ_RE.fullmatch(cleaned):
            result.add(cleaned)
    return result


def clean_comment_content(raw: str, *, max_len: int = 1200) -> str:
    # 去除不可见控制字符（保留换行与制表符），折叠过多连续空行
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", raw or "")
    cleaned = cleaned.replace("\r\n", "\n").replace("\r", "\n").strip()
    cleaned = re.sub(r"\n{4,}", "\n\n\n", cleaned)
    if not cleaned:
        raise ValueError("内容不能为空")
    if len(cleaned) > max_len:
        raise ValueError(f"内容超出长度上限（最多 {max_len} 字）")
    return cleaned


class UpstreamAuthClient:
    """封装对 https://trpg.chordvers.org/onebot-sequences 的后端安全调用。"""

    def __init__(
        self,
        *,
        base_url: str | None = None,
        token_getter: Callable[[], str | None] | None = None,
        transport: httpx.BaseTransport | None = None,
        timeout: float = 5.0,
    ) -> None:
        self.base_url = (base_url or os.getenv("ONEBOT_SEQUENCE_BASE_URL") or DEFAULT_AUTH_BASE).rstrip("/")
        self._token_getter = token_getter or (lambda: os.getenv("ONEBOT_SEQUENCE_TOKEN"))
        self._transport = transport
        self.timeout = timeout

    def configured(self) -> bool:
        return bool((self._token_getter() or "").strip())

    def request(self, path: str, *, method: str = "GET", body: dict[str, Any] | None = None) -> tuple[int, dict[str, Any]]:
        token = (self._token_getter() or "").strip()
        if not token:
            raise HTTPException(503, ERROR_MESSAGES["unauthorized"])
        if "?" in path:
            raise ValueError("上游验证码接口不允许携带 URL 查询参数")
        headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
            "Cache-Control": "no-store",
        }
        content_bytes: bytes | None = None
        if body is not None:
            raw_json = json.dumps(body, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
            if len(raw_json) > 2048:
                raise HTTPException(400, "请求正文超出上限")
            headers["Content-Type"] = "application/json"
            content_bytes = raw_json

        url = f"{self.base_url}{path}"
        try:
            with httpx.Client(
                timeout=self.timeout,
                follow_redirects=False,
                transport=self._transport,
            ) as client:
                resp = client.request(method=method, url=url, headers=headers, content=content_bytes)
        except httpx.HTTPError as exc:
            raise RuntimeError("upstream_network_error") from exc

        if 300 <= resp.status_code < 400:
            raise RuntimeError("upstream_unexpected_redirect")

        try:
            data = resp.json()
            if not isinstance(data, dict):
                data = {}
        except Exception:
            data = {}
        return resp.status_code, data


class CommunityStore:
    """SQLite 持久化存储：浏览器会话、登录成员档案、封禁名单、留言板与资源评论、点赞共鸣。"""

    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    @contextmanager
    def _conn(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(str(self.db_path), timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def _init_db(self) -> None:
        with self._conn() as conn:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS community_users (
                    qq TEXT PRIMARY KEY,
                    nickname TEXT NOT NULL,
                    avatar_url TEXT NOT NULL DEFAULT '',
                    card_count INTEGER NOT NULL DEFAULT 0,
                    group_id INTEGER NOT NULL DEFAULT 672204828,
                    role TEXT NOT NULL DEFAULT 'member',
                    banned_until INTEGER NOT NULL DEFAULT 0,
                    ban_reason TEXT NOT NULL DEFAULT '',
                    first_login_at INTEGER NOT NULL,
                    last_login_at INTEGER NOT NULL,
                    updated_at INTEGER NOT NULL
                );

                CREATE TABLE IF NOT EXISTS browser_sessions (
                    sid_hash TEXT PRIMARY KEY,
                    user_qq TEXT,
                    pending_challenge_id TEXT,
                    pending_qq TEXT,
                    pending_bot_qq TEXT,
                    pending_expires_at TEXT,
                    pending_created_at INTEGER,
                    created_at INTEGER NOT NULL,
                    updated_at INTEGER NOT NULL,
                    expires_at INTEGER NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_browser_sessions_user ON browser_sessions(user_qq);
                CREATE INDEX IF NOT EXISTS idx_browser_sessions_exp ON browser_sessions(expires_at);

                CREATE TABLE IF NOT EXISTS community_posts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    scope TEXT NOT NULL CHECK(scope IN ('board', 'resource')),
                    resource_path TEXT NOT NULL DEFAULT '',
                    resource_title TEXT NOT NULL DEFAULT '',
                    anchor_id TEXT NOT NULL DEFAULT '',
                    anchor_label TEXT NOT NULL DEFAULT '',
                    parent_id INTEGER,
                    reply_to_id INTEGER,
                    reply_to_qq TEXT NOT NULL DEFAULT '',
                    reply_to_nickname TEXT NOT NULL DEFAULT '',
                    author_qq TEXT NOT NULL,
                    author_nickname TEXT NOT NULL,
                    author_avatar TEXT NOT NULL DEFAULT '',
                    author_role_badge TEXT NOT NULL DEFAULT '',
                    content TEXT NOT NULL,
                    is_pinned INTEGER NOT NULL DEFAULT 0,
                    like_count INTEGER NOT NULL DEFAULT 0,
                    status TEXT NOT NULL DEFAULT 'active' CHECK(status IN ('active', 'deleted')),
                    deleted_by TEXT NOT NULL DEFAULT '',
                    created_at INTEGER NOT NULL,
                    updated_at INTEGER NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_posts_board ON community_posts(scope, status, is_pinned DESC, created_at DESC, id DESC);
                CREATE INDEX IF NOT EXISTS idx_posts_resource ON community_posts(scope, resource_path, status, is_pinned DESC, created_at ASC, id ASC);
                CREATE INDEX IF NOT EXISTS idx_posts_parent ON community_posts(parent_id, status, created_at ASC);
                CREATE INDEX IF NOT EXISTS idx_posts_author ON community_posts(author_qq, created_at DESC);

                CREATE TABLE IF NOT EXISTS community_post_likes (
                    post_id INTEGER NOT NULL,
                    user_qq TEXT NOT NULL,
                    created_at INTEGER NOT NULL,
                    PRIMARY KEY (post_id, user_qq),
                    FOREIGN KEY (post_id) REFERENCES community_posts(id) ON DELETE CASCADE
                );
                """
            )

    @staticmethod
    def _hash_sid(raw_sid: str, secret: bytes) -> str:
        return hmac.new(secret, raw_sid.encode("utf-8"), hashlib.sha256).hexdigest()

    def get_or_create_session(self, raw_sid: str | None, secret: bytes) -> tuple[str, dict[str, Any], bool]:
        now = now_ts()
        with self._conn() as conn:
            if raw_sid and re.fullmatch(r"[A-Za-z0-9_-]{24,120}", raw_sid):
                sid_hash = self._hash_sid(raw_sid, secret)
                row = conn.execute(
                    "SELECT * FROM browser_sessions WHERE sid_hash = ? AND expires_at > ?",
                    (sid_hash, now),
                ).fetchone()
                if row:
                    return raw_sid, dict(row), False

            new_sid = secrets.token_urlsafe(32)
            sid_hash = self._hash_sid(new_sid, secret)
            exp = now + BROWSER_TTL_SECONDS
            conn.execute(
                """
                INSERT INTO browser_sessions (sid_hash, created_at, updated_at, expires_at)
                VALUES (?, ?, ?, ?)
                """,
                (sid_hash, now, now, exp),
            )
            # 顺带清理过期很久的无用匿名会话
            conn.execute("DELETE FROM browser_sessions WHERE expires_at <= ?", (now - 3600,))
            row = conn.execute("SELECT * FROM browser_sessions WHERE sid_hash = ?", (sid_hash,)).fetchone()
            return new_sid, dict(row), True

    def get_session(self, raw_sid: str | None, secret: bytes) -> dict[str, Any] | None:
        if not raw_sid or not re.fullmatch(r"[A-Za-z0-9_-]{24,120}", raw_sid):
            return None
        now = now_ts()
        sid_hash = self._hash_sid(raw_sid, secret)
        with self._conn() as conn:
            row = conn.execute(
                "SELECT * FROM browser_sessions WHERE sid_hash = ? AND expires_at > ?",
                (sid_hash, now),
            ).fetchone()
            return dict(row) if row else None

    def save_pending_challenge(
        self,
        raw_sid: str,
        secret: bytes,
        *,
        challenge_id: str,
        qq: str,
        bot_qq: str,
        expires_at: str,
    ) -> None:
        now = now_ts()
        sid_hash = self._hash_sid(raw_sid, secret)
        with self._conn() as conn:
            conn.execute(
                """
                UPDATE browser_sessions
                SET pending_challenge_id = ?,
                    pending_qq = ?,
                    pending_bot_qq = ?,
                    pending_expires_at = ?,
                    pending_created_at = ?,
                    updated_at = ?,
                    expires_at = ?
                WHERE sid_hash = ?
                """,
                (challenge_id, qq, bot_qq, expires_at, now, now, now + BROWSER_TTL_SECONDS, sid_hash),
            )

    def clear_pending_challenge(self, raw_sid: str, secret: bytes) -> None:
        now = now_ts()
        sid_hash = self._hash_sid(raw_sid, secret)
        with self._conn() as conn:
            conn.execute(
                """
                UPDATE browser_sessions
                SET pending_challenge_id = NULL,
                    pending_qq = NULL,
                    pending_bot_qq = NULL,
                    pending_expires_at = NULL,
                    pending_created_at = NULL,
                    updated_at = ?
                WHERE sid_hash = ?
                """,
                (now, sid_hash),
            )

    def bind_logged_in_user(
        self,
        raw_sid: str,
        secret: bytes,
        *,
        qq: str,
        nickname: str,
        avatar_url: str,
        card_count: int,
        group_id: int,
    ) -> dict[str, Any]:
        now = now_ts()
        sid_hash = self._hash_sid(raw_sid, secret)
        safe_nick = (nickname or "").strip() or f"调查员{qq[-4:]}"
        safe_avatar = (avatar_url or "").strip() or f"https://q1.qlogo.cn/g?b=qq&nk={qq}&s=640"
        with self._conn() as conn:
            existing = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE community_users
                    SET nickname = ?,
                        avatar_url = ?,
                        card_count = ?,
                        group_id = ?,
                        last_login_at = ?,
                        updated_at = ?
                    WHERE qq = ?
                    """,
                    (safe_nick, safe_avatar, int(card_count or 0), int(group_id), now, now, qq),
                )
            else:
                conn.execute(
                    """
                    INSERT INTO community_users (
                        qq, nickname, avatar_url, card_count, group_id, role,
                        banned_until, ban_reason, first_login_at, last_login_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, 'member', 0, '', ?, ?, ?)
                    """,
                    (qq, safe_nick, safe_avatar, int(card_count or 0), int(group_id), now, now, now),
                )

            conn.execute(
                """
                UPDATE browser_sessions
                SET user_qq = ?,
                    pending_challenge_id = NULL,
                    pending_qq = NULL,
                    pending_bot_qq = NULL,
                    pending_expires_at = NULL,
                    pending_created_at = NULL,
                    updated_at = ?,
                    expires_at = ?
                WHERE sid_hash = ?
                """,
                (qq, now, now + BROWSER_TTL_SECONDS, sid_hash),
            )
            user_row = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            return dict(user_row)

    def logout_session(self, raw_sid: str | None, secret: bytes) -> None:
        if not raw_sid:
            return
        sid_hash = self._hash_sid(raw_sid, secret)
        with self._conn() as conn:
            conn.execute("DELETE FROM browser_sessions WHERE sid_hash = ?", (sid_hash,))

    def get_user(self, qq: str | None) -> dict[str, Any] | None:
        if not qq:
            return None
        with self._conn() as conn:
            row = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            return dict(row) if row else None

    def set_user_ban(self, qq: str, *, banned_until: int, reason: str = "", nickname_hint: str = "") -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            existing = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            if existing:
                conn.execute(
                    """
                    UPDATE community_users
                    SET banned_until = ?, ban_reason = ?, updated_at = ?
                    WHERE qq = ?
                    """,
                    (banned_until, reason.strip(), now, qq),
                )
            else:
                nick = nickname_hint.strip() or f"QQ {qq}"
                avatar = f"https://q1.qlogo.cn/g?b=qq&nk={qq}&s=640"
                conn.execute(
                    """
                    INSERT INTO community_users (
                        qq, nickname, avatar_url, card_count, group_id, role,
                        banned_until, ban_reason, first_login_at, last_login_at, updated_at
                    ) VALUES (?, ?, ?, 0, ?, 'member', ?, ?, ?, 0, ?)
                    """,
                    (qq, nick, avatar, EXPECTED_GROUP_ID, banned_until, reason.strip(), now, now),
                )
            # 若永久或当前处于封禁中，可保留会话让其看到被禁言原因，发言时拦截
            row = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            return dict(row)

    def set_user_role(self, qq: str, role: Literal["member", "moderator"]) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            existing = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            if not existing:
                raise HTTPException(404, "该 QQ 尚未登录过本站")
            conn.execute("UPDATE community_users SET role = ?, updated_at = ? WHERE qq = ?", (role, now, qq))
            row = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            return dict(row)

    def _serialize_post(
        self,
        row: sqlite3.Row | dict[str, Any],
        *,
        viewer_qq: str | None = None,
        viewer_is_mod: bool = False,
        liked_ids: set[int] | None = None,
        include_full_qq: bool = False,
    ) -> dict[str, Any]:
        d = dict(row)
        pid = int(d["id"])
        author_qq = str(d["author_qq"])
        reply_qq = str(d.get("reply_to_qq") or "")
        is_mine = bool(viewer_qq and viewer_qq == author_qq)
        item: dict[str, Any] = {
            "id": pid,
            "scope": d["scope"],
            "resource_path": d["resource_path"],
            "resource_title": d["resource_title"],
            "anchor_id": d["anchor_id"],
            "anchor_label": d["anchor_label"],
            "parent_id": d["parent_id"],
            "reply_to_id": d["reply_to_id"],
            "reply_to_nickname": d["reply_to_nickname"],
            "author": {
                "qq_masked": mask_qq(author_qq),
                "qq": author_qq if (include_full_qq or viewer_is_mod or is_mine) else None,
                "nickname": d["author_nickname"],
                "avatar_url": d["author_avatar"] or f"https://q1.qlogo.cn/g?b=qq&nk={author_qq}&s=640",
                "badge": d["author_role_badge"] or "",
            },
            "content": d["content"] if d["status"] == "active" else "",
            "is_pinned": bool(d["is_pinned"]),
            "like_count": int(d["like_count"] or 0),
            "liked_by_me": bool(liked_ids and pid in liked_ids),
            "can_delete": bool(d["status"] == "active" and (is_mine or viewer_is_mod)),
            "can_pin": bool(d["status"] == "active" and d["parent_id"] is None and viewer_is_mod),
            "status": d["status"],
            "created_at": ts_to_iso(d["created_at"]),
            "created_ts": int(d["created_at"]),
        }
        if reply_qq and (include_full_qq or viewer_is_mod):
            item["reply_to_qq"] = reply_qq
        return item

    def _liked_post_ids(self, conn: sqlite3.Connection, viewer_qq: str | None, post_ids: list[int]) -> set[int]:
        if not viewer_qq or not post_ids:
            return set()
        placeholders = ",".join("?" for _ in post_ids)
        rows = conn.execute(
            f"SELECT post_id FROM community_post_likes WHERE user_qq = ? AND post_id IN ({placeholders})",
            [viewer_qq, *post_ids],
        ).fetchall()
        return {int(r["post_id"]) for r in rows}

    def create_post(
        self,
        *,
        scope: Literal["board", "resource"],
        author: dict[str, Any],
        author_badge: str,
        content: str,
        resource_path: str = "",
        resource_title: str = "",
        anchor_id: str = "",
        anchor_label: str = "",
        reply_to_id: int | None = None,
        bypass_cooldown: bool = False,
    ) -> dict[str, Any]:
        cleaned = clean_comment_content(content, max_len=1200)
        now = now_ts()
        author_qq = str(author["qq"])

        with self._conn() as conn:
            cooldown = comment_cooldown_seconds()
            if not bypass_cooldown and cooldown > 0:
                recent = conn.execute(
                    "SELECT created_at FROM community_posts WHERE author_qq = ? ORDER BY id DESC LIMIT 1",
                    (author_qq,),
                ).fetchone()
                if recent and now - int(recent["created_at"]) < cooldown:
                    raise HTTPException(429, "发言太快啦，请稍等几秒再发送。")

            parent_id: int | None = None
            actual_reply_id: int | None = None
            reply_to_qq = ""
            reply_to_nickname = ""

            if reply_to_id is not None:
                target = conn.execute(
                    "SELECT * FROM community_posts WHERE id = ? AND status = 'active'",
                    (int(reply_to_id),),
                ).fetchone()
                if not target:
                    raise HTTPException(404, "要回复的留言或评论不存在或已被删除")
                if target["scope"] != scope or (scope == "resource" and target["resource_path"] != resource_path):
                    raise HTTPException(400, "不能跨版块或跨档案回复评论")
                # 单层楼中楼归并：如果回复的是子回复，parent_id 归属到根楼层，同时保留 reply_to_nickname
                parent_id = int(target["parent_id"]) if target["parent_id"] is not None else int(target["id"])
                actual_reply_id = int(target["id"])
                reply_to_qq = str(target["author_qq"])
                reply_to_nickname = str(target["author_nickname"])
                if not anchor_id and target["anchor_id"]:
                    anchor_id = str(target["anchor_id"])
                    anchor_label = str(target["anchor_label"])

            cursor = conn.execute(
                """
                INSERT INTO community_posts (
                    scope, resource_path, resource_title, anchor_id, anchor_label,
                    parent_id, reply_to_id, reply_to_qq, reply_to_nickname,
                    author_qq, author_nickname, author_avatar, author_role_badge,
                    content, is_pinned, like_count, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, 'active', ?, ?)
                """,
                (
                    scope,
                    resource_path,
                    resource_title,
                    anchor_id.strip()[:80],
                    anchor_label.strip()[:120],
                    parent_id,
                    actual_reply_id,
                    reply_to_qq,
                    reply_to_nickname,
                    author_qq,
                    str(author["nickname"]),
                    str(author.get("avatar_url") or ""),
                    author_badge,
                    cleaned,
                    now,
                    now,
                ),
            )
            post_id = int(cursor.lastrowid)
            row = conn.execute("SELECT * FROM community_posts WHERE id = ?", (post_id,)).fetchone()
            return self._serialize_post(row, viewer_qq=author_qq, viewer_is_mod=bool(author_badge))

    def list_threads(
        self,
        *,
        scope: Literal["board", "resource"],
        resource_path: str = "",
        anchor_id: str = "",
        limit: int = 30,
        offset: int = 0,
        sort: Literal["latest", "oldest", "hot"] = "latest",
        viewer_qq: str | None = None,
        viewer_is_mod: bool = False,
    ) -> dict[str, Any]:
        limit = max(1, min(limit, 100))
        offset = max(0, offset)
        where = ["scope = ?", "status = 'active'", "parent_id IS NULL"]
        params: list[Any] = [scope]
        if scope == "resource":
            where.append("resource_path = ?")
            params.append(resource_path)
            if anchor_id:
                where.append("anchor_id = ?")
                params.append(anchor_id)

        where_sql = " AND ".join(where)
        if sort == "oldest":
            order_sql = "is_pinned DESC, created_at ASC, id ASC"
        elif sort == "hot":
            order_sql = "is_pinned DESC, like_count DESC, created_at DESC, id DESC"
        else:
            order_sql = "is_pinned DESC, created_at DESC, id DESC"

        with self._conn() as conn:
            total_roots = int(
                conn.execute(f"SELECT COUNT(*) AS c FROM community_posts WHERE {where_sql}", params).fetchone()["c"]
            )
            if scope == "resource":
                total_all = int(
                    conn.execute(
                        "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'resource' AND resource_path = ? AND status = 'active'",
                        (resource_path,),
                    ).fetchone()["c"]
                )
                anchor_rows = conn.execute(
                    """
                    SELECT anchor_id, anchor_label, COUNT(*) AS c
                    FROM community_posts
                    WHERE scope = 'resource' AND resource_path = ? AND status = 'active' AND anchor_id != ''
                    GROUP BY anchor_id
                    """,
                    (resource_path,),
                ).fetchall()
                anchor_counts = {
                    str(r["anchor_id"]): {"count": int(r["c"]), "label": str(r["anchor_label"] or "")}
                    for r in anchor_rows
                }
            else:
                total_all = int(
                    conn.execute(
                        "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'board' AND status = 'active'"
                    ).fetchone()["c"]
                )
                anchor_counts = {}

            root_rows = conn.execute(
                f"SELECT * FROM community_posts WHERE {where_sql} ORDER BY {order_sql} LIMIT ? OFFSET ?",
                [*params, limit, offset],
            ).fetchall()

            root_ids = [int(r["id"]) for r in root_rows]
            reply_rows: list[sqlite3.Row] = []
            if root_ids:
                placeholders = ",".join("?" for _ in root_ids)
                reply_rows = conn.execute(
                    f"""
                    SELECT * FROM community_posts
                    WHERE status = 'active' AND parent_id IN ({placeholders})
                    ORDER BY created_at ASC, id ASC
                    """,
                    root_ids,
                ).fetchall()

            all_ids = root_ids + [int(r["id"]) for r in reply_rows]
            liked_ids = self._liked_post_ids(conn, viewer_qq, all_ids)

            replies_by_parent: dict[int, list[dict[str, Any]]] = {}
            for r in reply_rows:
                pid = int(r["parent_id"])
                replies_by_parent.setdefault(pid, []).append(
                    self._serialize_post(
                        r,
                        viewer_qq=viewer_qq,
                        viewer_is_mod=viewer_is_mod,
                        liked_ids=liked_ids,
                    )
                )

            items = []
            for r in root_rows:
                item = self._serialize_post(
                    r,
                    viewer_qq=viewer_qq,
                    viewer_is_mod=viewer_is_mod,
                    liked_ids=liked_ids,
                )
                replies = replies_by_parent.get(int(r["id"]), [])
                item["replies"] = replies
                item["reply_count"] = len(replies)
                items.append(item)

            return {
                "items": items,
                "count": len(items),
                "total_threads": total_roots,
                "total_posts": total_all,
                "anchor_counts": anchor_counts,
                "limit": limit,
                "offset": offset,
            }

    def list_recent_resource_comments(
        self,
        *,
        limit: int = 16,
        viewer_qq: str | None = None,
        viewer_is_mod: bool = False,
    ) -> list[dict[str, Any]]:
        limit = max(1, min(limit, 50))
        with self._conn() as conn:
            rows = conn.execute(
                """
                SELECT * FROM community_posts
                WHERE scope = 'resource' AND status = 'active'
                ORDER BY created_at DESC, id DESC
                LIMIT ?
                """,
                (limit,),
            ).fetchall()
            post_ids = [int(r["id"]) for r in rows]
            liked_ids = self._liked_post_ids(conn, viewer_qq, post_ids)
            return [
                self._serialize_post(
                    r,
                    viewer_qq=viewer_qq,
                    viewer_is_mod=viewer_is_mod,
                    liked_ids=liked_ids,
                )
                for r in rows
            ]

    def toggle_like(self, *, post_id: int, user_qq: str) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            post = conn.execute(
                "SELECT id, status FROM community_posts WHERE id = ?",
                (post_id,),
            ).fetchone()
            if not post or post["status"] != "active":
                raise HTTPException(404, "该留言或评论不存在")

            existing = conn.execute(
                "SELECT 1 FROM community_post_likes WHERE post_id = ? AND user_qq = ?",
                (post_id, user_qq),
            ).fetchone()
            if existing:
                conn.execute(
                    "DELETE FROM community_post_likes WHERE post_id = ? AND user_qq = ?",
                    (post_id, user_qq),
                )
                liked = False
            else:
                conn.execute(
                    "INSERT INTO community_post_likes (post_id, user_qq, created_at) VALUES (?, ?, ?)",
                    (post_id, user_qq, now),
                )
                liked = True

            count = int(
                conn.execute(
                    "SELECT COUNT(*) AS c FROM community_post_likes WHERE post_id = ?",
                    (post_id,),
                ).fetchone()["c"]
            )
            conn.execute(
                "UPDATE community_posts SET like_count = ?, updated_at = ? WHERE id = ?",
                (count, now, post_id),
            )
            return {"ok": True, "post_id": post_id, "liked": liked, "like_count": count}

    def set_pinned(self, *, post_id: int, pinned: bool) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            post = conn.execute("SELECT * FROM community_posts WHERE id = ?", (post_id,)).fetchone()
            if not post or post["status"] != "active":
                raise HTTPException(404, "该留言或评论不存在")
            if post["parent_id"] is not None:
                raise HTTPException(400, "只能置顶主楼层，不能置顶楼中楼回复")
            conn.execute(
                "UPDATE community_posts SET is_pinned = ?, updated_at = ? WHERE id = ?",
                (1 if pinned else 0, now, post_id),
            )
            updated = conn.execute("SELECT * FROM community_posts WHERE id = ?", (post_id,)).fetchone()
            return self._serialize_post(updated, viewer_is_mod=True, include_full_qq=True)

    def delete_post(self, *, post_id: int, actor: str, is_mod: bool, viewer_qq: str | None) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            post = conn.execute("SELECT * FROM community_posts WHERE id = ?", (post_id,)).fetchone()
            if not post:
                raise HTTPException(404, "该留言或评论不存在")
            if post["status"] == "deleted":
                return {"ok": True, "post_id": post_id, "already_deleted": True}
            if not is_mod and (not viewer_qq or str(post["author_qq"]) != viewer_qq):
                raise HTTPException(403, "只能删除自己发表的留言或评论")

            # 删除主楼时，其楼中楼一并标记删除，避免悬空孤儿回复
            conn.execute(
                """
                UPDATE community_posts
                SET status = 'deleted', deleted_by = ?, is_pinned = 0, updated_at = ?
                WHERE id = ? OR parent_id = ?
                """,
                (actor, now, post_id, post_id),
            )
            return {"ok": True, "post_id": post_id}

    def stats_overview(self) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            board_count = int(
                conn.execute(
                    "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'board' AND status = 'active'"
                ).fetchone()["c"]
            )
            resource_comment_count = int(
                conn.execute(
                    "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'resource' AND status = 'active'"
                ).fetchone()["c"]
            )
            member_count = int(conn.execute("SELECT COUNT(*) AS c FROM community_users").fetchone()["c"])
            banned_count = int(
                conn.execute(
                    "SELECT COUNT(*) AS c FROM community_users WHERE banned_until > ? OR banned_until = -1",
                    (now,),
                ).fetchone()["c"]
            )
            return {
                "board_post_count": board_count,
                "resource_comment_count": resource_comment_count,
                "member_count": member_count,
                "banned_count": banned_count,
            }

    def admin_list_posts(
        self,
        *,
        scope: Literal["all", "board", "resource"] = "all",
        status: Literal["active", "deleted", "all"] = "active",
        q: str = "",
        qq: str = "",
        limit: int = 50,
        offset: int = 0,
    ) -> dict[str, Any]:
        limit = max(1, min(limit, 200))
        offset = max(0, offset)
        where = ["1=1"]
        params: list[Any] = []
        if scope in ("board", "resource"):
            where.append("scope = ?")
            params.append(scope)
        if status in ("active", "deleted"):
            where.append("status = ?")
            params.append(status)
        if qq.strip():
            where.append("author_qq = ?")
            params.append(qq.strip())
        if q.strip():
            needle = f"%{q.strip()}%"
            where.append("(content LIKE ? OR author_nickname LIKE ? OR resource_title LIKE ? OR author_qq LIKE ?)")
            params.extend([needle, needle, needle, needle])

        where_sql = " AND ".join(where)
        with self._conn() as conn:
            total = int(
                conn.execute(f"SELECT COUNT(*) AS c FROM community_posts WHERE {where_sql}", params).fetchone()["c"]
            )
            rows = conn.execute(
                f"""
                SELECT * FROM community_posts
                WHERE {where_sql}
                ORDER BY created_at DESC, id DESC
                LIMIT ? OFFSET ?
                """,
                [*params, limit, offset],
            ).fetchall()
            items = []
            for r in rows:
                d = dict(r)
                item = self._serialize_post(d, viewer_is_mod=True, include_full_qq=True)
                # 管理台即使已删除也允许审计原文与删除操作者
                item["raw_content"] = d["content"]
                item["deleted_by"] = d["deleted_by"]
                items.append(item)
            return {"items": items, "count": len(items), "total": total, "limit": limit, "offset": offset}

    def admin_list_users(
        self,
        *,
        filter_mode: Literal["all", "banned", "moderator"] = "all",
        q: str = "",
        limit: int = 80,
        offset: int = 0,
    ) -> dict[str, Any]:
        now = now_ts()
        limit = max(1, min(limit, 200))
        offset = max(0, offset)
        where = ["1=1"]
        params: list[Any] = []
        if filter_mode == "banned":
            where.append("(banned_until = -1 OR banned_until > ?)")
            params.append(now)
        elif filter_mode == "moderator":
            where.append("role = 'moderator'")
        if q.strip():
            needle = f"%{q.strip()}%"
            where.append("(qq LIKE ? OR nickname LIKE ?)")
            params.extend([needle, needle])

        where_sql = " AND ".join(where)
        admin_qqs = parse_admin_qqs()
        with self._conn() as conn:
            total = int(
                conn.execute(f"SELECT COUNT(*) AS c FROM community_users WHERE {where_sql}", params).fetchone()["c"]
            )
            rows = conn.execute(
                f"""
                SELECT u.*,
                       (SELECT COUNT(*) FROM community_posts p WHERE p.author_qq = u.qq AND p.status = 'active') AS post_count
                FROM community_users u
                WHERE {where_sql}
                ORDER BY u.last_login_at DESC, u.updated_at DESC
                LIMIT ? OFFSET ?
                """,
                [*params, limit, offset],
            ).fetchall()
            items = []
            for r in rows:
                d = dict(r)
                banned_until = int(d["banned_until"] or 0)
                is_banned = banned_until == -1 or banned_until > now
                qq_val = str(d["qq"])
                items.append(
                    {
                        "qq": qq_val,
                        "nickname": d["nickname"],
                        "avatar_url": d["avatar_url"],
                        "card_count": int(d["card_count"] or 0),
                        "role": "env_admin" if qq_val in admin_qqs else d["role"],
                        "is_banned": is_banned,
                        "banned_until": None if banned_until == 0 else (-1 if banned_until == -1 else ts_to_iso(banned_until)),
                        "ban_reason": d["ban_reason"] if is_banned else "",
                        "post_count": int(d["post_count"] or 0),
                        "first_login_at": ts_to_iso(d["first_login_at"]),
                        "last_login_at": ts_to_iso(d["last_login_at"]) if d["last_login_at"] else None,
                    }
                )
            return {"items": items, "count": len(items), "total": total}


class BeginChallengeRequest(BaseModel):
    qq: str


class CreateBoardPostRequest(BaseModel):
    content: str = Field(min_length=1, max_length=1200)
    reply_to_id: Optional[int] = None


class CreateResourceCommentRequest(BaseModel):
    resource_path: str = Field(min_length=1, max_length=300)
    content: str = Field(min_length=1, max_length=1200)
    anchor_id: str = Field(default="", max_length=80)
    anchor_label: str = Field(default="", max_length=120)
    reply_to_id: Optional[int] = None


class PinPostRequest(BaseModel):
    pinned: bool


class AdminBanUserRequest(BaseModel):
    qq: str
    duration_hours: Optional[int] = Field(default=None, description="None 或 <=0 表示永久封禁；传 0 且 unban=True 解除封禁")
    unban: bool = False
    reason: str = Field(default="", max_length=200)
    nickname_hint: str = Field(default="", max_length=80)


class AdminRoleRequest(BaseModel):
    role: Literal["member", "moderator"]


class AdminAnnouncementRequest(BaseModel):
    content: str = Field(min_length=1, max_length=1200)
    nickname: str = Field(default="界域终端公告", max_length=40)
    pinned: bool = True


def build_community_router(
    *,
    store_getter: Callable[[], CommunityStore],
    upstream_getter: Callable[[], UpstreamAuthClient],
    secret_getter: Callable[[], bytes],
    is_admin_request: Callable[[Request], bool],
    resolve_public_resource: Callable[[str], tuple[str, str]],
) -> APIRouter:
    """构建挂载在 /api 下的社区与 QQ 登录路由器。"""
    router = APIRouter(tags=["community"])

    def no_store(response: Response) -> None:
        response.headers["Cache-Control"] = "private, no-store"

    def set_browser_cookie(response: Response, request: Request, sid: str) -> None:
        response.set_cookie(
            BROWSER_COOKIE,
            sid,
            max_age=BROWSER_TTL_SECONDS,
            httponly=True,
            samesite="lax",
            secure=request.url.scheme == "https",
            path="/",
        )

    def user_is_banned(user: dict[str, Any] | None) -> bool:
        if not user:
            return False
        until = int(user.get("banned_until") or 0)
        return until == -1 or until > now_ts()

    def user_is_qq_moderator(user: dict[str, Any] | None) -> bool:
        if not user or user_is_banned(user):
            return False
        qq = str(user["qq"])
        if qq in parse_admin_qqs():
            return True
        return user.get("role") == "moderator"

    def viewer_context(request: Request) -> tuple[dict[str, Any] | None, dict[str, Any] | None, bool, bool]:
        """返回 (session_row, user_row, is_admin_cookie, is_effective_moderator)"""
        raw_sid = request.cookies.get(BROWSER_COOKIE)
        sess = store_getter().get_session(raw_sid, secret_getter())
        user = store_getter().get_user(sess.get("user_qq")) if sess and sess.get("user_qq") else None
        admin_cookie = is_admin_request(request)
        is_mod = admin_cookie or user_is_qq_moderator(user)
        return sess, user, admin_cookie, is_mod

    def format_viewer(
        sess: dict[str, Any] | None,
        user: dict[str, Any] | None,
        *,
        admin_cookie: bool,
        is_mod: bool,
    ) -> dict[str, Any]:
        pending = None
        if sess and sess.get("pending_challenge_id") and sess.get("pending_qq"):
            pending = {
                "qq": sess["pending_qq"],
                "bot_qq": sess.get("pending_bot_qq") or EXPECTED_BOT_QQ,
                "expires_at": sess.get("pending_expires_at"),
            }
        if not user:
            return {
                "authenticated": False,
                "auth_configured": upstream_getter().configured(),
                "bot_qq": EXPECTED_BOT_QQ,
                "group_id": EXPECTED_GROUP_ID,
                "admin_authenticated": admin_cookie,
                "is_moderator": is_mod,
                "pending_challenge": pending,
                "user": None,
            }
        banned = user_is_banned(user)
        until = int(user.get("banned_until") or 0)
        qq_val = str(user["qq"])
        badge = "维护者" if qq_val in parse_admin_qqs() else ("协管" if user.get("role") == "moderator" else "")
        return {
            "authenticated": True,
            "auth_configured": upstream_getter().configured(),
            "bot_qq": EXPECTED_BOT_QQ,
            "group_id": EXPECTED_GROUP_ID,
            "admin_authenticated": admin_cookie,
            "is_moderator": is_mod,
            "pending_challenge": pending,
            "user": {
                "qq": qq_val,
                "qq_masked": mask_qq(qq_val),
                "nickname": user["nickname"],
                "avatar_url": user["avatar_url"] or f"https://q1.qlogo.cn/g?b=qq&nk={qq_val}&s=640",
                "card_count": int(user.get("card_count") or 0),
                "badge": badge,
                "is_banned": banned,
                "banned_until": None if not banned else (-1 if until == -1 else ts_to_iso(until)),
                "ban_reason": user.get("ban_reason") or "" if banned else "",
            },
        }

    def require_active_member(request: Request) -> tuple[dict[str, Any], bool, str]:
        """要求已通过 QQ 验证登录且未被封禁。返回 (user, is_mod, author_badge)。"""
        _sess, user, admin_cookie, is_mod = viewer_context(request)
        if not user:
            raise HTTPException(401, "请先使用群内 QQ 号完成私聊验证码登录后再发言")
        if user_is_banned(user):
            reason = (user.get("ban_reason") or "").strip()
            detail = f"当前账号已被限制发言（原因：{reason}）" if reason else "当前账号已被限制发言"
            raise HTTPException(403, detail)
        qq_val = str(user["qq"])
        badge = "维护者" if (qq_val in parse_admin_qqs() or admin_cookie) else ("协管" if user.get("role") == "moderator" else "")
        return user, is_mod, badge

    def require_moderator(request: Request) -> tuple[str, dict[str, Any] | None]:
        """支持 ADMIN_PASSWORD Cookie 或管理员 QQ 白名单/协管身份。"""
        _sess, user, admin_cookie, is_mod = viewer_context(request)
        if not is_mod:
            raise HTTPException(403, "需要管理员或社区协管权限")
        actor = f"qq:{user['qq']}" if user else "admin_console"
        return actor, user

    # ---------------- 1. QQ 私聊验证码三步登录接口 ----------------

    @router.get("/auth/me")
    def auth_me(request: Request, response: Response):
        no_store(response)
        sess, user, admin_cookie, is_mod = viewer_context(request)
        return format_viewer(sess, user, admin_cookie=admin_cookie, is_mod=is_mod)

    @router.post("/auth/qq/challenges")
    def auth_begin_qq_login(body: BeginChallengeRequest, request: Request, response: Response):
        no_store(response)
        if request.url.query:
            raise HTTPException(400, "创建验证请求不允许携带 URL 查询参数")
        proposed_qq = (body.qq or "").strip()
        if not QQ_RE.fullmatch(proposed_qq):
            raise HTTPException(400, "请输入有效的 QQ 号（5～12 位数字，不以 0 开头）")

        sid, _sess, _created = store_getter().get_or_create_session(
            request.cookies.get(BROWSER_COOKIE),
            secret_getter(),
        )
        set_browser_cookie(response, request, sid)

        try:
            http_status, data = upstream_getter().request(
                "/v1/auth/challenges",
                method="POST",
                body={"qq": proposed_qq},
            )
        except HTTPException:
            raise
        except Exception:
            raise HTTPException(503, ERROR_MESSAGES["auth_unavailable"])

        if http_status != 201:
            err_code = str(data.get("error") or "")
            msg = ERROR_MESSAGES.get(err_code, f"无法开始验证（HTTP {http_status}）")
            mapped_status = http_status if http_status in {400, 403, 429, 503} else 503
            raise HTTPException(mapped_status, msg)

        challenge_id = data.get("challenge_id")
        verification_code = data.get("verification_code")
        message_to_send = data.get("message_to_send") or verification_code
        bot_qq = str(data.get("bot_qq") or "")
        expires_at = str(data.get("expires_at") or "")

        if (
            data.get("qq") != proposed_qq
            or data.get("group_id") != EXPECTED_GROUP_ID
            or bot_qq != EXPECTED_BOT_QQ
            or data.get("status") != "pending"
            or not isinstance(challenge_id, str)
            or not challenge_id
            or not isinstance(verification_code, str)
            or not verification_code
        ):
            raise HTTPException(502, "验证服务返回数据格式异常，未建立验证会话")

        store_getter().save_pending_challenge(
            sid,
            secret_getter(),
            challenge_id=challenge_id,
            qq=proposed_qq,
            bot_qq=bot_qq,
            expires_at=expires_at,
        )
        # 原码仅在此响应交给当前浏览器页面，绝不返回 Bearer 密钥或 challenge_id
        return {
            "ok": True,
            "qq": proposed_qq,
            "group_id": EXPECTED_GROUP_ID,
            "bot_qq": bot_qq,
            "verification_code": verification_code,
            "message_to_send": message_to_send,
            "expires_at": expires_at,
            "status": "pending",
        }

    @router.get("/auth/qq/challenges/current")
    def auth_poll_qq_login(request: Request, response: Response):
        no_store(response)
        if request.url.query:
            raise HTTPException(400, "状态查询不允许携带 URL 查询参数")
        raw_sid = request.cookies.get(BROWSER_COOKIE)
        sess = store_getter().get_session(raw_sid, secret_getter())
        if not sess or not sess.get("pending_challenge_id") or not sess.get("pending_qq"):
            raise HTTPException(404, ERROR_MESSAGES["no_login_attempt"])

        challenge_id = str(sess["pending_challenge_id"])
        bound_qq = str(sess["pending_qq"])

        try:
            http_status, data = upstream_getter().request(
                f"/v1/auth/challenges/{challenge_id}",
                method="GET",
            )
        except HTTPException:
            raise
        except Exception:
            raise HTTPException(503, ERROR_MESSAGES["auth_unavailable"])

        if http_status == 404:
            store_getter().clear_pending_challenge(raw_sid, secret_getter())
            raise HTTPException(404, ERROR_MESSAGES["challenge_not_found"])

        if (
            http_status != 200
            or data.get("challenge_id") != challenge_id
            or data.get("qq") != bound_qq
            or data.get("group_id") != EXPECTED_GROUP_ID
        ):
            err_code = str(data.get("error") or "")
            raise HTTPException(503, ERROR_MESSAGES.get(err_code, f"暂时无法查询验证状态（HTTP {http_status}）"))

        status = str(data.get("status") or "pending")
        if status in {"expired", "denied", "consumed"}:
            # 若已经终态失效，清理挂起引用以便前端引导重新发码
            if status in {"expired", "denied"}:
                store_getter().clear_pending_challenge(raw_sid, secret_getter())
        return {
            "qq": bound_qq,
            "bot_qq": str(data.get("bot_qq") or EXPECTED_BOT_QQ),
            "status": status,
            "expires_at": data.get("expires_at") or sess.get("pending_expires_at"),
            "verified_at": data.get("verified_at"),
        }

    @router.post("/auth/qq/challenges/consume")
    async def auth_complete_qq_login(request: Request, response: Response):
        no_store(response)
        if request.url.query:
            raise HTTPException(400, "消费请求不允许携带 URL 查询参数")
        raw_body = (await request.body()).strip()
        if raw_body and raw_body != b"{}":
            raise HTTPException(400, "消费接口不接受自定义正文参数")

        raw_sid = request.cookies.get(BROWSER_COOKIE)
        sess = store_getter().get_session(raw_sid, secret_getter())
        if not sess or not sess.get("pending_challenge_id") or not sess.get("pending_qq"):
            raise HTTPException(400, ERROR_MESSAGES["no_login_attempt"])

        challenge_id = str(sess["pending_challenge_id"])
        bound_qq = str(sess["pending_qq"])

        try:
            http_status, data = upstream_getter().request(
                f"/v1/auth/challenges/{challenge_id}/consume",
                method="POST",
            )
        except HTTPException:
            raise
        except Exception:
            # 超时或响应丢失绝不放行
            raise HTTPException(503, ERROR_MESSAGES["unable_to_verify"])

        if http_status != 200:
            err_code = str(data.get("error") or "")
            if http_status in {403, 404, 410} or err_code in {"already_consumed", "not_in_group", "no_uploaded_card", "challenge_expired"}:
                store_getter().clear_pending_challenge(raw_sid, secret_getter())
            mapped = http_status if http_status in {400, 403, 404, 409, 410, 503} else 503
            raise HTTPException(mapped, ERROR_MESSAGES.get(err_code, ERROR_MESSAGES["unable_to_verify"]))

        # 严格按文档核对全部 8 项不变量
        if (
            data.get("passed") is not True
            or data.get("stale") is not False
            or data.get("auth_type") != "qq_private_code"
            or data.get("status") != "consumed"
            or data.get("group_id") != EXPECTED_GROUP_ID
            or data.get("qq") != bound_qq
            or data.get("challenge_id") != challenge_id
        ):
            store_getter().clear_pending_challenge(raw_sid, secret_getter())
            err_code = str(data.get("error") or "unable_to_verify")
            raise HTTPException(403, ERROR_MESSAGES.get(err_code, ERROR_MESSAGES["unable_to_verify"]))

        user = store_getter().bind_logged_in_user(
            raw_sid,
            secret_getter(),
            qq=bound_qq,
            nickname=str(data.get("group_nickname") or ""),
            avatar_url=str(data.get("avatar_url") or ""),
            card_count=int(data.get("card_count") or 0),
            group_id=EXPECTED_GROUP_ID,
        )
        set_browser_cookie(response, request, raw_sid)
        updated_sess = store_getter().get_session(raw_sid, secret_getter())
        admin_cookie = is_admin_request(request)
        is_mod = admin_cookie or user_is_qq_moderator(user)
        return {
            "ok": True,
            **format_viewer(updated_sess, user, admin_cookie=admin_cookie, is_mod=is_mod),
        }

    @router.post("/auth/logout")
    def auth_logout(request: Request, response: Response):
        no_store(response)
        raw_sid = request.cookies.get(BROWSER_COOKIE)
        store_getter().logout_session(raw_sid, secret_getter())
        response.delete_cookie(BROWSER_COOKIE, path="/")
        return {"ok": True}

    # ---------------- 2. 独立主页留言板与全站评论动态流 ----------------

    @router.get("/community/home")
    def community_home(
        request: Request,
        response: Response,
        limit: int = 25,
        offset: int = 0,
        sort: Literal["latest", "oldest", "hot"] = "latest",
    ):
        no_store(response)
        _sess, user, _admin_cookie, is_mod = viewer_context(request)
        viewer_qq = str(user["qq"]) if user else None
        board = store_getter().list_threads(
            scope="board",
            limit=limit,
            offset=offset,
            sort=sort,
            viewer_qq=viewer_qq,
            viewer_is_mod=is_mod,
        )
        recent_comments = store_getter().list_recent_resource_comments(
            limit=15,
            viewer_qq=viewer_qq,
            viewer_is_mod=is_mod,
        )
        stats = store_getter().stats_overview()
        return {
            "board": board,
            "recent_resource_comments": recent_comments,
            "stats": stats,
        }

    @router.post("/community/board")
    def create_board_post(body: CreateBoardPostRequest, request: Request, response: Response):
        no_store(response)
        user, _is_mod, badge = require_active_member(request)
        try:
            post = store_getter().create_post(
                scope="board",
                author=user,
                author_badge=badge,
                content=body.content,
                reply_to_id=body.reply_to_id,
            )
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, "post": post}

    # ---------------- 3. 序列档案评论、能力卡锚点关联与点赞共鸣 ----------------

    @router.get("/community/comments")
    def list_resource_comments(
        request: Request,
        response: Response,
        path: str = Query(..., min_length=1),
        anchor_id: str = "",
        limit: int = 50,
        offset: int = 0,
        sort: Literal["latest", "oldest", "hot"] = "latest",
    ):
        no_store(response)
        rel_path, title = resolve_public_resource(path)
        _sess, user, _admin_cookie, is_mod = viewer_context(request)
        viewer_qq = str(user["qq"]) if user else None
        data = store_getter().list_threads(
            scope="resource",
            resource_path=rel_path,
            anchor_id=anchor_id.strip(),
            limit=limit,
            offset=offset,
            sort=sort,
            viewer_qq=viewer_qq,
            viewer_is_mod=is_mod,
        )
        return {"resource_path": rel_path, "resource_title": title, **data}

    @router.post("/community/comments")
    def create_resource_comment(body: CreateResourceCommentRequest, request: Request, response: Response):
        no_store(response)
        user, _is_mod, badge = require_active_member(request)
        rel_path, title = resolve_public_resource(body.resource_path)
        anchor_id = body.anchor_id.strip()
        if anchor_id and not re.fullmatch(r"[A-Za-z0-9._-]{1,80}", anchor_id):
            raise HTTPException(400, "能力卡锚点标识不合法")
        try:
            post = store_getter().create_post(
                scope="resource",
                author=user,
                author_badge=badge,
                content=body.content,
                resource_path=rel_path,
                resource_title=title,
                anchor_id=anchor_id,
                anchor_label=body.anchor_label.strip(),
                reply_to_id=body.reply_to_id,
            )
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, "post": post}

    @router.post("/community/posts/{post_id}/like")
    def toggle_post_like(post_id: int, request: Request, response: Response):
        no_store(response)
        user, _is_mod, _badge = require_active_member(request)
        return store_getter().toggle_like(post_id=post_id, user_qq=str(user["qq"]))

    @router.delete("/community/posts/{post_id}")
    def delete_community_post(post_id: int, request: Request, response: Response):
        no_store(response)
        _sess, user, _admin_cookie, is_mod = viewer_context(request)
        if not user and not is_mod:
            raise HTTPException(401, "请先登录")
        actor = f"qq:{user['qq']}" if user else "admin_console"
        return store_getter().delete_post(
            post_id=post_id,
            actor=actor,
            is_mod=is_mod,
            viewer_qq=str(user["qq"]) if user else None,
        )

    @router.post("/community/posts/{post_id}/pin")
    def pin_community_post(post_id: int, body: PinPostRequest, request: Request, response: Response):
        no_store(response)
        require_moderator(request)
        post = store_getter().set_pinned(post_id=post_id, pinned=body.pinned)
        return {"ok": True, "post": post}

    # ---------------- 4. 社区后台管理与管理员 QQ 快捷管控 ----------------

    @router.get("/community/moderation/posts")
    def moderation_list_posts(
        request: Request,
        response: Response,
        scope: Literal["all", "board", "resource"] = "all",
        status: Literal["active", "deleted", "all"] = "active",
        q: str = "",
        qq: str = "",
        limit: int = 50,
        offset: int = 0,
    ):
        no_store(response)
        require_moderator(request)
        return {
            "stats": store_getter().stats_overview(),
            **store_getter().admin_list_posts(scope=scope, status=status, q=q, qq=qq, limit=limit, offset=offset),
        }

    @router.get("/community/moderation/users")
    def moderation_list_users(
        request: Request,
        response: Response,
        filter_mode: Literal["all", "banned", "moderator"] = "all",
        q: str = "",
        limit: int = 80,
        offset: int = 0,
    ):
        no_store(response)
        require_moderator(request)
        return store_getter().admin_list_users(filter_mode=filter_mode, q=q, limit=limit, offset=offset)

    @router.post("/community/moderation/ban")
    def moderation_ban_user(body: AdminBanUserRequest, request: Request, response: Response):
        no_store(response)
        actor, _user = require_moderator(request)
        target_qq = (body.qq or "").strip()
        if not QQ_RE.fullmatch(target_qq):
            raise HTTPException(400, "请输入有效的目标 QQ 号（5～12 位数字）")
        if target_qq in parse_admin_qqs() and not is_admin_request(request):
            raise HTTPException(403, "不能封禁站点维护者白名单 QQ")

        if body.unban:
            banned_until = 0
            reason = ""
        elif body.duration_hours is None or body.duration_hours <= 0:
            banned_until = -1
            reason = body.reason.strip() or f"由管理员（{actor}）限制发言"
        else:
            banned_until = now_ts() + int(body.duration_hours) * 3600
            reason = body.reason.strip() or f"禁言 {body.duration_hours} 小时"

        row = store_getter().set_user_ban(
            target_qq,
            banned_until=banned_until,
            reason=reason,
            nickname_hint=body.nickname_hint,
        )
        return {"ok": True, "qq": target_qq, "banned_until": row["banned_until"], "ban_reason": row["ban_reason"]}

    @router.post("/community/moderation/users/{qq}/role")
    def moderation_set_role(qq: str, body: AdminRoleRequest, request: Request, response: Response):
        no_store(response)
        # 任命/撤销协管仅允许 ADMIN_PASSWORD 控制台或环境变量白名单维护者执行
        _sess, user, admin_cookie, _is_mod = viewer_context(request)
        is_super = admin_cookie or bool(user and str(user["qq"]) in parse_admin_qqs())
        if not is_super:
            raise HTTPException(403, "仅站点维护者可调整协管权限")
        if not QQ_RE.fullmatch(qq.strip()):
            raise HTTPException(400, "QQ 号格式不合法")
        row = store_getter().set_user_role(qq.strip(), body.role)
        return {"ok": True, "qq": row["qq"], "role": row["role"]}

    @router.post("/community/moderation/announcement")
    def moderation_post_announcement(body: AdminAnnouncementRequest, request: Request, response: Response):
        no_store(response)
        _actor, user = require_moderator(request)
        # 允许 ADMIN_PASSWORD 后台在不登录个人 QQ 的情况下直接发布置顶站点公告
        if user:
            author_info = user
        else:
            author_info = {
                "qq": "10000",
                "nickname": body.nickname.strip() or "界域终端公告",
                "avatar_url": "",
            }
        try:
            post = store_getter().create_post(
                scope="board",
                author=author_info,
                author_badge="系统公告" if not user else "维护者",
                content=body.content,
                bypass_cooldown=True,
            )
            if body.pinned:
                post = store_getter().set_pinned(post_id=int(post["id"]), pinned=True)
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, "post": post}

    return router
