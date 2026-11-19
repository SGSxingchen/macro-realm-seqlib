"""主群 QQ 验证、持久会话和轮回者论坛。

网页挑战绑定 HttpOnly Cookie 对应的浏览器会话。主群机器人核验身份后，
页面使用一次性确认凭据提交，身份记录与浏览器会话在同一事务中更新。
历史 QQ 私聊接口保留原有独立契约，资源评论继续使用去序号标识及历史路径映射。
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
from pathlib import Path, PurePosixPath
from typing import Any, Callable, Iterator, Literal, Optional

import httpx
from fastapi import APIRouter, HTTPException, Query, Request, Response
from pydantic import BaseModel, ConfigDict, Field

QQ_RE = re.compile(r"^[1-9][0-9]{4,11}$")
OPENID_HEX_RE = re.compile(r"^[0-9A-F]{16,64}$")
USER_KEY_RE = re.compile(r"^(?:[1-9][0-9]{4,11}|oid:[0-9A-F]{16,64})$")
BOT_SHORT_CODE_RE = re.compile(r"^[0-9]{6}$")

EXPECTED_GROUP_ID = 672204828
OFFICIAL_BOT_APP_ID = os.getenv("COMMUNITY_BOT_APP_ID", "1903175013").strip() or "1903175013"
OFFICIAL_BOT_NAME = os.getenv("COMMUNITY_BOT_NAME", "阿瓦娜").strip() or "阿瓦娜"
DEFAULT_ALLOWED_OPEN_GROUPS = "OpenQQ-Group:4019653007-6C93AE22F9628FC5D947F07C55643B91"
BOT_SHORT_CODE_TTL_SECONDS = 300

# 私聊验证码接收 BOT 的 QQ 号（历史兼容）；换号时优先改环境变量 COMMUNITY_BOT_QQ
EXPECTED_BOT_QQ = os.getenv("COMMUNITY_BOT_QQ", "412318056").strip() or "412318056"
DEFAULT_AUTH_BASE = "https://trpg.chordvers.org/onebot-sequences"

BROWSER_COOKIE = "seqlib_community_sid"
BROWSER_TTL_SECONDS = 30 * 86400  # 30 天持久登录会话
COMMENT_COOLDOWN_SECONDS = 3

ERROR_MESSAGES: dict[str, str] = {
    "invalid_request": "QQ 号或请求参数格式不合法。",
    "unauthorized": "网站后端访问密钥错误或未配置，请联系站点维护者检查 ONEBOT_SEQUENCE_TOKEN。",
    "not_in_group": "该 QQ 当前不在交易群（672204828），无法获取验证码或登录。",
    "no_uploaded_card": "该 QQ 尚未在交易群指定目录上传过 .xlsx 角色卡，暂未通过轮回者发言核验。",
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
    if not qq:
        return ""
    if qq.startswith("oid:"):
        raw = qq[4:]
        if len(raw) <= 6:
            return f"ID:{raw}"
        return f"ID:{raw[:4]}…{raw[-4:]}"
    if len(qq) < 5:
        return qq
    if len(qq) <= 7:
        return f"{qq[:2]}***{qq[-2:]}"
    return f"{qq[:3]}****{qq[-3:]}"


def normalize_open_user_id(raw: str) -> str:
    """将海豹 OpenQQ:4019653007-HEX、oid:HEX 或裸 HEX 归一化为 16~64 位大写 HEX openid。"""
    val = (raw or "").strip()
    if val.lower().startswith("oid:"):
        val = val[4:].strip()
    elif "-" in val and ":" in val:
        val = val.rsplit("-", 1)[-1].strip()
    val = val.upper()
    if not OPENID_HEX_RE.fullmatch(val):
        raise ValueError("无效的官方 Bot 用户 OpenID")
    return val


def normalize_user_key(raw: str) -> str:
    """兼容历史数字 QQ、oid:<HEX>、OpenQQ:<uin>-<HEX> 或裸 <HEX> 用户标识。"""
    val = (raw or "").strip()
    if QQ_RE.fullmatch(val):
        return val
    try:
        return f"oid:{normalize_open_user_id(val)}"
    except ValueError:
        raise ValueError("请输入有效的用户标识（QQ 号或官方 Bot ID）")


def normalize_open_group_token(raw: str) -> str:
    val = (raw or "").strip()
    if "-" in val and val.lower().startswith("openqq-group:"):
        val = val.rsplit("-", 1)[-1].strip()
    return val.upper()


def parse_allowed_open_groups() -> set[str]:
    raw = os.getenv("COMMUNITY_ALLOWED_OPEN_GROUPS", DEFAULT_ALLOWED_OPEN_GROUPS)
    result: set[str] = set()
    for part in re.split(r"[,;\s]+", raw or DEFAULT_ALLOWED_OPEN_GROUPS):
        tok = normalize_open_group_token(part)
        if tok:
            result.add(tok)
    return result


def bot_bridge_token() -> str:
    return (
        os.getenv("COMMUNITY_BOT_TOKEN")
        or os.getenv("ONEBOT_SEQUENCE_TOKEN")
        or ""
    ).strip()


def bot_auth_configured() -> bool:
    return bool(bot_bridge_token())


def default_avatar_for_user(user_key: str, app_id: str | None = None) -> str:
    if user_key.startswith("oid:"):
        oid = user_key[4:]
        aid = (app_id or OFFICIAL_BOT_APP_ID).strip() or OFFICIAL_BOT_APP_ID
        return f"https://q.qlogo.cn/qqapp/{aid}/{oid}/640"
    if QQ_RE.fullmatch(user_key):
        return f"https://q1.qlogo.cn/g?b=qq&nk={user_key}&s=640"
    return ""


def clean_nickname(raw: str, *, max_len: int = 32) -> str:
    cleaned = re.sub(r"[\x00-\x1f\x7f]+", " ", raw or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if not cleaned:
        raise ValueError("昵称不能为空")
    if len(cleaned) > max_len:
        raise ValueError(f"昵称长度不能超过 {max_len} 个字符")
    return cleaned


def parse_admin_qqs() -> set[str]:
    raw = os.getenv("COMMUNITY_ADMIN_QQS", "")
    extra_oids = os.getenv("COMMUNITY_ADMIN_OPENIDS", "FF834522288AB2D2993970BE14970D3A")
    result: set[str] = set()
    for part in re.split(r"[,;\s]+", f"{raw},{extra_oids}"):
        cleaned = part.strip()
        if not cleaned:
            continue
        if QQ_RE.fullmatch(cleaned):
            result.add(cleaned)
            continue
        try:
            result.add(f"oid:{normalize_open_user_id(cleaned)}")
        except ValueError:
            continue
    return result


def compute_resource_key(rel_path: str) -> str:
    """提取不受编号重排 (`renumber_resources.py`) 和侧别子目录挪动影响的稳定资源标识。

    例如：
    - `序列库/职业/神秘侧/001】天师.txt` -> `职业::天师`
    - `序列库/特质改造/异化改造类/005】从者化《Fate》.txt` -> `特质改造::从者化《Fate》`
    - `序列库/新过审序列/012】剑圣（职业·战技侧）.txt` -> `剑圣`（剥离新过审括号标注以便入正式库时可归一化关联）
    """
    norm = (rel_path or "").replace("\\", "/").strip("/")
    if not norm:
        return ""
    p = PurePosixPath(norm)
    parts = p.parts
    top_kind = ""
    if len(parts) >= 3 and parts[0] in ("序列库", "荣誉室"):
        top_kind = parts[1]
    stem = p.stem
    bare = re.sub(r"^\d+\s*[】\]\.]\s*", "", stem).strip()
    bare = re.sub(r"\s+", " ", bare)
    # 兼容新过审目录转正：若位于新过审序列且末尾带（职业·战技侧）类分类标注，提取对应大类
    if top_kind == "新过审序列":
        m = re.match(
            r"^(.+?)[（(]\s*(职业|技能表|特质改造|特质|能量池|公共建筑)(?:[·/\-\s][^）)]*)?[）)]$",
            bare,
        )
        if m:
            mapped_top = "特质改造" if m.group(2) == "特质" else m.group(2)
            return f"{mapped_top}::{m.group(1).strip()}"
    if top_kind == "特质":
        top_kind = "特质改造"
    return f"{top_kind}::{bare}" if top_kind else bare


def normalize_anchor_label(label: str) -> str:
    """剥离能力名中的品阶后缀（如 `（A级）`、`[EX]`），防止平衡性调整升降阶后锚点失联。"""
    cleaned = re.sub(r"[（(\[【]\s*(?:EX|[SABCDEF][+-]?)\s*级?\s*[）)\]】]", "", label or "", flags=re.IGNORECASE)
    return re.sub(r"\s+", " ", cleaned).strip()[:120]


def clean_post_title(raw: str, *, max_len: int = 100) -> str:
    cleaned = re.sub(r"[\x00-\x1f\x7f]+", " ", raw or "")
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    if len(cleaned) > max_len:
        raise ValueError(f"标题超出长度上限（最多 {max_len} 字）")
    return cleaned


def clean_comment_content(raw: str, *, max_len: int = 4000) -> str:
    cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", raw or "")
    cleaned = cleaned.replace("\r\n", "\n").replace("\r", "\n").strip()
    cleaned = re.sub(r"\n{4,}", "\n\n\n", cleaned)
    if not cleaned:
        raise ValueError("内容不能为空")
    if len(cleaned) > max_len:
        raise ValueError(f"内容超出长度上限（最多 {max_len} 字）")
    return cleaned


def fallback_thread_title(title: str, content: str) -> str:
    t = (title or "").strip()
    if t:
        return t
    first_line = next((line.strip() for line in (content or "").splitlines() if line.strip()), "")
    if len(first_line) <= 44:
        return first_line or "无题通讯"
    return first_line[:42] + "…"


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

    def request(
        self,
        path: str,
        *,
        method: str = "GET",
        body: dict[str, Any] | None = None,
        params: dict[str, Any] | None = None,
    ) -> tuple[int, dict[str, Any]]:
        token = (self._token_getter() or "").strip()
        if not token:
            raise HTTPException(503, ERROR_MESSAGES["unauthorized"])
        if params is None and "?" in path:
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
                resp = client.request(method=method, url=url, headers=headers, content=content_bytes, params=params)
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
    """SQLite 持久化存储：浏览器会话、轮回者档案、论坛主题帖、抗改名序列评论与路径别名。"""

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
                    title TEXT NOT NULL DEFAULT '',
                    resource_key TEXT NOT NULL DEFAULT '',
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
                    last_reply_at INTEGER NOT NULL DEFAULT 0,
                    last_reply_nickname TEXT NOT NULL DEFAULT '',
                    status TEXT NOT NULL DEFAULT 'active' CHECK(status IN ('active', 'deleted')),
                    deleted_by TEXT NOT NULL DEFAULT '',
                    created_at INTEGER NOT NULL,
                    updated_at INTEGER NOT NULL
                );

                CREATE TABLE IF NOT EXISTS community_post_likes (
                    post_id INTEGER NOT NULL,
                    user_qq TEXT NOT NULL,
                    created_at INTEGER NOT NULL,
                    PRIMARY KEY (post_id, user_qq),
                    FOREIGN KEY (post_id) REFERENCES community_posts(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS community_resource_aliases (
                    path TEXT PRIMARY KEY,
                    canonical_key TEXT NOT NULL,
                    latest_path TEXT NOT NULL DEFAULT '',
                    latest_title TEXT NOT NULL DEFAULT '',
                    updated_at INTEGER NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_res_alias_key ON community_resource_aliases(canonical_key);

                CREATE TABLE IF NOT EXISTS community_web_challenges (
                    id INTEGER PRIMARY KEY,
                    sid_hash TEXT NOT NULL REFERENCES browser_sessions(sid_hash) ON DELETE CASCADE,
                    code TEXT NOT NULL UNIQUE,
                    status TEXT NOT NULL DEFAULT 'pending' CHECK(status IN ('pending', 'verified', 'consumed', 'expired')),
                    user_key TEXT NOT NULL DEFAULT '',
                    sealdice_nickname TEXT NOT NULL DEFAULT '',
                    expires_at INTEGER NOT NULL
                );
                CREATE INDEX IF NOT EXISTS idx_web_challenges_sid ON community_web_challenges(sid_hash, id DESC);
                CREATE INDEX IF NOT EXISTS idx_web_challenges_exp ON community_web_challenges(expires_at);
                """
            )
            # 在线平滑升级旧表结构（兼容已部署库）
            cols = {str(r["name"]) for r in conn.execute("PRAGMA table_info(community_posts)").fetchall()}
            migrations = [
                ("title", "ALTER TABLE community_posts ADD COLUMN title TEXT NOT NULL DEFAULT ''"),
                ("resource_key", "ALTER TABLE community_posts ADD COLUMN resource_key TEXT NOT NULL DEFAULT ''"),
                ("last_reply_at", "ALTER TABLE community_posts ADD COLUMN last_reply_at INTEGER NOT NULL DEFAULT 0"),
                ("last_reply_nickname", "ALTER TABLE community_posts ADD COLUMN last_reply_nickname TEXT NOT NULL DEFAULT ''"),
            ]
            for col_name, ddl in migrations:
                if col_name not in cols:
                    conn.execute(ddl)

            sess_cols = {str(r["name"]) for r in conn.execute("PRAGMA table_info(browser_sessions)").fetchall()}
            if "pending_code" not in sess_cols:
                conn.execute("ALTER TABLE browser_sessions ADD COLUMN pending_code TEXT")

            conn.executescript(
                """
                CREATE INDEX IF NOT EXISTS idx_posts_board ON community_posts(scope, status, is_pinned DESC, last_reply_at DESC, id DESC);
                CREATE INDEX IF NOT EXISTS idx_posts_res_key ON community_posts(scope, resource_key, status, is_pinned DESC, created_at ASC, id ASC);
                CREATE INDEX IF NOT EXISTS idx_posts_resource ON community_posts(scope, resource_path, status, is_pinned DESC, created_at ASC, id ASC);
                CREATE INDEX IF NOT EXISTS idx_posts_parent ON community_posts(parent_id, status, created_at ASC);
                CREATE INDEX IF NOT EXISTS idx_posts_author ON community_posts(author_qq, created_at DESC);
                """
            )

            # 回填旧数据的 last_reply_at、last_reply_nickname 与 resource_key（保护已产生的线上帖子与回复）
            conn.execute(
                """
                UPDATE community_posts
                SET last_reply_at = COALESCE(
                        (SELECT MAX(r.created_at) FROM community_posts r WHERE r.parent_id = community_posts.id AND r.status = 'active'),
                        created_at
                    ),
                    last_reply_nickname = COALESCE(
                        (SELECT r.author_nickname FROM community_posts r WHERE r.parent_id = community_posts.id AND r.status = 'active' ORDER BY r.created_at DESC, r.id DESC LIMIT 1),
                        author_nickname
                    )
                WHERE parent_id IS NULL AND last_reply_at = 0
                """
            )
            legacy_res = conn.execute(
                "SELECT id, resource_path, resource_title, anchor_label FROM community_posts WHERE scope = 'resource' AND resource_key = '' AND resource_path != ''"
            ).fetchall()
            now = now_ts()
            for r in legacy_res:
                rk = compute_resource_key(str(r["resource_path"]))
                norm_label = normalize_anchor_label(str(r["anchor_label"] or ""))
                conn.execute(
                    "UPDATE community_posts SET resource_key = ?, anchor_label = ? WHERE id = ?",
                    (rk, norm_label, int(r["id"])),
                )
                if rk:
                    conn.execute(
                        """
                        INSERT INTO community_resource_aliases (path, canonical_key, latest_path, latest_title, updated_at)
                        VALUES (?, ?, ?, ?, ?)
                        ON CONFLICT(path) DO UPDATE SET canonical_key = excluded.canonical_key, updated_at = excluded.updated_at
                        """,
                        (str(r["resource_path"]), rk, str(r["resource_path"]), str(r["resource_title"] or ""), now),
                    )

    # ---------------- 资源抗重排 / 抗改名别名与自愈逻辑 ----------------

    def resolve_canonical_key_with_conn(self, conn: sqlite3.Connection, rel_path: str, current_title: str = "") -> str:
        norm_path = (rel_path or "").replace("\\", "/").strip("/")
        if not norm_path:
            return ""
        now = now_ts()
        alias_row = conn.execute(
            "SELECT canonical_key FROM community_resource_aliases WHERE path = ?",
            (norm_path,),
        ).fetchone()
        computed_key = compute_resource_key(norm_path)
        canonical_key = str(alias_row["canonical_key"]) if alias_row and alias_row["canonical_key"] else computed_key
        if not canonical_key:
            return ""

        # 记录当前有效路径，并自愈同 canonical_key 下旧帖子的展示路径与标题（例如 012】天师 -> 015】天师）
        conn.execute(
            """
            INSERT INTO community_resource_aliases (path, canonical_key, latest_path, latest_title, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(path) DO UPDATE SET
                canonical_key = excluded.canonical_key,
                latest_path = excluded.latest_path,
                latest_title = CASE WHEN excluded.latest_title != '' THEN excluded.latest_title ELSE community_resource_aliases.latest_title END,
                updated_at = excluded.updated_at
            """,
            (norm_path, canonical_key, norm_path, current_title, now),
        )
        if computed_key and computed_key != canonical_key:
            # 把去序号键也桥接到最初的 canonical_key，保证重排后无论哪条路径都能命中同一条评论链
            conn.execute(
                """
                UPDATE community_posts
                SET resource_key = ?
                WHERE scope = 'resource' AND resource_key = ?
                """,
                (canonical_key, computed_key),
            )
        if current_title:
            conn.execute(
                """
                UPDATE community_resource_aliases
                SET latest_path = ?, latest_title = ?, updated_at = ?
                WHERE canonical_key = ?
                """,
                (norm_path, current_title, now, canonical_key),
            )
            conn.execute(
                """
                UPDATE community_posts
                SET resource_path = ?, resource_title = ?
                WHERE scope = 'resource' AND resource_key = ? AND (resource_path != ? OR resource_title != ?)
                """,
                (norm_path, current_title, canonical_key, norm_path, current_title),
            )
        return canonical_key

    def record_resource_move(self, old_path: str, new_path: str, new_title: str = "") -> None:
        """当后台发生文件改名、跨目录移动或移入荣誉室时，自动继承原资源的 canonical_key 并刷新评论路径。"""
        old_norm = (old_path or "").replace("\\", "/").strip("/")
        new_norm = (new_path or "").replace("\\", "/").strip("/")
        if not old_norm or not new_norm:
            return
        now = now_ts()
        with self._conn() as conn:
            old_alias = conn.execute(
                "SELECT canonical_key, latest_title FROM community_resource_aliases WHERE path = ?",
                (old_norm,),
            ).fetchone()
            canonical_key = (
                str(old_alias["canonical_key"])
                if old_alias and old_alias["canonical_key"]
                else compute_resource_key(old_norm)
            )
            if not canonical_key:
                canonical_key = compute_resource_key(new_norm)
            title_val = new_title.strip() or (str(old_alias["latest_title"]) if old_alias else "")

            for p in (old_norm, new_norm):
                conn.execute(
                    """
                    INSERT INTO community_resource_aliases (path, canonical_key, latest_path, latest_title, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(path) DO UPDATE SET
                        canonical_key = excluded.canonical_key,
                        latest_path = excluded.latest_path,
                        latest_title = CASE WHEN excluded.latest_title != '' THEN excluded.latest_title ELSE community_resource_aliases.latest_title END,
                        updated_at = excluded.updated_at
                    """,
                    (p, canonical_key, new_norm, title_val, now),
                )
            conn.execute(
                "UPDATE community_resource_aliases SET latest_path = ?, latest_title = CASE WHEN ? != '' THEN ? ELSE latest_title END, updated_at = ? WHERE canonical_key = ?",
                (new_norm, title_val, title_val, now, canonical_key),
            )
            new_computed = compute_resource_key(new_norm)
            keys_to_unify = {k for k in (canonical_key, compute_resource_key(old_norm), new_computed) if k}
            if keys_to_unify:
                placeholders = ",".join("?" for _ in keys_to_unify)
                conn.execute(
                    f"""
                    UPDATE community_posts
                    SET resource_key = ?,
                        resource_path = ?,
                        resource_title = CASE WHEN ? != '' THEN ? ELSE resource_title END
                    WHERE scope = 'resource' AND (resource_key IN ({placeholders}) OR resource_path = ?)
                    """,
                    [canonical_key, new_norm, title_val, title_val, *keys_to_unify, old_norm],
                )

    # ---------------- 会话与用户管理 ----------------

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
        verification_code: str,
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
                    pending_code = ?,
                    pending_expires_at = ?,
                    pending_created_at = ?,
                    updated_at = ?,
                    expires_at = ?
                WHERE sid_hash = ?
                """,
                (challenge_id, qq, bot_qq, verification_code, expires_at, now, now, now + BROWSER_TTL_SECONDS, sid_hash),
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
                    pending_code = NULL,
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
        preserve_existing_nickname: bool = False,
    ) -> dict[str, Any]:
        now = now_ts()
        sid_hash = self._hash_sid(raw_sid, secret)
        safe_nick = (nickname or "").strip() or f"轮回者{qq[-4:]}"
        safe_avatar = (avatar_url or "").strip() or default_avatar_for_user(qq)
        with self._conn() as conn:
            existing = conn.execute("SELECT * FROM community_users WHERE qq = ?", (qq,)).fetchone()
            if existing:
                final_nick = (
                    str(existing["nickname"]).strip() or safe_nick
                    if preserve_existing_nickname
                    else safe_nick
                )
                final_cards = max(int(existing["card_count"] or 0), int(card_count or 0))
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
                    (final_nick, safe_avatar, final_cards, int(group_id), now, now, qq),
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
                    pending_code = NULL,
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

    @staticmethod
    def _latest_web_challenge(conn: sqlite3.Connection, sid_hash: str) -> sqlite3.Row | None:
        return conn.execute(
            """SELECT id, code, status, user_key, sealdice_nickname, expires_at
               FROM community_web_challenges WHERE sid_hash = ? ORDER BY id DESC LIMIT 1""",
            (sid_hash,),
        ).fetchone()

    @staticmethod
    def _web_challenge_view(
        conn: sqlite3.Connection, row: sqlite3.Row | None, secret: bytes
    ) -> dict[str, Any]:
        if row is None or row["status"] in ("consumed", "expired") or row["expires_at"] <= now_ts():
            return {"status": "expired"}
        view = {"status": row["status"], "expires_at": ts_to_iso(row["expires_at"])}
        if row["status"] == "pending":
            return {**view, "code": row["code"], "command": f"/登录 {row['code']}", "bot_name": OFFICIAL_BOT_NAME}

        user_key = row["user_key"]
        existing = conn.execute("SELECT nickname FROM community_users WHERE qq = ?", (user_key,)).fetchone()
        nickname = row["sealdice_nickname"] or f"轮回者{user_key[-4:]}"
        return {
            **view,
            "is_new_user": existing is None,
            "user_key_masked": mask_qq(user_key),
            "suggested_nickname": existing["nickname"] if existing else nickname,
            "sealdice_nickname": row["sealdice_nickname"],
            "avatar_url": default_avatar_for_user(user_key),
            "confirmation": CommunityStore._web_confirmation(row, secret),
        }

    @staticmethod
    def _web_confirmation(row: sqlite3.Row, secret: bytes) -> str:
        data = f"web-login:{row['id']}:{row['code']}:{row['user_key']}:{row['expires_at']}"
        return hmac.new(secret, data.encode(), hashlib.sha256).hexdigest()

    def begin_web_challenge(self, sid_hash: str, secret: bytes, *, renew: bool) -> dict[str, Any]:
        with self._conn() as conn:
            conn.execute("BEGIN IMMEDIATE")
            now = now_ts()
            row = self._latest_web_challenge(conn, sid_hash)
            if row and row["expires_at"] > now and (
                row["status"] == "pending" or (row["status"] == "verified" and not renew)
            ):
                return self._web_challenge_view(conn, row, secret)
            if row and row["status"] in ("pending", "verified"):
                conn.execute("UPDATE community_web_challenges SET status = 'expired' WHERE id = ?", (row["id"],))
            conn.execute(
                """DELETE FROM community_web_challenges WHERE id IN (
                       SELECT id FROM community_web_challenges WHERE expires_at <= ? ORDER BY expires_at LIMIT 100
                   )""",
                (now - BOT_SHORT_CODE_TTL_SECONDS,),
            )
            for _ in range(20):
                code = f"{secrets.randbelow(1_000_000):06d}"
                inserted = conn.execute(
                    """INSERT INTO community_web_challenges (sid_hash, code, expires_at) VALUES (?, ?, ?)
                       ON CONFLICT(code) DO NOTHING""",
                    (sid_hash, code, now + BOT_SHORT_CODE_TTL_SECONDS),
                )
                if inserted.rowcount:
                    return self._web_challenge_view(conn, self._latest_web_challenge(conn, sid_hash), secret)
            raise HTTPException(503, "暂时无法生成登录码，请稍后重试")

    def current_web_challenge(self, sid_hash: str, secret: bytes) -> dict[str, Any]:
        with self._conn() as conn:
            return self._web_challenge_view(conn, self._latest_web_challenge(conn, sid_hash), secret)

    def verify_web_challenge(self, code: str, user_key: str, nickname: str) -> None:
        if not BOT_SHORT_CODE_RE.fullmatch(code):
            raise HTTPException(400, "请输入 6 位数字验证码")
        with self._conn() as conn:
            conn.execute("BEGIN IMMEDIATE")
            now = now_ts()
            row = conn.execute(
                "SELECT id, status, expires_at FROM community_web_challenges WHERE code = ?", (code,)
            ).fetchone()
            if row is None:
                raise HTTPException(404, "没有找到该验证码，请检查网页上的指令")
            if row["status"] == "expired" or row["expires_at"] <= now:
                raise HTTPException(410, "验证码已过期，请回网页重新取码")
            if row["status"] != "pending":
                raise HTTPException(409, "该验证码已核验，请在网页核对身份")
            conn.execute(
                """UPDATE community_web_challenges SET status = 'verified', user_key = ?, sealdice_nickname = ?
                   WHERE id = ? AND status = 'pending'""",
                (user_key, re.sub(r"[\x00-\x1f\x7f]+", " ", nickname).strip()[:32], row["id"]),
            )

    def consume_web_challenge(
        self, sid_hash: str, secret: bytes, *, confirmation: str, nickname: str
    ) -> dict[str, Any]:
        with self._conn() as conn:
            conn.execute("BEGIN IMMEDIATE")
            now = now_ts()
            row = self._latest_web_challenge(conn, sid_hash)
            if row is None or row["status"] != "verified" or row["expires_at"] <= now:
                raise HTTPException(409, "当前浏览器没有可确认的登录验证")
            if not re.fullmatch(r"[0-9a-f]{64}", confirmation) or not hmac.compare_digest(
                self._web_confirmation(row, secret), confirmation
            ):
                raise HTTPException(409, "待确认身份已变化，请重新核对")
            conn.execute("UPDATE community_web_challenges SET status = 'consumed' WHERE id = ?", (row["id"],))

            user_key = str(row["user_key"])
            existing = conn.execute("SELECT * FROM community_users WHERE qq = ?", (user_key,)).fetchone()
            avatar = default_avatar_for_user(user_key)

            provided_nick = (nickname or "").strip()
            if existing:
                final_nick = existing["nickname"]
            else:
                try:
                    final_nick = clean_nickname(provided_nick)
                except ValueError as exc:
                    raise HTTPException(400, str(exc)) from exc

            if existing:
                conn.execute(
                    """UPDATE community_users SET avatar_url = ?, last_login_at = ?, updated_at = ? WHERE qq = ?""",
                    (avatar, now, now, user_key),
                )
                if avatar != existing["avatar_url"]:
                    conn.execute(
                        "UPDATE community_posts SET author_avatar = ? WHERE author_qq = ?",
                        (avatar, user_key),
                    )
            else:
                conn.execute(
                    """
                    INSERT INTO community_users (
                        qq, nickname, avatar_url, card_count, group_id, role,
                        banned_until, ban_reason, first_login_at, last_login_at, updated_at
                    ) VALUES (?, ?, ?, 0, ?, 'member', 0, '', ?, ?, ?)
                    """,
                    (user_key, final_nick, avatar, EXPECTED_GROUP_ID, now, now, now),
                )

            updated_session = conn.execute(
                """
                UPDATE browser_sessions
                SET user_qq = ?,
                    pending_challenge_id = NULL,
                    pending_qq = NULL,
                    pending_bot_qq = NULL,
                    pending_code = NULL,
                    pending_expires_at = NULL,
                    pending_created_at = NULL,
                    updated_at = ?,
                    expires_at = ?
                WHERE sid_hash = ? AND user_qq IS NULL AND expires_at > ?
                """,
                (user_key, now, now + BROWSER_TTL_SECONDS, sid_hash, now),
            )
            if updated_session.rowcount != 1:
                raise HTTPException(409, "当前浏览器会话已变化，请刷新页面")
            user_row = conn.execute("SELECT * FROM community_users WHERE qq = ?", (user_key,)).fetchone()
            return dict(user_row)

    @staticmethod
    def _sync_author_profile_in_posts(
        conn: sqlite3.Connection,
        *,
        user_key: str,
        nickname: str,
        avatar_url: str = "",
    ) -> None:
        if avatar_url:
            conn.execute(
                "UPDATE community_posts SET author_nickname = ?, author_avatar = ? WHERE author_qq = ?",
                (nickname, avatar_url, user_key),
            )
        else:
            conn.execute(
                "UPDATE community_posts SET author_nickname = ? WHERE author_qq = ?",
                (nickname, user_key),
            )
        conn.execute(
            "UPDATE community_posts SET reply_to_nickname = ? WHERE reply_to_qq = ?",
            (nickname, user_key),
        )

    def update_user_nickname(self, user_key: str, nickname: str) -> dict[str, Any]:
        try:
            final_nick = clean_nickname(nickname)
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        now = now_ts()
        with self._conn() as conn:
            existing = conn.execute("SELECT * FROM community_users WHERE qq = ?", (user_key,)).fetchone()
            if not existing:
                raise HTTPException(404, "用户不存在")
            conn.execute(
                "UPDATE community_users SET nickname = ?, updated_at = ? WHERE qq = ?",
                (final_nick, now, user_key),
            )
            self._sync_author_profile_in_posts(
                conn,
                user_key=user_key,
                nickname=final_nick,
                avatar_url=str(existing["avatar_url"] or ""),
            )
            row = conn.execute("SELECT * FROM community_users WHERE qq = ?", (user_key,)).fetchone()
            return dict(row)

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
                nick = nickname_hint.strip() or (f"轮回者 {mask_qq(qq)}" if qq.startswith("oid:") else f"QQ {qq}")
                avatar = default_avatar_for_user(qq)
                conn.execute(
                    """
                    INSERT INTO community_users (
                        qq, nickname, avatar_url, card_count, group_id, role,
                        banned_until, ban_reason, first_login_at, last_login_at, updated_at
                    ) VALUES (?, ?, ?, 0, ?, 'member', ?, ?, ?, 0, ?)
                    """,
                    (qq, nick, avatar, EXPECTED_GROUP_ID, banned_until, reason.strip(), now, now),
                )
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

    # ---------------- 论坛主题帖与序列评论核心方法 ----------------

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
        raw_title = str(d.get("title") or "")
        content_active = str(d["content"]) if d["status"] == "active" else ""
        display_title = (
            fallback_thread_title(raw_title, content_active)
            if d["scope"] == "board" and d["parent_id"] is None and d["status"] == "active"
            else raw_title
        )
        last_reply_ts = int(d.get("last_reply_at") or d["created_at"])
        item: dict[str, Any] = {
            "id": pid,
            "scope": d["scope"],
            "title": display_title,
            "resource_key": d.get("resource_key") or "",
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
                "avatar_url": d["author_avatar"] or default_avatar_for_user(author_qq),
                "badge": d["author_role_badge"] or "",
            },
            "content": content_active,
            "excerpt": re.sub(r"\s+", " ", content_active).strip()[:120],
            "is_pinned": bool(d["is_pinned"]),
            "like_count": int(d["like_count"] or 0),
            "liked_by_me": bool(liked_ids and pid in liked_ids),
            "can_delete": bool(d["status"] == "active" and (is_mine or viewer_is_mod)),
            "can_pin": bool(d["status"] == "active" and d["parent_id"] is None and viewer_is_mod),
            "status": d["status"],
            "created_at": ts_to_iso(d["created_at"]),
            "created_ts": int(d["created_at"]),
            "last_reply_at": ts_to_iso(last_reply_ts),
            "last_reply_ts": last_reply_ts,
            "last_reply_nickname": str(d.get("last_reply_nickname") or d["author_nickname"]),
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
        title: str = "",
        resource_path: str = "",
        resource_title: str = "",
        anchor_id: str = "",
        anchor_label: str = "",
        reply_to_id: int | None = None,
        bypass_cooldown: bool = False,
    ) -> dict[str, Any]:
        cleaned = clean_comment_content(content, max_len=4000 if scope == "board" else 1600)
        cleaned_title = clean_post_title(title, max_len=100) if title else ""
        now = now_ts()
        author_qq = str(author["qq"])
        author_nick = str(author["nickname"])

        with self._conn() as conn:
            cooldown = comment_cooldown_seconds()
            if not bypass_cooldown and cooldown > 0:
                recent = conn.execute(
                    "SELECT created_at FROM community_posts WHERE author_qq = ? ORDER BY id DESC LIMIT 1",
                    (author_qq,),
                ).fetchone()
                if recent and now - int(recent["created_at"]) < cooldown:
                    raise HTTPException(429, "发言太快啦，请稍等几秒再发送。")

            res_key = ""
            if scope == "resource":
                res_key = self.resolve_canonical_key_with_conn(conn, resource_path, resource_title)

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
                    raise HTTPException(404, "要回复的主题或评论不存在或已被删除")
                if target["scope"] != scope:
                    raise HTTPException(400, "不能跨版块回复")
                if scope == "resource":
                    target_key = str(target["resource_key"] or compute_resource_key(str(target["resource_path"])))
                    if target_key != res_key and str(target["resource_path"]) != resource_path:
                        raise HTTPException(400, "不能跨序列档案回复评论")

                parent_id = int(target["parent_id"]) if target["parent_id"] is not None else int(target["id"])
                actual_reply_id = int(target["id"])
                reply_to_qq = str(target["author_qq"])
                reply_to_nickname = str(target["author_nickname"])
                cleaned_title = ""  # 楼层跟帖不单独带主题标题
                if not anchor_id and target["anchor_id"]:
                    anchor_id = str(target["anchor_id"])
                    anchor_label = str(target["anchor_label"])
            elif scope == "board":
                if not cleaned_title:
                    cleaned_title = fallback_thread_title("", cleaned)

            norm_anchor_label = normalize_anchor_label(anchor_label) if anchor_label else ""

            cursor = conn.execute(
                """
                INSERT INTO community_posts (
                    scope, title, resource_key, resource_path, resource_title, anchor_id, anchor_label,
                    parent_id, reply_to_id, reply_to_qq, reply_to_nickname,
                    author_qq, author_nickname, author_avatar, author_role_badge,
                    content, is_pinned, like_count, last_reply_at, last_reply_nickname, status, created_at, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0, ?, ?, 'active', ?, ?)
                """,
                (
                    scope,
                    cleaned_title,
                    res_key,
                    resource_path,
                    resource_title,
                    anchor_id.strip()[:80],
                    norm_anchor_label,
                    parent_id,
                    actual_reply_id,
                    reply_to_qq,
                    reply_to_nickname,
                    author_qq,
                    author_nick,
                    str(author.get("avatar_url") or ""),
                    author_badge,
                    cleaned,
                    now if parent_id is None else 0,
                    author_nick if parent_id is None else "",
                    now,
                    now,
                ),
            )
            post_id = int(cursor.lastrowid)

            # 如果是跟帖回复，更新主楼的最后回复时间与最后回复人（经典论坛顶帖机制，确保同秒回复也严格大于先前时间戳）
            if parent_id is not None:
                conn.execute(
                    """
                    UPDATE community_posts
                    SET last_reply_at = MAX(?, last_reply_at + 1), last_reply_nickname = ?, updated_at = ?
                    WHERE id = ?
                    """,
                    (now, author_nick, now, parent_id),
                )

            row = conn.execute("SELECT * FROM community_posts WHERE id = ?", (post_id,)).fetchone()
            return self._serialize_post(row, viewer_qq=author_qq, viewer_is_mod=bool(author_badge))

    def get_board_thread_detail(
        self,
        *,
        thread_id: int,
        viewer_qq: str | None = None,
        viewer_is_mod: bool = False,
    ) -> dict[str, Any]:
        with self._conn() as conn:
            root = conn.execute(
                "SELECT * FROM community_posts WHERE id = ? AND scope = 'board' AND parent_id IS NULL AND status = 'active'",
                (thread_id,),
            ).fetchone()
            if not root:
                raise HTTPException(404, "该论坛主题不存在或已被删除")
            reply_rows = conn.execute(
                """
                SELECT * FROM community_posts
                WHERE parent_id = ? AND status = 'active'
                ORDER BY created_at ASC, id ASC
                """,
                (thread_id,),
            ).fetchall()
            all_ids = [int(root["id"]), *[int(r["id"]) for r in reply_rows]]
            liked_ids = self._liked_post_ids(conn, viewer_qq, all_ids)

            thread = self._serialize_post(
                root,
                viewer_qq=viewer_qq,
                viewer_is_mod=viewer_is_mod,
                liked_ids=liked_ids,
            )
            thread["floor"] = 1
            replies = []
            for idx, r in enumerate(reply_rows, start=2):
                rep = self._serialize_post(
                    r,
                    viewer_qq=viewer_qq,
                    viewer_is_mod=viewer_is_mod,
                    liked_ids=liked_ids,
                )
                rep["floor"] = idx
                replies.append(rep)
            thread["replies"] = replies
            thread["reply_count"] = len(replies)
            return {"thread": thread}

    def list_threads(
        self,
        *,
        scope: Literal["board", "resource"],
        resource_path: str = "",
        resource_title: str = "",
        anchor_id: str = "",
        anchor_label: str = "",
        q: str = "",
        limit: int = 30,
        offset: int = 0,
        sort: Literal["latest", "created", "oldest", "hot"] = "latest",
        include_replies: bool = True,
        viewer_qq: str | None = None,
        viewer_is_mod: bool = False,
    ) -> dict[str, Any]:
        limit = max(1, min(limit, 100))
        offset = max(0, offset)

        with self._conn() as conn:
            where = ["scope = ?", "status = 'active'", "parent_id IS NULL"]
            params: list[Any] = [scope]

            canonical_key = ""
            if scope == "resource":
                canonical_key = self.resolve_canonical_key_with_conn(conn, resource_path, resource_title)
                where.append("(resource_key = ? OR resource_path = ?)")
                params.extend([canonical_key, resource_path])
                if anchor_id or anchor_label:
                    norm_lbl = normalize_anchor_label(anchor_label) if anchor_label else ""
                    if anchor_id and norm_lbl:
                        where.append("(anchor_id = ? OR (anchor_label != '' AND anchor_label = ?))")
                        params.extend([anchor_id, norm_lbl])
                    elif anchor_id:
                        where.append("anchor_id = ?")
                        params.append(anchor_id)
                    elif norm_lbl:
                        where.append("anchor_label = ?")
                        params.append(norm_lbl)
            elif q.strip():
                needle = f"%{q.strip()}%"
                where.append("(title LIKE ? OR content LIKE ? OR author_nickname LIKE ?)")
                params.extend([needle, needle, needle])

            where_sql = " AND ".join(where)
            if scope == "board":
                if sort == "created":
                    order_sql = "is_pinned DESC, created_at DESC, id DESC"
                elif sort == "oldest":
                    order_sql = "is_pinned DESC, created_at ASC, id ASC"
                elif sort == "hot":
                    order_sql = "is_pinned DESC, like_count DESC, last_reply_at DESC, id DESC"
                else:
                    # 默认论坛按最后回复顶帖排序
                    order_sql = "is_pinned DESC, last_reply_at DESC, id DESC"
            else:
                if sort == "oldest":
                    order_sql = "is_pinned DESC, created_at ASC, id ASC"
                elif sort == "hot":
                    order_sql = "is_pinned DESC, like_count DESC, created_at DESC, id DESC"
                else:
                    order_sql = "is_pinned DESC, created_at DESC, id DESC"

            total_roots = int(
                conn.execute(f"SELECT COUNT(*) AS c FROM community_posts WHERE {where_sql}", params).fetchone()["c"]
            )
            if scope == "resource":
                total_all = int(
                    conn.execute(
                        "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'resource' AND (resource_key = ? OR resource_path = ?) AND status = 'active'",
                        (canonical_key, resource_path),
                    ).fetchone()["c"]
                )
                anchor_rows = conn.execute(
                    """
                    SELECT anchor_id, anchor_label, COUNT(*) AS c
                    FROM community_posts
                    WHERE scope = 'resource' AND (resource_key = ? OR resource_path = ?) AND status = 'active' AND (anchor_id != '' OR anchor_label != '')
                    GROUP BY COALESCE(NULLIF(anchor_label, ''), anchor_id)
                    """,
                    (canonical_key, resource_path),
                ).fetchall()
                anchor_counts: dict[str, dict[str, Any]] = {}
                label_counts: dict[str, int] = {}
                for r in anchor_rows:
                    aid = str(r["anchor_id"] or "")
                    albl = str(r["anchor_label"] or "")
                    cnt = int(r["c"])
                    key = aid or albl
                    anchor_counts[key] = {"count": cnt, "label": albl or aid, "anchor_id": aid}
                    if albl:
                        label_counts[albl] = label_counts.get(albl, 0) + cnt
            else:
                total_all = int(
                    conn.execute(
                        "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'board' AND status = 'active'"
                    ).fetchone()["c"]
                )
                anchor_counts = {}
                label_counts = {}

            root_rows = conn.execute(
                f"SELECT * FROM community_posts WHERE {where_sql} ORDER BY {order_sql} LIMIT ? OFFSET ?",
                [*params, limit, offset],
            ).fetchall()

            root_ids = [int(r["id"]) for r in root_rows]
            reply_counts_map: dict[int, int] = {}
            replies_by_parent: dict[int, list[dict[str, Any]]] = {}

            if root_ids:
                placeholders = ",".join("?" for _ in root_ids)
                if include_replies:
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
                    for pid, reps in replies_by_parent.items():
                        reply_counts_map[pid] = len(reps)
                else:
                    liked_ids = self._liked_post_ids(conn, viewer_qq, root_ids)
                    cnt_rows = conn.execute(
                        f"""
                        SELECT parent_id, COUNT(*) AS c
                        FROM community_posts
                        WHERE status = 'active' AND parent_id IN ({placeholders})
                        GROUP BY parent_id
                        """,
                        root_ids,
                    ).fetchall()
                    reply_counts_map = {int(r["parent_id"]): int(r["c"]) for r in cnt_rows}
            else:
                liked_ids = set()

            items = []
            for r in root_rows:
                pid = int(r["id"])
                item = self._serialize_post(
                    r,
                    viewer_qq=viewer_qq,
                    viewer_is_mod=viewer_is_mod,
                    liked_ids=liked_ids,
                )
                if include_replies:
                    item["replies"] = replies_by_parent.get(pid, [])
                item["reply_count"] = reply_counts_map.get(pid, 0)
                items.append(item)

            return {
                "items": items,
                "count": len(items),
                "total_threads": total_roots,
                "total_posts": total_all,
                "resource_key": canonical_key,
                "anchor_counts": anchor_counts,
                "label_counts": label_counts,
                "limit": limit,
                "offset": offset,
            }

    def list_recent_resource_comments(
        self,
        *,
        limit: int = 16,
        viewer_qq: str | None = None,
        viewer_is_mod: bool = False,
        path_resolver: Callable[[str, str], tuple[str, str] | None] | None = None,
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
            out: list[dict[str, Any]] = []
            for r in rows:
                item = self._serialize_post(
                    r,
                    viewer_qq=viewer_qq,
                    viewer_is_mod=viewer_is_mod,
                    liked_ids=liked_ids,
                )
                if path_resolver:
                    resolved = path_resolver(str(r["resource_path"]), str(r["resource_key"] or ""))
                    if resolved:
                        live_path, live_title = resolved
                        if live_path != r["resource_path"] or live_title != r["resource_title"]:
                            conn.execute(
                                "UPDATE community_posts SET resource_path = ?, resource_title = ? WHERE id = ?",
                                (live_path, live_title, int(r["id"])),
                            )
                            item["resource_path"] = live_path
                            item["resource_title"] = live_title
                out.append(item)
            return out

    def toggle_like(self, *, post_id: int, user_qq: str) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            post = conn.execute(
                "SELECT id, status FROM community_posts WHERE id = ?",
                (post_id,),
            ).fetchone()
            if not post or post["status"] != "active":
                raise HTTPException(404, "该帖子或评论不存在")

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
                raise HTTPException(404, "该帖子或评论不存在")
            if post["parent_id"] is not None:
                raise HTTPException(400, "只能置顶主楼层，不能置顶跟帖回复")
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
                raise HTTPException(404, "该帖子或评论不存在")
            if post["status"] == "deleted":
                return {"ok": True, "post_id": post_id, "already_deleted": True}
            if not is_mod and (not viewer_qq or str(post["author_qq"]) != viewer_qq):
                raise HTTPException(403, "只能删除自己发表的帖子或评论")

            parent_id = post["parent_id"]
            conn.execute(
                """
                UPDATE community_posts
                SET status = 'deleted', deleted_by = ?, is_pinned = 0, updated_at = ?
                WHERE id = ? OR parent_id = ?
                """,
                (actor, now, post_id, post_id),
            )
            # 若删除的是某主题下的最后一楼回复，重新校准主楼的 last_reply_at
            if parent_id is not None:
                latest_rep = conn.execute(
                    "SELECT created_at, author_nickname FROM community_posts WHERE parent_id = ? AND status = 'active' ORDER BY created_at DESC, id DESC LIMIT 1",
                    (int(parent_id),),
                ).fetchone()
                if latest_rep:
                    conn.execute(
                        "UPDATE community_posts SET last_reply_at = ?, last_reply_nickname = ? WHERE id = ?",
                        (int(latest_rep["created_at"]), str(latest_rep["author_nickname"]), int(parent_id)),
                    )
                else:
                    conn.execute(
                        "UPDATE community_posts SET last_reply_at = created_at, last_reply_nickname = author_nickname WHERE id = ?",
                        (int(parent_id),),
                    )
            return {"ok": True, "post_id": post_id}

    def stats_overview(self) -> dict[str, Any]:
        now = now_ts()
        with self._conn() as conn:
            board_threads = int(
                conn.execute(
                    "SELECT COUNT(*) AS c FROM community_posts WHERE scope = 'board' AND parent_id IS NULL AND status = 'active'"
                ).fetchone()["c"]
            )
            board_posts = int(
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
                "board_thread_count": board_threads,
                "board_post_count": board_posts,
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
            where.append("(title LIKE ? OR content LIKE ? OR author_nickname LIKE ? OR resource_title LIKE ? OR author_qq LIKE ?)")
            params.extend([needle, needle, needle, needle, needle])

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


class BotVerifyCodeRequest(BaseModel):
    code: str = Field(min_length=6, max_length=6)
    open_user_id: str = Field(min_length=4, max_length=120)
    open_group_id: str = Field(min_length=1, max_length=120)
    sealdice_nickname: str = Field(default="", max_length=80)
    bot_app_id: str = Field(default="", max_length=32)


class WebChallengeConsumeRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    nickname: str = Field(default="", max_length=40)


class UpdateNicknameRequest(BaseModel):
    nickname: str = Field(min_length=1, max_length=40)


class CreateBoardPostRequest(BaseModel):
    title: str = Field(default="", max_length=100)
    content: str = Field(min_length=1, max_length=4000)
    reply_to_id: Optional[int] = None


class CreateResourceCommentRequest(BaseModel):
    resource_path: str = Field(min_length=1, max_length=300)
    content: str = Field(min_length=1, max_length=1600)
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
    title: str = Field(default="", max_length=100)
    content: str = Field(min_length=1, max_length=4000)
    nickname: str = Field(default="界域终端公告", max_length=40)
    pinned: bool = True


def build_community_router(
    *,
    store_getter: Callable[[], CommunityStore],
    upstream_getter: Callable[[], UpstreamAuthClient],
    secret_getter: Callable[[], bytes],
    is_admin_request: Callable[[Request], bool],
    resolve_public_resource: Callable[[str], tuple[str, str]],
    find_live_resource_by_key: Callable[[str, str], tuple[str, str] | None] | None = None,
) -> APIRouter:
    """构建挂载在 /api 下的轮回者论坛、防断链评论与 QQ 登录路由器。"""
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
                "verification_code": sess.get("pending_code") or "",
                "expires_at": sess.get("pending_expires_at"),
            }
        auth_ready = bot_auth_configured() or upstream_getter().configured()
        if not user:
            return {
                "authenticated": False,
                "auth_configured": auth_ready,
                "auth_mode": "official_bot_code",
                "bot_name": OFFICIAL_BOT_NAME,
                "bot_app_id": OFFICIAL_BOT_APP_ID,
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
            "auth_configured": auth_ready,
            "auth_mode": "official_bot_code",
            "bot_name": OFFICIAL_BOT_NAME,
            "bot_app_id": OFFICIAL_BOT_APP_ID,
            "bot_qq": EXPECTED_BOT_QQ,
            "group_id": EXPECTED_GROUP_ID,
            "admin_authenticated": admin_cookie,
            "is_moderator": is_mod,
            "pending_challenge": pending,
            "user": {
                "qq": qq_val,
                "qq_masked": mask_qq(qq_val),
                "nickname": user["nickname"],
                "avatar_url": user["avatar_url"] or default_avatar_for_user(qq_val),
                "card_count": int(user.get("card_count") or 0),
                "badge": badge,
                "is_banned": banned,
                "banned_until": None if not banned else (-1 if until == -1 else ts_to_iso(until)),
                "ban_reason": user.get("ban_reason") or "" if banned else "",
            },
        }

    def require_active_member(request: Request) -> tuple[dict[str, Any], bool, str]:
        _sess, user, admin_cookie, is_mod = viewer_context(request)
        if not user:
            raise HTTPException(401, "请先在主群发送 /登录 获取 6 位登录码完成认证")
        if user_is_banned(user):
            reason = (user.get("ban_reason") or "").strip()
            detail = f"当前轮回者账号已被限制发言（原因：{reason}）" if reason else "当前轮回者账号已被限制发言"
            raise HTTPException(403, detail)
        qq_val = str(user["qq"])
        badge = "维护者" if (qq_val in parse_admin_qqs() or admin_cookie) else ("协管" if user.get("role") == "moderator" else "")
        return user, is_mod, badge

    def require_moderator(request: Request) -> tuple[str, dict[str, Any] | None]:
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
            verification_code=verification_code,
            expires_at=expires_at,
        )
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
            raise HTTPException(503, ERROR_MESSAGES["unable_to_verify"])

        if http_status != 200:
            err_code = str(data.get("error") or "")
            if http_status in {403, 404, 410} or err_code in {"already_consumed", "not_in_group", "no_uploaded_card", "challenge_expired"}:
                store_getter().clear_pending_challenge(raw_sid, secret_getter())
            mapped = http_status if http_status in {400, 403, 404, 409, 410, 503} else 503
            raise HTTPException(mapped, ERROR_MESSAGES.get(err_code, ERROR_MESSAGES["unable_to_verify"]))

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

    @router.post("/auth/web-challenges")
    def auth_web_challenge_begin(request: Request, response: Response, renew: bool = False):
        no_store(response)
        if not bot_auth_configured():
            raise HTTPException(503, "服务器尚未配置机器人桥接密钥")
        store = store_getter()
        secret = secret_getter()
        sid, session, _ = store.get_or_create_session(request.cookies.get(BROWSER_COOKIE), secret)
        if session["user_qq"]:
            raise HTTPException(409, "当前浏览器已登录")
        view = store.begin_web_challenge(session["sid_hash"], secret, renew=renew)
        set_browser_cookie(response, request, sid)
        return view

    @router.get("/auth/web-challenges/current")
    def auth_web_challenge_current(request: Request, response: Response):
        no_store(response)
        store = store_getter()
        secret = secret_getter()
        session = store.get_session(request.cookies.get(BROWSER_COOKIE), secret)
        if not session or session["user_qq"]:
            return {"status": "expired"}
        return store.current_web_challenge(session["sid_hash"], secret)

    @router.post("/auth/bot/verify-code")
    def auth_bot_verify_code(body: BotVerifyCodeRequest, request: Request, response: Response):
        no_store(response)
        expected = bot_bridge_token()
        if not expected:
            raise HTTPException(503, "服务器尚未配置机器人桥接密钥")
        scheme, _, provided = request.headers.get("Authorization", "").partition(" ")
        if scheme.lower() != "bearer" or not hmac.compare_digest(provided.strip().encode(), expected.encode()):
            raise HTTPException(401, "机器人桥接鉴权未通过")
        group = normalize_open_group_token(body.open_group_id)
        allowed = parse_allowed_open_groups()
        if not OPENID_HEX_RE.fullmatch(group) or (group not in allowed and "*" not in allowed):
            raise HTTPException(403, "仅限在宏观界域主群内核验身份")
        if body.bot_app_id and body.bot_app_id != OFFICIAL_BOT_APP_ID:
            raise HTTPException(403, "机器人身份与站点配置不符")
        try:
            user_key = f"oid:{normalize_open_user_id(body.open_user_id)}"
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        store_getter().verify_web_challenge(body.code, user_key, body.sealdice_nickname)
        return {"ok": True, "status": "verified"}

    @router.post("/auth/web-challenges/consume")
    def auth_web_challenge_consume(body: WebChallengeConsumeRequest, request: Request, response: Response):
        no_store(response)
        store = store_getter()
        secret = secret_getter()
        sid = request.cookies.get(BROWSER_COOKIE)
        session = store.get_session(sid, secret)
        if not sid or not session or session["user_qq"]:
            raise HTTPException(409, "当前浏览器没有可确认的登录验证")
        user = store.consume_web_challenge(
            session["sid_hash"], secret, confirmation=request.headers.get("X-Login-Confirmation", ""), nickname=body.nickname
        )
        set_browser_cookie(response, request, sid)
        admin_cookie = is_admin_request(request)
        return {
            "ok": True,
            **format_viewer(store.get_session(sid, secret), user, admin_cookie=admin_cookie,
                            is_mod=admin_cookie or user_is_qq_moderator(user)),
        }

    @router.patch("/auth/profile")
    def auth_update_profile(body: UpdateNicknameRequest, request: Request, response: Response):
        """已登录用户自助修改论坛展示昵称（同步更新历史发帖展示昵称）。"""
        no_store(response)
        sess, user, admin_cookie, _is_mod = viewer_context(request)
        if not user:
            raise HTTPException(401, "请先登录后再修改昵称")
        if user_is_banned(user):
            raise HTTPException(403, "当前账号处于禁言状态，暂不可修改昵称")
        updated_user = store_getter().update_user_nickname(str(user["qq"]), body.nickname)
        is_mod = admin_cookie or user_is_qq_moderator(updated_user)
        return {
            "ok": True,
            **format_viewer(sess, updated_user, admin_cookie=admin_cookie, is_mod=is_mod),
        }

    @router.post("/auth/logout")
    def auth_logout(request: Request, response: Response):
        no_store(response)
        raw_sid = request.cookies.get(BROWSER_COOKIE)
        store_getter().logout_session(raw_sid, secret_getter())
        response.delete_cookie(BROWSER_COOKIE, path="/")
        return {"ok": True}

    # ---------------- 2. 轮回者公共论坛（纯主题列表 + 楼层详情） ----------------

    @router.get("/community/home")
    def community_home(
        request: Request,
        response: Response,
        q: str = "",
        limit: int = 35,
        offset: int = 0,
        sort: Literal["latest", "created", "oldest", "hot"] = "latest",
    ):
        no_store(response)
        _sess, user, _admin_cookie, is_mod = viewer_context(request)
        viewer_qq = str(user["qq"]) if user else None
        board = store_getter().list_threads(
            scope="board",
            q=q,
            limit=limit,
            offset=offset,
            sort=sort,
            include_replies=False,
            viewer_qq=viewer_qq,
            viewer_is_mod=is_mod,
        )
        recent_comments = store_getter().list_recent_resource_comments(
            limit=15,
            viewer_qq=viewer_qq,
            viewer_is_mod=is_mod,
            path_resolver=find_live_resource_by_key,
        )
        stats = store_getter().stats_overview()
        return {
            "board": board,
            "recent_resource_comments": recent_comments,
            "stats": stats,
        }

    @router.get("/community/threads/{thread_id}")
    def get_board_thread(thread_id: int, request: Request, response: Response):
        no_store(response)
        _sess, user, _admin_cookie, is_mod = viewer_context(request)
        viewer_qq = str(user["qq"]) if user else None
        return store_getter().get_board_thread_detail(
            thread_id=thread_id,
            viewer_qq=viewer_qq,
            viewer_is_mod=is_mod,
        )

    @router.post("/community/board")
    def create_board_post(body: CreateBoardPostRequest, request: Request, response: Response):
        no_store(response)
        user, _is_mod, badge = require_active_member(request)
        if body.reply_to_id is None and not (body.title or "").strip() and len((body.content or "").strip()) < 2:
            raise HTTPException(400, "请输入主题标题与正文内容")
        try:
            post = store_getter().create_post(
                scope="board",
                author=user,
                author_badge=badge,
                title=body.title,
                content=body.content,
                reply_to_id=body.reply_to_id,
            )
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, "post": post}

    # ---------------- 3. 序列档案评论（抗序号重排 + 能力卡双重锚点） ----------------

    @router.get("/community/comments")
    def list_resource_comments(
        request: Request,
        response: Response,
        path: str = Query(..., min_length=1),
        anchor_id: str = "",
        anchor_label: str = "",
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
            resource_title=title,
            anchor_id=anchor_id.strip(),
            anchor_label=anchor_label.strip(),
            limit=limit,
            offset=offset,
            sort=sort,
            include_replies=True,
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
        try:
            target_qq = normalize_user_key(body.qq)
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        if target_qq in parse_admin_qqs() and not is_admin_request(request):
            raise HTTPException(403, "不能封禁站点维护者白名单账号")

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

    @router.post("/community/moderation/users/{qq:path}/role")
    def moderation_set_role(qq: str, body: AdminRoleRequest, request: Request, response: Response):
        no_store(response)
        _sess, user, admin_cookie, _is_mod = viewer_context(request)
        is_super = admin_cookie or bool(user and str(user["qq"]) in parse_admin_qqs())
        if not is_super:
            raise HTTPException(403, "仅站点维护者可调整协管权限")
        try:
            target_qq = normalize_user_key(qq)
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        row = store_getter().set_user_role(target_qq, body.role)
        return {"ok": True, "qq": row["qq"], "role": row["role"]}

    @router.post("/community/moderation/announcement")
    def moderation_post_announcement(body: AdminAnnouncementRequest, request: Request, response: Response):
        no_store(response)
        _actor, user = require_moderator(request)
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
                title=body.title.strip() or fallback_thread_title("", body.content),
                content=body.content,
                bypass_cooldown=True,
            )
            if body.pinned:
                post = store_getter().set_pinned(post_id=int(post["id"]), pinned=True)
        except ValueError as exc:
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, "post": post}

    return router
