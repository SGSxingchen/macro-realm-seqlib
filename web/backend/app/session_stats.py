from __future__ import annotations

import hashlib
import json
import math
import os
import socket
import sqlite3
import time
import urllib.error
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from contextlib import contextmanager
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Iterator, Protocol


@dataclass(frozen=True)
class SessionTextFile:
    filename: str
    content: str


class SessionStatsExtractor(Protocol):
    def extract(self, *, filename: str, content: str, month: str) -> dict[str, Any]:
        ...


class OpenAIExtractor:
    def __init__(
        self,
        model_name: str,
        *,
        api_key: str | None = None,
        base_url: str | None = None,
        api_mode: str | None = None,
        timeout: int | None = None,
        urlopen=urllib.request.urlopen,
    ) -> None:
        self.model_name = model_name
        self.api_key = api_key
        self.base_url = (base_url or os.getenv("OPENAI_BASE_URL") or "https://api.openai.com/v1").rstrip("/")
        self.api_mode = (api_mode or os.getenv("OPENAI_API_MODE") or "auto").strip().lower()
        self.timeout = int(timeout or os.getenv("OPENAI_TIMEOUT_SECONDS") or 180)
        self.request_retries = max(0, int(os.getenv("OPENAI_REQUEST_RETRIES") or 1))
        self.urlopen = urlopen

    def extract(self, *, filename: str, content: str, month: str) -> dict[str, Any]:
        api_key = self.api_key or os.getenv("OPENAI_API_KEY") or os.getenv("THIRD_PARTY_API_KEY")
        if not api_key:
            raise RuntimeError("未配置 OPENAI_API_KEY 或 THIRD_PARTY_API_KEY")

        prompt = self._prompt(filename=filename, content=content, month=month)
        if self._use_chat_completions():
            return self._extract_chat_completions(api_key, prompt)
        return self._extract_responses(api_key, prompt)

    def _use_chat_completions(self) -> bool:
        if self.api_mode in {"chat", "chat_completions", "chat-completions"}:
            return True
        if self.api_mode in {"responses", "response"}:
            return False
        return self.model_name.lower().startswith("deepseek")

    def _extract_responses(self, api_key: str, prompt: str) -> dict[str, Any]:
        payload = {
            "model": self.model_name,
            "input": [
                {
                    "role": "user",
                    "content": [{"type": "input_text", "text": prompt}],
                }
            ],
            "text": {
                "format": {
                    "type": "json_schema",
                    "name": "session_summary",
                    "strict": True,
                    "schema": self._schema(),
                }
            },
        }
        try:
            raw = self._post_json(f"{self.base_url}/responses", api_key, payload)
        except urllib.error.HTTPError as exc:
            detail = exc.read().decode("utf-8", errors="replace")
            raise RuntimeError(f"OpenAI 请求失败：HTTP {exc.code} {detail}") from exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"OpenAI 请求失败：{exc.reason}") from exc

        return self._parse_result(raw, self._response_text)

    def _extract_chat_completions(self, api_key: str, prompt: str) -> dict[str, Any]:
        payload = self._chat_payload(prompt, schema_mode=True)
        try:
            raw = self._post_json(f"{self.base_url}/chat/completions", api_key, payload)
        except urllib.error.HTTPError as exc:
            if exc.code not in {400, 422}:
                detail = exc.read().decode("utf-8", errors="replace")
                raise RuntimeError(f"OpenAI 请求失败：HTTP {exc.code} {detail}") from exc
            payload = self._chat_payload(prompt, schema_mode=False)
            try:
                raw = self._post_json(f"{self.base_url}/chat/completions", api_key, payload)
            except urllib.error.HTTPError as fallback_exc:
                detail = fallback_exc.read().decode("utf-8", errors="replace")
                raise RuntimeError(f"OpenAI 请求失败：HTTP {fallback_exc.code} {detail}") from fallback_exc
        except urllib.error.URLError as exc:
            raise RuntimeError(f"OpenAI 请求失败：{exc.reason}") from exc
        return self._parse_result(raw, self._chat_response_text)

    def _post_json(self, url: str, api_key: str, payload: dict[str, Any]) -> str:
        request = urllib.request.Request(
            url,
            data=json.dumps(payload, ensure_ascii=False).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {api_key}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        last_timeout: BaseException | None = None
        for attempt in range(self.request_retries + 1):
            try:
                with self.urlopen(request, timeout=self.timeout) as response:
                    return response.read().decode("utf-8")
            except (TimeoutError, socket.timeout, urllib.error.URLError) as exc:
                if not self._is_timeout_error(exc):
                    raise
                last_timeout = exc
                if attempt >= self.request_retries:
                    break
                time.sleep(0.8 * (attempt + 1))
        raise RuntimeError(f"模型请求超时：{self.timeout} 秒内没有返回，已重试 {self.request_retries} 次") from last_timeout

    @staticmethod
    def _is_timeout_error(exc: BaseException) -> bool:
        if isinstance(exc, (TimeoutError, socket.timeout)):
            return True
        if isinstance(exc, urllib.error.URLError):
            reason = exc.reason
            if isinstance(reason, BaseException):
                return OpenAIExtractor._is_timeout_error(reason)
            return "timed out" in str(reason).lower() or "timeout" in str(reason).lower()
        return False

    def _chat_payload(self, prompt: str, *, schema_mode: bool) -> dict[str, Any]:
        response_format: dict[str, Any]
        if schema_mode:
            response_format = {
                "type": "json_schema",
                "json_schema": {
                    "name": "session_summary",
                    "strict": True,
                    "schema": self._schema(),
                },
            }
        else:
            response_format = {"type": "json_object"}
        return {
            "model": self.model_name,
            "messages": [{"role": "user", "content": prompt}],
            "response_format": response_format,
        }

    @staticmethod
    def _prompt(*, filename: str, content: str, month: str) -> str:
        return (
            "从下面的 TRPG 结团文本中抽取结构化统计数据，只返回 JSON。"
            "字段：title 字符串、duration_hours 数字、kp 对象{name, qq}、"
            "players 数组[{name, qq}]、confidence 数字、warnings 字符串数组。"
            "无法识别 QQ 时填 null；duration_hours 必须是数字小时。"
            f"\n月份：{month}\n文件名：{filename}\n原文：\n{content}"
        )

    @staticmethod
    def _response_text(data: dict[str, Any]) -> str:
        if isinstance(data.get("output_text"), str):
            return data["output_text"]
        chunks: list[str] = []
        for item in data.get("output", []):
            if not isinstance(item, dict):
                continue
            for content in item.get("content", []):
                if isinstance(content, dict) and isinstance(content.get("text"), str):
                    chunks.append(content["text"])
        return "".join(chunks)

    @staticmethod
    def _chat_response_text(data: dict[str, Any]) -> str:
        choices = data.get("choices")
        if not isinstance(choices, list) or not choices:
            return ""
        first = choices[0]
        if not isinstance(first, dict):
            return ""
        message = first.get("message")
        if not isinstance(message, dict):
            return ""
        content = message.get("content")
        return content if isinstance(content, str) else ""

    @staticmethod
    def _parse_result(raw: str, text_reader: Callable[[dict[str, Any]], str]) -> dict[str, Any]:
        try:
            data = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise RuntimeError("OpenAI 响应不是合法 JSON") from exc
        text = text_reader(data)
        if not text:
            raise RuntimeError("OpenAI 返回为空")
        try:
            result = json.loads(text)
        except json.JSONDecodeError as exc:
            raise RuntimeError("OpenAI 返回不是合法 JSON") from exc
        if not isinstance(result, dict):
            raise RuntimeError("OpenAI 返回 JSON 顶层不是对象")
        return result

    @staticmethod
    def _schema() -> dict[str, Any]:
        return {
            "type": "object",
            "additionalProperties": False,
            "required": ["title", "duration_hours", "kp", "players", "confidence", "warnings"],
            "properties": {
                "title": {"type": "string"},
                "duration_hours": {"type": "number"},
                "kp": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": ["name", "qq"],
                    "properties": {
                        "name": {"type": "string"},
                        "qq": {"type": ["string", "null"]},
                    },
                },
                "players": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["name", "qq"],
                        "properties": {
                            "name": {"type": "string"},
                            "qq": {"type": ["string", "null"]},
                        },
                    },
                },
                "confidence": {"type": "number"},
                "warnings": {"type": "array", "items": {"type": "string"}},
            },
        }


class SessionStatsService:
    def __init__(self, db_path: Path) -> None:
        self.db_path = Path(db_path)
        self.db_path.parent.mkdir(parents=True, exist_ok=True)
        self._init_db()

    def import_text_files(
        self,
        month: str,
        files: list[SessionTextFile],
        extractor: SessionStatsExtractor | None,
        model_name: str,
        concurrency: int = 1,
        progress_callback: Callable[[dict[str, Any], int, int, int], None] | None = None,
    ) -> dict[str, Any]:
        actual_extractor = extractor or OpenAIExtractor(model_name)
        if int(concurrency or 1) > 1:
            return self._import_text_files_concurrently(
                month=month,
                files=files,
                extractor=actual_extractor,
                model_name=model_name,
                concurrency=concurrency,
                progress_callback=progress_callback,
            )
        items: list[dict[str, Any]] = []
        success_count = 0
        failure_count = 0
        skip_count = 0

        with self._connect() as conn:
            batch_id = self._create_batch(conn, month, model_name, len(files))

        for file in files:
            content_hash = self._content_hash(file.content)
            with self._connect() as conn:
                existing = self._find_existing_session_by_content(conn, month, content_hash)
                if existing:
                    item = self._skip_item(file.filename, existing)
                    items.append(item)
                    skip_count += 1
                    if progress_callback:
                        progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)
                    continue
            with self._connect() as conn:
                if self._is_duplicate(conn, month, file.filename, content_hash):
                    reason = "重复导入：同一月份、文件名和内容已经存在"
                    self._record_error(conn, batch_id, month, file.filename, content_hash, reason, model_name, None)
                    item = {"filename": file.filename, "success": False, "reason": reason}
                    items.append(item)
                    failure_count += 1
                    if progress_callback:
                        progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)
                    continue

            raw_payload: Any = None
            try:
                extracted = actual_extractor.extract(filename=file.filename, content=file.content, month=month)
                raw_payload = extracted
                reason = self._validation_error(extracted)
                if reason:
                    raise ValueError(reason)
                with self._connect() as conn:
                    existing = self._find_existing_session_by_content(conn, month, content_hash)
                    if existing:
                        item = self._skip_item(file.filename, existing)
                        items.append(item)
                        skip_count += 1
                        if progress_callback:
                            progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)
                        continue
                    session_id = self._insert_session(
                        conn,
                        batch_id=batch_id,
                        month=month,
                        file=file,
                        content_hash=content_hash,
                        payload=extracted,
                        model_name=model_name,
                    )
            except Exception as exc:
                reason = str(exc) or exc.__class__.__name__
                with self._connect() as conn:
                    self._record_error(conn, batch_id, month, file.filename, content_hash, reason, model_name, raw_payload)
                item = {"filename": file.filename, "success": False, "reason": reason}
                items.append(item)
                failure_count += 1
                if progress_callback:
                    progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)
                continue

            item = {"filename": file.filename, "success": True, "session_id": session_id}
            items.append(item)
            success_count += 1
            if progress_callback:
                progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)

        with self._connect() as conn:
            self._finish_batch(conn, batch_id, success_count, failure_count)

        return {"success_count": success_count, "failure_count": failure_count, "skip_count": skip_count, "items": items}

    def _import_text_files_concurrently(
        self,
        *,
        month: str,
        files: list[SessionTextFile],
        extractor: SessionStatsExtractor,
        model_name: str,
        concurrency: int,
        progress_callback: Callable[[dict[str, Any], int, int, int], None] | None = None,
    ) -> dict[str, Any]:
        items: list[dict[str, Any]] = []
        success_count = 0
        failure_count = 0
        skip_count = 0
        worker_count = max(1, min(int(concurrency or 1), len(files) or 1))

        with self._connect() as conn:
            batch_id = self._create_batch(conn, month, model_name, len(files))

        def process_file(file: SessionTextFile) -> dict[str, Any]:
            content_hash = self._content_hash(file.content)
            with self._connect() as conn:
                existing = self._find_existing_session_by_content(conn, month, content_hash)
                if existing:
                    return self._skip_item(file.filename, existing)
            raw_payload: Any = None
            try:
                with self._connect() as conn:
                    if self._is_duplicate(conn, month, file.filename, content_hash):
                        reason = "重复导入：同一月份、文件名和内容已经存在"
                        self._record_error(conn, batch_id, month, file.filename, content_hash, reason, model_name, None)
                        return {"filename": file.filename, "success": False, "reason": reason}

                extracted = extractor.extract(filename=file.filename, content=file.content, month=month)
                raw_payload = extracted
                reason = self._validation_error(extracted)
                if reason:
                    raise ValueError(reason)
                with self._connect() as conn:
                    existing = self._find_existing_session_by_content(conn, month, content_hash)
                    if existing:
                        return self._skip_item(file.filename, existing)
                with self._connect() as conn:
                    session_id = self._insert_session(
                        conn,
                        batch_id=batch_id,
                        month=month,
                        file=file,
                        content_hash=content_hash,
                        payload=extracted,
                        model_name=model_name,
                    )
                return {"filename": file.filename, "success": True, "session_id": session_id}
            except Exception as exc:
                reason = str(exc) or exc.__class__.__name__
                with self._connect() as conn:
                    self._record_error(conn, batch_id, month, file.filename, content_hash, reason, model_name, raw_payload)
                return {"filename": file.filename, "success": False, "reason": reason}

        with ThreadPoolExecutor(max_workers=worker_count) as executor:
            futures = [executor.submit(process_file, file) for file in files]
            for future in as_completed(futures):
                item = future.result()
                items.append(item)
                if item.get("skipped") is True:
                    skip_count += 1
                elif item.get("success") is True:
                    success_count += 1
                else:
                    failure_count += 1
                if progress_callback:
                    progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)

        with self._connect() as conn:
            self._finish_batch(conn, batch_id, success_count, failure_count)

        return {"success_count": success_count, "failure_count": failure_count, "skip_count": skip_count, "items": items}

    def get_overview(self, month: str) -> dict[str, Any]:
        with self._connect() as conn:
            session_count = conn.execute("select count(*) from sessions where month = ?", (month,)).fetchone()[0]
            participant_count = conn.execute(
                """
                select count(*)
                from session_participants sp
                join sessions s on s.id = sp.session_id
                where s.month = ?
                """,
                (month,),
            ).fetchone()[0]
            totals = conn.execute(
                """
                select
                    coalesce(sum(sp.duration_hours), 0) as total_game_hours,
                    coalesce(sum(case when sp.is_host = 1 then sp.duration_hours else 0 end), 0) as total_host_hours
                from session_participants sp
                join sessions s on s.id = sp.session_id
                where s.month = ?
                """,
                (month,),
            ).fetchone()
        return {
            "session_count": int(session_count),
            "participant_count": int(participant_count),
            "total_game_hours": float(totals["total_game_hours"]),
            "total_host_hours": float(totals["total_host_hours"]),
        }

    def list_players(self, month: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                select
                    p.id,
                    p.name,
                    p.qq,
                    count(sp.session_id) as game_count,
                    coalesce(sum(sp.duration_hours), 0) as game_hours,
                    coalesce(sum(sp.reincarnation_count), 0) as reincarnation_count,
                    coalesce(sum(case when sp.is_host = 1 then 1 else 0 end), 0) as host_count,
                    coalesce(sum(case when sp.is_host = 1 then sp.duration_hours else 0 end), 0) as host_hours
                from session_participants sp
                join sessions s on s.id = sp.session_id
                join players p on p.id = sp.player_id
                where s.month = ?
                group by p.id, p.name, p.qq
                order by game_count desc, host_count desc, p.name asc
                """,
                (month,),
            ).fetchall()
        return [self._player_row(row) for row in rows]

    def list_sessions(self, month: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                select
                    s.id,
                    s.month,
                    s.title,
                    s.duration_hours,
                    s.source_filename,
                    s.model_name,
                    s.confidence,
                    s.created_at,
                    kp.name as kp_name,
                    kp.qq as kp_qq,
                    coalesce(sum(case when sp.is_host = 0 then 1 else 0 end), 0) as pl_count
                from sessions s
                join players kp on kp.id = s.kp_player_id
                left join session_participants sp on sp.session_id = s.id
                where s.month = ?
                group by s.id
                order by s.created_at desc, s.id desc
                """,
                (month,),
            ).fetchall()
        return [dict(row) for row in rows]

    def get_session_detail(self, session_id: int) -> dict[str, Any] | None:
        with self._connect() as conn:
            row = conn.execute(
                """
                select
                    s.id,
                    s.month,
                    s.title,
                    s.duration_hours,
                    s.source_filename,
                    s.model_name,
                    s.confidence,
                    s.raw_payload,
                    s.created_at,
                    kp.id as kp_id,
                    kp.name as kp_name,
                    kp.qq as kp_qq
                from sessions s
                join players kp on kp.id = s.kp_player_id
                where s.id = ?
                """,
                (session_id,),
            ).fetchone()
            if row is None:
                return None
            participants = conn.execute(
                """
                select
                    p.id,
                    p.name,
                    p.qq,
                    sp.role,
                    sp.duration_hours,
                    sp.reincarnation_count,
                    sp.is_host
                from session_participants sp
                join players p on p.id = sp.player_id
                where sp.session_id = ?
                order by sp.is_host desc, sp.role asc, p.name asc
                """,
                (session_id,),
            ).fetchall()
        detail = dict(row)
        raw_payload = detail.get("raw_payload")
        if isinstance(raw_payload, str):
            try:
                detail["raw_payload"] = json.loads(raw_payload)
            except json.JSONDecodeError:
                pass
        detail["kp"] = {"id": detail.pop("kp_id"), "name": detail.pop("kp_name"), "qq": detail.pop("kp_qq")}
        detail["participants"] = [
            {
                "id": p["id"],
                "name": p["name"],
                "qq": p["qq"],
                "role": p["role"],
                "duration_hours": float(p["duration_hours"]),
                "reincarnation_count": int(p["reincarnation_count"]),
                "is_host": bool(p["is_host"]),
            }
            for p in participants
        ]
        return detail

    def update_session(self, session_id: int, *, title: str | None = None, duration_hours: float | None = None) -> dict[str, Any] | None:
        updates: list[str] = []
        params: list[Any] = []
        if title is not None:
            clean_title = self._clean_text(title)
            if not clean_title:
                raise ValueError("标题不能为空")
            updates.append("title = ?")
            params.append(clean_title)
        if duration_hours is not None:
            duration = float(duration_hours)
            if not math.isfinite(duration) or duration <= 0:
                raise ValueError("时长必须为正数")
            updates.append("duration_hours = ?")
            params.append(duration)

        if updates:
            with self._connect() as conn:
                params.append(session_id)
                cursor = conn.execute(f"update sessions set {', '.join(updates)} where id = ?", params)
                if cursor.rowcount == 0:
                    return None
                if duration_hours is not None:
                    conn.execute("update session_participants set duration_hours = ? where session_id = ?", (float(duration_hours), session_id))
        return self.get_session_detail(session_id)

    def list_errors(self, month: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                select id, month, filename, content_hash, reason, model_name, raw_payload, created_at
                from import_errors
                where month = ?
                order by created_at desc, id desc
                """,
                (month,),
            ).fetchall()
        return [dict(row) for row in rows]

    def list_import_batches(self, month: str) -> list[dict[str, Any]]:
        with self._connect() as conn:
            rows = conn.execute(
                """
                select id, month, file_count, success_count, failure_count, model_name, created_at
                from import_batches
                where month = ?
                order by created_at desc, id desc
                """,
                (month,),
            ).fetchall()
        return [dict(row) for row in rows]

    def delete_session(self, session_id: int) -> bool:
        with self._connect() as conn:
            cursor = conn.execute("delete from sessions where id = ?", (session_id,))
            return cursor.rowcount > 0

    def cleanup_exact_duplicate_sessions_for_month(self, month: str) -> dict[str, Any]:
        kept: list[dict[str, Any]] = []
        deleted: list[dict[str, Any]] = []
        with self._connect() as conn:
            rows = conn.execute(
                """
                select id, content_hash, source_filename
                from sessions
                where month = ?
                order by content_hash asc, created_at asc, id asc
                """,
                (month,),
            ).fetchall()

            current_hash: str | None = None
            keep_row: sqlite3.Row | None = None
            duplicate_count = 0
            delete_ids: list[int] = []

            def finish_group() -> None:
                nonlocal duplicate_count
                if keep_row is not None and duplicate_count:
                    kept.append(
                        {
                            "content_hash": keep_row["content_hash"],
                            "session_id": int(keep_row["id"]),
                            "source_filename": keep_row["source_filename"],
                            "duplicate_count": duplicate_count,
                        }
                    )
                duplicate_count = 0

            for row in rows:
                if row["content_hash"] != current_hash:
                    finish_group()
                    current_hash = row["content_hash"]
                    keep_row = row
                    continue

                duplicate_count += 1
                session_id = int(row["id"])
                delete_ids.append(session_id)
                deleted.append(
                    {
                        "content_hash": row["content_hash"],
                        "session_id": session_id,
                        "source_filename": row["source_filename"],
                    }
                )
            finish_group()

            if delete_ids:
                placeholders = ",".join("?" for _ in delete_ids)
                conn.execute(f"delete from sessions where id in ({placeholders})", delete_ids)

        return {"deleted_count": len(deleted), "kept": kept, "deleted": deleted}

    def _init_db(self) -> None:
        with self._connect() as conn:
            conn.executescript(
                """
                pragma foreign_keys = on;

                create table if not exists import_batches (
                    id integer primary key autoincrement,
                    month text not null,
                    file_count integer not null default 0,
                    success_count integer not null default 0,
                    failure_count integer not null default 0,
                    model_name text not null,
                    created_at text not null
                );

                create table if not exists players (
                    id integer primary key autoincrement,
                    name text not null,
                    qq text,
                    created_at text not null,
                    updated_at text not null
                );

                create table if not exists sessions (
                    id integer primary key autoincrement,
                    batch_id integer not null references import_batches(id) on delete cascade,
                    month text not null,
                    title text not null,
                    duration_hours real not null,
                    kp_player_id integer not null references players(id),
                    source_filename text not null,
                    content_hash text not null,
                    model_name text not null,
                    confidence real,
                    raw_payload text not null,
                    created_at text not null,
                    unique(month, source_filename, content_hash)
                );

                create table if not exists session_participants (
                    session_id integer not null references sessions(id) on delete cascade,
                    player_id integer not null references players(id),
                    role text not null check(role in ('kp', 'pl')),
                    duration_hours real not null,
                    reincarnation_count integer not null default 1,
                    is_host integer not null default 0,
                    primary key(session_id, player_id)
                );

                create table if not exists import_errors (
                    id integer primary key autoincrement,
                    batch_id integer references import_batches(id) on delete set null,
                    month text not null,
                    filename text not null,
                    content_hash text not null,
                    reason text not null,
                    model_name text not null,
                    raw_payload text,
                    created_at text not null
                );

                create index if not exists idx_players_name on players(name);
                create unique index if not exists idx_players_qq on players(qq) where qq is not null;
                """
            )

    @contextmanager
    def _connect(self) -> Iterator[sqlite3.Connection]:
        conn = sqlite3.connect(self.db_path, timeout=30)
        conn.row_factory = sqlite3.Row
        conn.execute("pragma foreign_keys = on")
        conn.execute("pragma busy_timeout = 30000")
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    @staticmethod
    def _now() -> str:
        return time.strftime("%Y-%m-%dT%H:%M:%S%z", time.localtime())

    @staticmethod
    def _content_hash(content: str) -> str:
        return hashlib.sha256(content.encode("utf-8")).hexdigest()

    @staticmethod
    def _clean_text(value: Any) -> str:
        if value is None:
            return ""
        return str(value).strip()

    @staticmethod
    def _optional_text(value: Any) -> str | None:
        text = SessionStatsService._clean_text(value)
        return text or None

    def _create_batch(self, conn: sqlite3.Connection, month: str, model_name: str, file_count: int) -> int:
        cursor = conn.execute(
            """
            insert into import_batches(month, file_count, success_count, failure_count, model_name, created_at)
            values (?, ?, 0, 0, ?, ?)
            """,
            (month, file_count, model_name, self._now()),
        )
        return int(cursor.lastrowid)

    def _finish_batch(self, conn: sqlite3.Connection, batch_id: int, success_count: int, failure_count: int) -> None:
        conn.execute(
            "update import_batches set success_count = ?, failure_count = ? where id = ?",
            (success_count, failure_count, batch_id),
        )

    def _is_duplicate(self, conn: sqlite3.Connection, month: str, filename: str, content_hash: str) -> bool:
        row = conn.execute(
            "select 1 from sessions where month = ? and source_filename = ? and content_hash = ?",
            (month, filename, content_hash),
        ).fetchone()
        return row is not None

    def _find_existing_session_by_content(self, conn: sqlite3.Connection, month: str, content_hash: str) -> sqlite3.Row | None:
        return conn.execute(
            """
            select id, source_filename
            from sessions
            where month = ? and content_hash = ?
            order by created_at asc, id asc
            limit 1
            """,
            (month, content_hash),
        ).fetchone()

    @staticmethod
    def _skip_item(filename: str, existing: sqlite3.Row) -> dict[str, Any]:
        return {
            "filename": filename,
            "success": True,
            "skipped": True,
            "reason": "SKIP：重复导入，同月已存在同内容团文件",
            "session_id": int(existing["id"]),
            "source_filename": existing["source_filename"],
        }

    def _record_error(
        self,
        conn: sqlite3.Connection,
        batch_id: int,
        month: str,
        filename: str,
        content_hash: str,
        reason: str,
        model_name: str,
        raw_payload: Any,
    ) -> None:
        conn.execute(
            """
            insert into import_errors(batch_id, month, filename, content_hash, reason, model_name, raw_payload, created_at)
            values (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (batch_id, month, filename, content_hash, reason, model_name, self._json_or_none(raw_payload), self._now()),
        )

    def _insert_session(
        self,
        conn: sqlite3.Connection,
        *,
        batch_id: int,
        month: str,
        file: SessionTextFile,
        content_hash: str,
        payload: dict[str, Any],
        model_name: str,
    ) -> int:
        duration_hours = float(payload["duration_hours"])
        kp = payload["kp"]
        kp_id = self._upsert_player(conn, kp["name"], kp.get("qq"), prefer_qq=True)
        cursor = conn.execute(
            """
            insert into sessions(
                batch_id, month, title, duration_hours, kp_player_id, source_filename,
                content_hash, model_name, confidence, raw_payload, created_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                batch_id,
                month,
                self._clean_text(payload["title"]),
                duration_hours,
                kp_id,
                file.filename,
                content_hash,
                model_name,
                payload.get("confidence"),
                json.dumps(payload, ensure_ascii=False),
                self._now(),
            ),
        )
        session_id = int(cursor.lastrowid)
        self._insert_participant(conn, session_id, kp_id, "kp", duration_hours, is_host=True)

        seen_player_ids = {kp_id}
        for player in payload["players"]:
            player_id = self._upsert_player(conn, player["name"], player.get("qq"), prefer_qq=True)
            if player_id in seen_player_ids:
                continue
            seen_player_ids.add(player_id)
            self._insert_participant(conn, session_id, player_id, "pl", duration_hours, is_host=False)
        return session_id

    def _upsert_player(self, conn: sqlite3.Connection, name: str, qq: Any, *, prefer_qq: bool) -> int:
        clean_name = self._clean_text(name)
        clean_qq = self._optional_text(qq)
        now = self._now()
        row = None
        if prefer_qq and clean_qq:
            row = conn.execute("select id, name, qq from players where qq = ?", (clean_qq,)).fetchone()
        if row is None:
            row = conn.execute(
                """
                select id, name, qq
                from players
                where name = ?
                order by case when qq is null then 0 else 1 end, id
                limit 1
                """,
                (clean_name,),
            ).fetchone()
        if prefer_qq and clean_qq and row is not None and row["qq"] != clean_qq:
            row = None
        if row:
            updates: list[str] = []
            params: list[Any] = []
            if clean_qq and not row["qq"]:
                updates.append("qq = ?")
                params.append(clean_qq)
            if prefer_qq and row["name"] != clean_name:
                updates.append("name = ?")
                params.append(clean_name)
            if updates:
                updates.append("updated_at = ?")
                params.append(now)
                params.append(row["id"])
                conn.execute(f"update players set {', '.join(updates)} where id = ?", params)
            return int(row["id"])
        cursor = conn.execute(
            "insert into players(name, qq, created_at, updated_at) values (?, ?, ?, ?)",
            (clean_name, clean_qq, now, now),
        )
        return int(cursor.lastrowid)

    def _insert_participant(
        self,
        conn: sqlite3.Connection,
        session_id: int,
        player_id: int,
        role: str,
        duration_hours: float,
        *,
        is_host: bool,
    ) -> None:
        conn.execute(
            """
            insert into session_participants(
                session_id, player_id, role, duration_hours, reincarnation_count, is_host
            )
            values (?, ?, ?, ?, 1, ?)
            """,
            (session_id, player_id, role, duration_hours, 1 if is_host else 0),
        )

    def _validation_error(self, payload: Any) -> str | None:
        if not isinstance(payload, dict):
            return "抽取结果不是对象"
        problems: list[str] = []
        if not self._clean_text(payload.get("title")):
            problems.append("缺少标题")
        duration = payload.get("duration_hours")
        if duration is None or duration == "":
            problems.append("缺少时长")
        else:
            try:
                duration_value = float(duration)
                if not math.isfinite(duration_value) or duration_value <= 0:
                    problems.append("时长必须为正数")
            except (TypeError, ValueError):
                problems.append("时长不是数字")
        confidence = payload.get("confidence")
        if confidence is not None and confidence != "":
            try:
                confidence_value = float(confidence)
                if not math.isfinite(confidence_value) or confidence_value < 0 or confidence_value > 1:
                    problems.append("confidence 必须在 0 到 1 之间")
            except (TypeError, ValueError):
                problems.append("confidence 不是数字")
        kp = payload.get("kp")
        if not isinstance(kp, dict) or not self._clean_text(kp.get("name")):
            problems.append("缺少KP")
        players = payload.get("players")
        if not isinstance(players, list):
            problems.append("缺少玩家列表")
        else:
            if not players:
                problems.append("缺少玩家列表")
            for index, player in enumerate(players, start=1):
                if not isinstance(player, dict) or not self._clean_text(player.get("name")):
                    problems.append(f"玩家列表第 {index} 项缺少姓名")
        return "；".join(problems) if problems else None

    @staticmethod
    def _json_or_none(value: Any) -> str | None:
        if value is None:
            return None
        try:
            return json.dumps(value, ensure_ascii=False)
        except TypeError:
            return str(value)

    @staticmethod
    def _player_row(row: sqlite3.Row) -> dict[str, Any]:
        return {
            "id": row["id"],
            "name": row["name"],
            "qq": row["qq"],
            "game_count": int(row["game_count"]),
            "game_hours": float(row["game_hours"]),
            "reincarnation_count": int(row["reincarnation_count"]),
            "host_count": int(row["host_count"]),
            "host_hours": float(row["host_hours"]),
        }
