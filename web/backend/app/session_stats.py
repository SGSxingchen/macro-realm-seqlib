from __future__ import annotations

import hashlib
import json
import math
import os
import re
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

QQ_CLEAN_RE = re.compile(r"(?<!\d)([1-9][0-9]{4,11})(?!\d)")
DURATION_EXPLICIT_RE = re.compile(
    r"(?:【\s*时长\s*】|(?:模组|本团|结团|总计|共计|总)?\s*时长)\s*[:：]?\s*(?:[^\r\n]{0,24}?总计\s*)?(\d+(?:\.\d+)?)\s*(?:小时|h|H|hr|hrs)?"
)
DURATION_FILENAME_RE = re.compile(r"(?<![.\d])(\d+(?:\.\d+)?)\s*(?:小时|h|H)(?![A-Za-z])")
DURATION_INLINE_RE = re.compile(r"(?<!\.ra)(?<!ra)(?<![.\d])(\d+(?:\.\d+)?)\s*(?:小时|h|H)\s*(?:结团|第一|第二|第三|第\d+|\)|）|$)")


def clean_person_identity(raw_name: Any, raw_qq: Any = None) -> tuple[str, str | None]:
    """修复结团文件常见姓名/QQ 瑕疵：
    1. 拆出混在名字里的 QQ 号，如『雨桦（595561352）』、『艾丝特1546003281』（纯数字卡名如『10086』不误当 QQ）；
    2. 剥离 KP 复合写法里的附加卡名，如『林雨霁（3909862519）+林若真（卡名）』-> name='林雨霁', qq='3909862519'（不误伤『苏早琴+1』类名字）；
    3. 剥离前缀『卡名：』『PC：』『PL：』『KP：』『主持：』与后缀『（卡名）』『(KP)』等噪音。
    """
    qq_str: str | None = None
    if raw_qq is not None:
        m_qq = QQ_CLEAN_RE.search(str(raw_qq).strip())
        if m_qq:
            qq_str = m_qq.group(1)

    text = str(raw_name or "").replace("\u3000", " ").strip()
    if not text:
        return "", qq_str

    # 去除常见字段头前缀
    text = re.sub(r"^(?:卡名|角色名|玩家|主持|主持人|KP|PL|PC)\s*[:：]\s*", "", text, flags=re.IGNORECASE).strip()

    # 若还没拿到 QQ，尝试从名字字符串提取 5~12 位 QQ 号；若整串本身就是纯数字卡名（如『10086』），不误判为 QQ
    if not qq_str:
        m_in_name = QQ_CLEAN_RE.search(text)
        if m_in_name and m_in_name.group(0) != text:
            qq_str = m_in_name.group(1)

    # 仅在确有『（卡名）/（角色）/（PC）』复合身份标注时按 + 切分，避免误切『苏早琴+1』
    if ("+" in text or "＋" in text) and re.search(r"[（(\[【]\s*(?:卡名|角色名|角色|PC)\s*[)）\]】]", text, flags=re.IGNORECASE):
        parts = [p.strip() for p in re.split(r"[+＋]", text) if p.strip()]
        if len(parts) >= 2:
            non_card = [p for p in parts if not re.search(r"[（(\[【]\s*(?:卡名|角色名|角色|PC)\s*[)）\]】]", p, flags=re.IGNORECASE)]
            text = non_card[0] if non_card else parts[0]

    # 剥离括号内的纯 QQ 号或角色标注，如（3909862519）、（卡名）、（KP）
    text = re.sub(r"[（(\[【]\s*(?:QQ\s*[:：]?\s*)?[1-9][0-9]{4,11}\s*[)）\]】]", "", text, flags=re.IGNORECASE)
    text = re.sub(r"[（(\[【]\s*(?:卡名|角色名|角色|KP|PL|PC|认证KP|非认证KP)\s*[)）\]】]", "", text, flags=re.IGNORECASE)
    # 若名字末尾紧贴该 QQ 号（如『艾丝特1546003281』且前面已有非数字字符），剥离尾部 QQ
    if qq_str and text.endswith(qq_str) and len(text) > len(qq_str):
        prefix = text[: -len(qq_str)].rstrip(" -_·:：")
        if prefix:
            text = prefix

    # 清理残余空括号与首尾空白
    text = re.sub(r"[（(\[【]\s*[)）\]】]", "", text).strip(" -_—·\t")
    return text, qq_str


def infer_duration_hours_fallback(filename: str, content: str) -> float | None:
    """当 LLM 未能提取出时长或填了 null/0 时，从文件名与正文中安全回退提取小时数。"""
    stem = Path(filename or "").stem
    # 1. 优先查正文中明确的【时长】9h / 模组时长：32小时 / 时长：11.10~11.15 总计21H
    for m in DURATION_EXPLICIT_RE.finditer(content or ""):
        try:
            val = float(m.group(1))
            if 0.5 <= val <= 300:
                return val
        except ValueError:
            continue

    # 2. 查文件名中的明确时长标记，如『地城追凶8h结团.txt』『隔离区h(31小时)第一车』『游乐场-35.5h』
    for m in DURATION_FILENAME_RE.finditer(stem):
        try:
            val = float(m.group(1))
            if 0.5 <= val <= 300:
                return val
        except ValueError:
            continue

    # 3. 查正文首行标题里的 XXh / XX小时（预先剥除 .ra 检定表达式，防 .ra3.5h 尾部匹配）
    first_lines = "\n".join((content or "").splitlines()[:5])
    sanitized_head = re.sub(r"\.?ra\s*\d+(?:\.\d+)?\s*(?:h|H|小时)?(?:\s*\*\s*\d+)?", " ", first_lines, flags=re.IGNORECASE)
    for m in DURATION_INLINE_RE.finditer(sanitized_head):
        try:
            val = float(m.group(1))
            if 0.5 <= val <= 300:
                return val
        except ValueError:
            continue
    return None


def strip_markdown_code_fence(text: str) -> str:
    cleaned = (text or "").strip()
    # 移除可能混入的 <think>...</think> 块
    cleaned = re.sub(r"<think>[\s\S]*?</think>", "", cleaned, flags=re.IGNORECASE).strip()
    if cleaned.startswith("```"):
        cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
        cleaned = re.sub(r"\s*```$", "", cleaned)
    return cleaned.strip()


@dataclass(frozen=True)
class SessionTextFile:
    filename: str
    content: str
    source_file_id: str | None = None
    uploader_qq: str | None = None
    uploader_name: str | None = None


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
            "你是“宏观界域”TRPG跑团结团报告结构化提取器。请从下面的结团文本与文件名中提取统计信息，严格只返回 JSON 对象。\n"
            "字段要求：\n"
            "1. title：模组/团名称字符串（去掉末尾多余的『结团报告』『.txt』等文件后缀，保留车次说明即可）。\n"
            "2. duration_hours：本团数字小时数（如 9、3.5、35.5）。注意：时长可能写在正文【时长】/模组时长/总计XXH中，也可能只写在文件名或首行标题（如『刮刮乐8h结团』『31小时』『35.5h』）里，请务必综合提取；不要把『.ra35.5*2』检定当成额外时长。\n"
            "3. kp：主持人对象 {name, qq}。常位于文首【KP】或文末『KP：/主持：/kp结算』处。若写成『林雨霁（3909862519）+林若真（卡名）』，name 填主持本人名『林雨霁』，qq 填『3909862519』；没有 QQ 则填 null。\n"
            "4. players：参团玩家数组 [{name, qq}]。每个拥有独立评分/扮演/影响/评语/个人额外结算段落的角色算一名玩家（注意排除 KP 本人，且绝不要输出 name 为空或 null 的占位元素）。若玩家名带括号 QQ 如『梅普露（2577535590）』，拆分为 name='梅普露', qq='2577535590'；若只有『卡名：芬』无 QQ，则 name='芬', qq=null。\n"
            "5. confidence：0 到 1 之间的置信度数字。\n"
            "6. warnings：字符串数组（如有乱码或不确定项可简短说明，正常填 []）。"
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
        text = strip_markdown_code_fence(text_reader(data))
        if not text:
            raise RuntimeError("OpenAI 返回为空")
        try:
            result = json.loads(text)
        except json.JSONDecodeError as exc:
            # 容错提取最外层 {...} JSON 对象
            start = text.find("{")
            end = text.rfind("}")
            if start != -1 and end > start:
                try:
                    result = json.loads(text[start : end + 1])
                except json.JSONDecodeError:
                    raise RuntimeError("OpenAI 返回不是合法 JSON") from exc
            else:
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
                existing_by_fid = self._find_existing_session_by_source_file_id(conn, month, file.source_file_id)
                if existing_by_fid and existing_by_fid["content_hash"] == content_hash:
                    item = self._skip_item(file.filename, existing_by_fid)
                    items.append(item)
                    skip_count += 1
                    if progress_callback:
                        progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)
                    continue
                existing = self._find_existing_session_by_content(conn, month, content_hash)
                if existing:
                    if file.source_file_id and not existing["source_file_id"]:
                        conn.execute("update sessions set source_file_id = ? where id = ?", (file.source_file_id, existing["id"]))
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
                extracted = self._normalize_extracted_payload(extracted, filename=file.filename, content=file.content)
                reason = self._validation_error(extracted)
                if reason:
                    raise ValueError(reason)
                with self._connect() as conn:
                    existing = self._find_existing_session_by_content(conn, month, content_hash)
                    if existing:
                        if file.source_file_id and not existing["source_file_id"]:
                            conn.execute("update sessions set source_file_id = ? where id = ?", (file.source_file_id, existing["id"]))
                        item = self._skip_item(file.filename, existing)
                        items.append(item)
                        skip_count += 1
                        if progress_callback:
                            progress_callback(item, success_count + failure_count + skip_count, success_count, failure_count)
                        continue
                    # 如果同一群文件 file_id 更新了正文内容，原子替换旧版本而不重复记两场团
                    old_same_fid = self._find_existing_session_by_source_file_id(conn, month, file.source_file_id)
                    if old_same_fid:
                        conn.execute("delete from sessions where id = ?", (old_same_fid["id"],))
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
                existing_by_fid = self._find_existing_session_by_source_file_id(conn, month, file.source_file_id)
                if existing_by_fid and existing_by_fid["content_hash"] == content_hash:
                    return self._skip_item(file.filename, existing_by_fid)
                existing = self._find_existing_session_by_content(conn, month, content_hash)
                if existing:
                    if file.source_file_id and not existing["source_file_id"]:
                        conn.execute("update sessions set source_file_id = ? where id = ?", (file.source_file_id, existing["id"]))
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
                extracted = self._normalize_extracted_payload(extracted, filename=file.filename, content=file.content)
                reason = self._validation_error(extracted)
                if reason:
                    raise ValueError(reason)
                with self._connect() as conn:
                    existing = self._find_existing_session_by_content(conn, month, content_hash)
                    if existing:
                        if file.source_file_id and not existing["source_file_id"]:
                            conn.execute("update sessions set source_file_id = ? where id = ?", (file.source_file_id, existing["id"]))
                        return self._skip_item(file.filename, existing)
                    old_same_fid = self._find_existing_session_by_source_file_id(conn, month, file.source_file_id)
                    if old_same_fid:
                        conn.execute("delete from sessions where id = ?", (old_same_fid["id"],))
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

    def has_imported_report(self, *, month: str, source_file_id: str | None, content_hash: str) -> bool:
        with self._connect() as conn:
            if source_file_id:
                row = self._find_existing_session_by_source_file_id(conn, month, source_file_id)
                if row and row["content_hash"] == content_hash:
                    return True
            return self._find_existing_session_by_content(conn, month, content_hash) is not None

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

    def update_session(
        self,
        session_id: int,
        *,
        title: str | None = None,
        duration_hours: float | None = None,
        kp: dict[str, Any] | None = None,
        players: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any] | None:
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

        with self._connect() as conn:
            current = conn.execute(
                "select id, duration_hours, kp_player_id from sessions where id = ?",
                (session_id,),
            ).fetchone()
            if current is None:
                return None

            effective_duration = float(duration_hours) if duration_hours is not None else float(current["duration_hours"])
            kp_player_id = int(current["kp_player_id"])

            if kp is not None:
                kp_name, kp_qq = clean_person_identity(kp.get("name"), kp.get("qq"))
                if not kp_name:
                    raise ValueError("KP 姓名不能为空")
                kp_player_id = self._upsert_player(conn, kp_name, kp_qq, prefer_qq=True)
                updates.append("kp_player_id = ?")
                params.append(kp_player_id)

            if updates:
                conn.execute(f"update sessions set {', '.join(updates)} where id = ?", [*params, session_id])

            if kp is not None or players is not None:
                # 重建本团参与者关系（当仅改 kp 或仅改 players 时保留另一方既有成员）
                if players is None:
                    existing_pls = conn.execute(
                        """
                        select p.name, p.qq
                        from session_participants sp
                        join players p on p.id = sp.player_id
                        where sp.session_id = ? and sp.is_host = 0
                        order by sp.rowid asc
                        """,
                        (session_id,),
                    ).fetchall()
                    pl_specs = [{"name": r["name"], "qq": r["qq"]} for r in existing_pls]
                else:
                    pl_specs = []
                    for idx, item in enumerate(players, start=1):
                        if not isinstance(item, dict):
                            raise ValueError(f"第 {idx} 名玩家格式无效")
                        p_name, p_qq = clean_person_identity(item.get("name"), item.get("qq"))
                        if not p_name:
                            continue
                        pl_specs.append({"name": p_name, "qq": p_qq})
                    if not pl_specs:
                        raise ValueError("至少需要保留 1 名参团玩家（PL）")

                conn.execute("delete from session_participants where session_id = ?", (session_id,))
                self._insert_participant(conn, session_id, kp_player_id, "kp", effective_duration, is_host=True)
                seen_ids = {kp_player_id}
                for spec in pl_specs:
                    pid = self._upsert_player(conn, spec["name"], spec["qq"], prefer_qq=True)
                    if pid in seen_ids:
                        continue
                    seen_ids.add(pid)
                    self._insert_participant(conn, session_id, pid, "pl", effective_duration, is_host=False)
            elif duration_hours is not None:
                conn.execute(
                    "update session_participants set duration_hours = ? where session_id = ?",
                    (effective_duration, session_id),
                )

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
            cols = {row["name"] for row in conn.execute("pragma table_info(sessions)").fetchall()}
            if "source_file_id" not in cols:
                conn.execute("alter table sessions add column source_file_id text")
            conn.execute(
                "create index if not exists idx_sessions_month_file_id on sessions(month, source_file_id) where source_file_id is not null"
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
            select id, source_filename, source_file_id, content_hash
            from sessions
            where month = ? and content_hash = ?
            order by created_at asc, id asc
            limit 1
            """,
            (month, content_hash),
        ).fetchone()

    def _find_existing_session_by_source_file_id(
        self, conn: sqlite3.Connection, month: str, source_file_id: str | None
    ) -> sqlite3.Row | None:
        if not source_file_id:
            return None
        return conn.execute(
            """
            select id, source_filename, source_file_id, content_hash
            from sessions
            where month = ? and source_file_id = ?
            order by created_at desc, id desc
            limit 1
            """,
            (month, source_file_id),
        ).fetchone()

    def _normalize_extracted_payload(self, payload: Any, *, filename: str, content: str) -> Any:
        """对 LLM 抽取结果执行确定性瑕疵修复：
        - 标题缺失时回退为文件名去扩展名；
        - 时长缺失/非正数时从正文与文件名正则兜底提取；
        - 清洗 KP 与 PL 的混写 QQ / 附加卡名 / 字段前缀；
        - 若全量玩家列表混入个别全空占位对象（如 {name: null, qq: null}）且存在有效玩家，自动剔除空噪音项；
        - 自动剔除被重复塞进 players 里的 KP 本人（同名或同 QQ）。
        """
        if not isinstance(payload, dict):
            return payload
        out = dict(payload)

        # 1. 标题兜底：仅当原字段已经是空/None但键存在时，若用户原本明确传了空标题由校验拦截，只有 None 时用文件名兜底
        if out.get("title") is None and filename:
            stem = Path(filename).stem.strip()
            if stem:
                out["title"] = stem

        # 2. 时长兜底：当 LLM 返回 None / 0 / 空串时，尝试从原文或文件名中提取明确小时数
        dur = out.get("duration_hours")
        need_fallback = False
        if dur is None or dur == "":
            need_fallback = True
        else:
            try:
                if float(dur) <= 0:
                    need_fallback = True
            except (TypeError, ValueError):
                need_fallback = True
        if need_fallback:
            inferred = infer_duration_hours_fallback(filename, content)
            if inferred is not None:
                out["duration_hours"] = inferred
                warnings = list(out.get("warnings") or []) if isinstance(out.get("warnings"), list) else []
                warnings.append(f"时长由文件名/正文规则兜底提取：{inferred}h")
                out["warnings"] = warnings

        # 3. KP 清洗
        kp = out.get("kp")
        kp_clean_name = ""
        kp_clean_qq: str | None = None
        if isinstance(kp, dict):
            kp_clean_name, kp_clean_qq = clean_person_identity(kp.get("name"), kp.get("qq"))
            if not kp_clean_name and kp_clean_qq:
                kp_clean_name = f"KP_{kp_clean_qq[-4:]}"
            out["kp"] = {"name": kp_clean_name, "qq": kp_clean_qq}

        # 4. Players 清洗：过滤纯 null 噪音占位，但保留显式残缺字典（如 {}）在无其他信息时的校验语义；
        #    当至少有 1 名有效具名玩家时，丢弃 LLM 因乱码产生的 {name: null, qq: null} 垃圾项。
        players = out.get("players")
        if isinstance(players, list):
            cleaned_valid: list[dict[str, Any]] = []
            cleaned_all: list[dict[str, Any]] = []
            for p in players:
                if not isinstance(p, dict):
                    cleaned_all.append(p)
                    continue
                # 区分完全无 name 键的非法结构和 LLM 输出的 {"name": null, "qq": null}
                has_name_key = "name" in p
                p_name, p_qq = clean_person_identity(p.get("name"), p.get("qq"))
                if not p_name and p_qq:
                    p_name = f"玩家_{p_qq[-4:]}"
                if p_name:
                    # 排除被 LLM 误重复列入 PL 的 KP 本人
                    if kp_clean_qq and p_qq and kp_clean_qq == p_qq:
                        continue
                    if not kp_clean_qq and not p_qq and kp_clean_name and p_name == kp_clean_name:
                        continue
                    entry = {"name": p_name, "qq": p_qq}
                    cleaned_valid.append(entry)
                    cleaned_all.append(entry)
                else:
                    if has_name_key and p.get("name") is None and not p_qq:
                        # LLM 因局部乱码吐出的 {"name": null, "qq": null}
                        continue
                    cleaned_all.append(p)
            if cleaned_valid and len(cleaned_valid) == len(cleaned_all):
                out["players"] = cleaned_valid
            elif cleaned_valid and all(isinstance(x, dict) and "name" in x for x in players):
                out["players"] = cleaned_valid
            else:
                out["players"] = cleaned_all

        return out

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
                source_file_id, content_hash, model_name, confidence, raw_payload, created_at
            )
            values (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                batch_id,
                month,
                self._clean_text(payload["title"]),
                duration_hours,
                kp_id,
                file.filename,
                file.source_file_id,
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

    def _merge_orphan_player_into(self, conn: sqlite3.Connection, source_id: int, target_id: int) -> None:
        if source_id == target_id:
            return
        conn.execute("update sessions set kp_player_id = ? where kp_player_id = ?", (target_id, source_id))
        # 对同一 session 若 target 已在参与者表中，删掉重复的 source 行，其余迁移为 target_id
        conn.execute(
            """
            delete from session_participants
            where player_id = ?
              and session_id in (select session_id from session_participants where player_id = ?)
            """,
            (source_id, target_id),
        )
        conn.execute("update session_participants set player_id = ? where player_id = ?", (target_id, source_id))
        conn.execute("delete from players where id = ?", (source_id,))

    def _upsert_player(self, conn: sqlite3.Connection, name: str, qq: Any, *, prefer_qq: bool) -> int:
        clean_name = self._clean_text(name)
        clean_qq = self._optional_text(qq)
        now = self._now()

        by_qq = conn.execute("select id, name, qq from players where qq = ?", (clean_qq,)).fetchone() if (prefer_qq and clean_qq) else None
        by_name = conn.execute(
            """
            select id, name, qq
            from players
            where name = ?
            order by case when qq is null then 0 else 1 end, id
            limit 1
            """,
            (clean_name,),
        ).fetchone()

        if by_qq is not None:
            target_id = int(by_qq["id"])
            # 若此前存在同名且 qq IS NULL 的孤儿记录，自动归并历史参团记录到该 QQ 账号下
            if by_name is not None and int(by_name["id"]) != target_id and not by_name["qq"]:
                self._merge_orphan_player_into(conn, int(by_name["id"]), target_id)
            if prefer_qq and by_qq["name"] != clean_name:
                conn.execute("update players set name = ?, updated_at = ? where id = ?", (clean_name, now, target_id))
            return target_id

        row = by_name
        # 仅当同名记录已经绑定了另一个不同的非空 QQ 时（同名不同人），才不复用而新建独立玩家行
        if prefer_qq and clean_qq and row is not None and row["qq"] is not None and row["qq"] != clean_qq:
            row = None

        if row is not None:
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
