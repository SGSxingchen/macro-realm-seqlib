from __future__ import annotations

import base64
import difflib
import hashlib
import hmac
import json
import locale
import os
import re
import secrets
import shutil
import sys
import subprocess
import threading
import time
import uuid
from pathlib import Path
from typing import Annotated, Iterable, Literal

from fastapi import Cookie, Depends, FastAPI, File, Form, HTTPException, Query, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel, Field

from .search import SearchIndex
from .session_stats import SessionStatsService, SessionTextFile

REPO_ROOT = Path(__file__).resolve().parents[3]
ALLOWED_ROOTS = ("序列库", "荣誉室")
PUBLIC_ROOTS = ("序列库",)
TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "gbk", "gb2312", "big5")
PROCESS_ENCODINGS = tuple(dict.fromkeys(("utf-8", locale.getpreferredencoding(False), sys.getfilesystemencoding(), "gbk", "gb2312")))
SESSION_COOKIE = "seqlib_admin"
NORMALIZATION_REVIEW_DIR = REPO_ROOT / "web" / "normalization_reviews"
NORMALIZATION_SIGNATURES_REQUIRED = 1
SESSION_STATS_DB = REPO_ROOT / "web" / "data" / "session_stats.sqlite"
SESSION_STATS_SERVICE: SessionStatsService | None = None
SESSION_STATS_EXTRACTOR = None
SESSION_IMPORT_JOBS: dict[str, dict] = {}
SESSION_IMPORT_JOBS_LOCK = threading.Lock()

app = FastAPI(title="宏观界域强化序列库 Web API", version="0.2.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.getenv("WEB_CORS_ORIGINS", "http://localhost:5173").split(","),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 全局搜索索引：仅索引公开 root，进程内 mtime 缓存
SEARCH_INDEX = SearchIndex(REPO_ROOT, PUBLIC_ROOTS)


@app.on_event("startup")
def _warm_index() -> None:
    try:
        SEARCH_INDEX.refresh()
    except Exception:
        # 启动期失败不阻塞服务，首个查询会再尝试构建
        pass


def read_text(path: Path) -> tuple[str, str]:
    for enc in TEXT_ENCODINGS:
        try:
            return path.read_text(encoding=enc), enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return path.read_text(encoding="utf-8", errors="replace"), "utf-8-replace"


def write_text_utf8(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8", newline="\n")


def now_iso() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%S%z", time.localtime())


def sort_key(name: str):
    match = re.match(r"(\d+)", name)
    if match:
        return (0, int(match.group(1)), name)
    return (1, 0, name)


def strip_number_prefix(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r"^\d+】", "", stem).strip() or stem


def title_for(path: Path) -> str:
    try:
        text, _ = read_text(path)
        first = next((line.strip().lstrip("\ufeff") for line in text.splitlines() if line.strip()), "")
        return first or strip_number_prefix(path.name)
    except Exception:
        return strip_number_prefix(path.name)


def rel_posix(path: Path) -> str:
    return path.relative_to(REPO_ROOT).as_posix()


def safe_resource_path(rel_path: str, *, must_exist: bool = False, allowed_roots: tuple[str, ...] = ALLOWED_ROOTS) -> Path:
    if "\x00" in rel_path:
        raise HTTPException(400, "路径包含非法字符")
    rel = Path(rel_path)
    if rel.is_absolute() or ".." in rel.parts:
        raise HTTPException(400, "路径越界")
    if not rel.parts or rel.parts[0] not in allowed_roots:
        root_hint = " 或 ".join(f"{root}/" for root in allowed_roots)
        raise HTTPException(400, f"路径必须位于 {root_hint} 内")
    if rel.suffix.lower() != ".txt":
        raise HTTPException(400, "只允许 .txt 文件")
    full = (REPO_ROOT / rel).resolve()
    roots = [(REPO_ROOT / root).resolve() for root in allowed_roots]
    if not any(full == root or root in full.parents for root in roots):
        raise HTTPException(400, "路径越界")
    if must_exist and not full.is_file():
        raise HTTPException(404, "资源不存在")
    return full



def public_resource_path(rel_path: str, *, must_exist: bool = False) -> Path:
    return safe_resource_path(rel_path, must_exist=must_exist, allowed_roots=PUBLIC_ROOTS)

def resource_entry(path: Path) -> dict:
    stat = path.stat()
    rel = Path(rel_posix(path))
    category = "/".join(rel.parts[1:-1])
    return {
        "path": rel.as_posix(),
        "filename": path.name,
        "title": title_for(path),
        "root": rel.parts[0],
        "category": category,
        "mtime": stat.st_mtime,
        "size": stat.st_size,
    }


def iter_txt_files(roots: tuple[str, ...] = PUBLIC_ROOTS) -> Iterable[Path]:
    for root in roots:
        base = REPO_ROOT / root
        if base.exists():
            yield from sorted(base.rglob("*.txt"), key=lambda p: rel_posix(p))


def admin_password() -> str | None:
    return os.getenv("ADMIN_PASSWORD")


def secret_key() -> bytes:
    seed = os.getenv("ADMIN_SECRET") or admin_password() or "dev-unconfigured-secret"
    return hashlib.sha256(seed.encode("utf-8")).digest()


def sign_payload(payload: dict) -> str:
    raw = json.dumps(payload, separators=(",", ":"), ensure_ascii=False).encode("utf-8")
    body = base64.urlsafe_b64encode(raw).decode().rstrip("=")
    sig = hmac.new(secret_key(), body.encode(), hashlib.sha256).hexdigest()
    return f"{body}.{sig}"


def verify_token(token: str | None) -> bool:
    if not token:
        return False
    try:
        body, sig = token.rsplit(".", 1)
        good = hmac.new(secret_key(), body.encode(), hashlib.sha256).hexdigest()
        if not hmac.compare_digest(sig, good):
            return False
        raw = base64.urlsafe_b64decode(body + "=" * (-len(body) % 4))
        payload = json.loads(raw)
        return payload.get("exp", 0) >= time.time() and payload.get("role") == "admin"
    except Exception:
        return False


def require_admin(token: str | None = Cookie(default=None, alias=SESSION_COOKIE)) -> None:
    if not admin_password():
        raise HTTPException(503, "未配置 ADMIN_PASSWORD，后台写操作已禁用")
    if not verify_token(token):
        raise HTTPException(401, "需要管理员登录")


def session_stats_service() -> SessionStatsService:
    global SESSION_STATS_SERVICE
    if SESSION_STATS_SERVICE is None:
        SESSION_STATS_SERVICE = SessionStatsService(SESSION_STATS_DB)
    return SESSION_STATS_SERVICE


def validate_stats_month(month: str) -> str:
    value = month.strip()
    if not re.fullmatch(r"\d{4}-\d{2}", value):
        raise HTTPException(400, "月份格式必须为 YYYY-MM")
    mm = int(value[-2:])
    if mm < 1 or mm > 12:
        raise HTTPException(400, "月份必须在 01 到 12 之间")
    return value


def validate_import_concurrency(concurrency: int) -> int:
    if concurrency < 1:
        return 1
    if concurrency > 6:
        return 6
    return concurrency


def decode_uploaded_txt(data: bytes) -> tuple[str, str]:
    for enc in TEXT_ENCODINGS:
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def session_stats_extractor():
    return SESSION_STATS_EXTRACTOR


def import_job_snapshot(job_id: str) -> dict:
    with SESSION_IMPORT_JOBS_LOCK:
        job = SESSION_IMPORT_JOBS.get(job_id)
        if not job:
            raise HTTPException(404, "导入任务不存在")
        return dict(job)


def update_import_job(job_id: str, **updates) -> None:
    with SESSION_IMPORT_JOBS_LOCK:
        job = SESSION_IMPORT_JOBS.get(job_id)
        if not job:
            return
        job.update(updates)
        job["updated_at"] = now_iso()


def run_session_import_job(job_id: str, month: str, files: list[SessionTextFile], model_name: str, concurrency: int) -> None:
    update_import_job(job_id, status="running", current_filename="")

    def on_progress(item: dict, processed_count: int, success_count: int, failure_count: int) -> None:
        current = str(item.get("filename") or "")
        with SESSION_IMPORT_JOBS_LOCK:
            job = SESSION_IMPORT_JOBS.get(job_id)
            if not job:
                return
            job["processed_count"] = processed_count
            job["success_count"] = success_count
            job["failure_count"] = failure_count
            if item.get("skipped") is True:
                job["skip_count"] = int(job.get("skip_count") or 0) + 1
            job["current_filename"] = current
            job.setdefault("items", []).append(item)
            job["updated_at"] = now_iso()

    try:
        result = session_stats_service().import_text_files(
            month,
            files,
            extractor=session_stats_extractor(),
            model_name=model_name,
            concurrency=concurrency,
            progress_callback=on_progress,
        )
    except Exception as exc:
        update_import_job(job_id, status="failed", error=str(exc) or exc.__class__.__name__, current_filename="")
        return

    update_import_job(
        job_id,
        status="completed",
        processed_count=len(files),
        success_count=result["success_count"],
        failure_count=result["failure_count"],
        skip_count=result.get("skip_count", 0),
        items=result["items"],
        current_filename="",
        error="",
    )


def run_cmd(args: list[str], timeout: int = 120) -> dict:
    started = time.time()
    try:
        p = subprocess.run(
            args,
            cwd=REPO_ROOT,
            capture_output=True,
            timeout=timeout,
        )
        return {
            "cmd": args,
            "returncode": p.returncode,
            "stdout": decode_process_output(p.stdout),
            "stderr": decode_process_output(p.stderr),
            "seconds": round(time.time() - started, 3),
        }
    except subprocess.TimeoutExpired as e:
        return {"cmd": args, "returncode": 124, "stdout": e.stdout or "", "stderr": f"超时: {e}", "seconds": timeout}


def git(*args: str) -> dict:
    return run_cmd(["git", "-c", "core.quotepath=false", *args])


def decode_process_output(data: bytes | str | None) -> str:
    if data is None:
        return ""
    if isinstance(data, str):
        return data
    for enc in PROCESS_ENCODINGS:
        try:
            return data.decode(enc)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def git_bytes(*args: str, timeout: int = 120) -> dict:
    started = time.time()
    try:
        p = subprocess.run(
            ["git", "-c", "core.quotepath=false", *args],
            cwd=REPO_ROOT,
            capture_output=True,
            timeout=timeout,
        )
        return {
            "cmd": ["git", "-c", "core.quotepath=false", *args],
            "returncode": p.returncode,
            "stdout": p.stdout,
            "stderr": decode_process_output(p.stderr),
            "seconds": round(time.time() - started, 3),
        }
    except subprocess.TimeoutExpired as e:
        return {"cmd": ["git", "-c", "core.quotepath=false", *args], "returncode": 124, "stdout": b"", "stderr": f"超时: {e}", "seconds": timeout}


class LoginRequest(BaseModel):
    password: str


class SaveRequest(BaseModel):
    path: str
    content: str = ""
    overwrite: bool = False


class MoveRequest(BaseModel):
    old_path: str
    new_path: str
    overwrite: bool = False


class DeleteRequest(BaseModel):
    path: str


class NormalizationReviewCreateRequest(BaseModel):
    resource_path: str
    normalized_content: str
    original_content: str | None = None
    title: str | None = None
    note: str | None = None


class NormalizationReviewSignRequest(BaseModel):
    signer: str = Field(min_length=1, max_length=80)
    note: str | None = Field(default=None, max_length=500)


class PackageFileRequest(BaseModel):
    path: str
    content: str = ""
    overwrite: bool = False


class MoveToHonorRequest(BaseModel):
    path: str
    category: str | None = None
    title: str | None = None
    filename: str | None = None
    target_path: str | None = None
    overwrite: bool = False


class SessionStatsSessionUpdateRequest(BaseModel):
    title: str | None = Field(default=None, max_length=200)
    duration_hours: float | None = Field(default=None, gt=0)


class PublishRequest(BaseModel):
    version: str = Field(pattern=r"^v?\d+(?:\.\d+)*(?:[-\w.]*)?$")
    message: str = "更新序列库 latest"
    branch: str = Field(default="main", pattern=r"^[A-Za-z0-9._/-]+$")
    push: bool = True


def wiki_credentials_configured() -> bool:
    return bool((os.getenv("WIKI_USER") or os.getenv("FANDOM_USER")) and (os.getenv("WIKI_PASSWORD") or os.getenv("FANDOM_PASSWORD")))


def safe_review_id(review_id: str) -> str:
    if not re.fullmatch(r"[A-Za-z0-9._-]{8,80}", review_id):
        raise HTTPException(400, "审核任务 ID 非法")
    return review_id


def review_file_path(review_id: str) -> Path:
    safe_id = safe_review_id(review_id)
    full = (NORMALIZATION_REVIEW_DIR / f"{safe_id}.json").resolve()
    if NORMALIZATION_REVIEW_DIR.resolve() not in full.parents:
        raise HTTPException(400, "审核任务路径越界")
    return full


def load_review(review_id: str) -> dict:
    full = review_file_path(review_id)
    if not full.is_file():
        raise HTTPException(404, "审核任务不存在")
    try:
        return json.loads(full.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        raise HTTPException(500, "审核任务文件损坏")


def save_review(review: dict) -> None:
    review_id = safe_review_id(str(review.get("id", "")))
    full = review_file_path(review_id)
    full.parent.mkdir(parents=True, exist_ok=True)
    full.write_text(json.dumps(review, ensure_ascii=False, indent=2) + "\n", encoding="utf-8", newline="\n")


def review_summary(review: dict) -> dict:
    signatures = review.get("signatures") or []
    required = int(review.get("required_signatures") or NORMALIZATION_SIGNATURES_REQUIRED)
    status = "approved" if len(signatures) >= required else "pending"
    review["status"] = status
    return {
        "id": review.get("id"),
        "resource_path": review.get("resource_path"),
        "title": review.get("title"),
        "status": status,
        "signature_count": len(signatures),
        "required_signatures": required,
        "updated_at": review.get("updated_at"),
        "created_at": review.get("created_at"),
    }


@app.get("/api/health")
def health():
    return {"ok": True, "repo": str(REPO_ROOT), "admin_configured": bool(admin_password()), "wiki_configured": wiki_credentials_configured()}


@app.get("/api/session-stats/overview")
def session_stats_overview(month: str):
    month = validate_stats_month(month)
    return session_stats_service().get_overview(month)


@app.get("/api/session-stats/players")
def session_stats_players(month: str, sort: Literal["games", "hours", "hosts", "name"] = "hours"):
    month = validate_stats_month(month)
    items = session_stats_service().list_players(month)
    if sort == "games":
        items.sort(key=lambda x: (-x["game_count"], -x["game_hours"], x["name"]))
    elif sort == "hosts":
        items.sort(key=lambda x: (-x["host_count"], -x["host_hours"], x["name"]))
    elif sort == "name":
        items.sort(key=lambda x: x["name"])
    else:
        items.sort(key=lambda x: (-x["game_hours"], -x["game_count"], x["name"]))
    return {"items": items, "count": len(items), "month": month}


@app.get("/api/session-stats/sessions")
def session_stats_sessions(month: str):
    month = validate_stats_month(month)
    items = session_stats_service().list_sessions(month)
    return {"items": items, "count": len(items), "month": month}


@app.get("/api/session-stats/errors")
def session_stats_errors(month: str, _admin: None = Depends(require_admin)):
    month = validate_stats_month(month)
    items = session_stats_service().list_errors(month)
    return {"items": items, "count": len(items), "month": month}


@app.post("/api/session-stats/import")
async def session_stats_import(
    month: str = Form(...),
    files: list[UploadFile] = File(...),
    _admin: None = Depends(require_admin),
):
    month = validate_stats_month(month)
    text_files: list[SessionTextFile] = []
    for file in files:
        filename = file.filename or "未命名.txt"
        if not filename.lower().endswith(".txt"):
            raise HTTPException(400, f"只支持上传 .txt：{filename}")
        data = await file.read()
        if len(data) > 2_000_000:
            raise HTTPException(413, f"TXT 过大：{filename}")
        content, _encoding = decode_uploaded_txt(data)
        text_files.append(SessionTextFile(filename=filename, content=content))
    model_name = os.getenv("OPENAI_MODEL", "deepseek-v4-pro")
    return session_stats_service().import_text_files(month, text_files, extractor=session_stats_extractor(), model_name=model_name)


@app.post("/api/session-stats/import-jobs")
async def session_stats_import_job_create(
    month: str = Form(...),
    concurrency: int = Form(1),
    files: list[UploadFile] = File(...),
    _admin: None = Depends(require_admin),
):
    month = validate_stats_month(month)
    concurrency = validate_import_concurrency(concurrency)
    text_files: list[SessionTextFile] = []
    for file in files:
        filename = file.filename or "未命名.txt"
        if not filename.lower().endswith(".txt"):
            raise HTTPException(400, f"只支持上传 .txt：{filename}")
        data = await file.read()
        if len(data) > 2_000_000:
            raise HTTPException(413, f"TXT 过大：{filename}")
        content, _encoding = decode_uploaded_txt(data)
        text_files.append(SessionTextFile(filename=filename, content=content))
    if not text_files:
        raise HTTPException(400, "至少需要上传一个 TXT 文件")

    job_id = uuid.uuid4().hex
    now = now_iso()
    job = {
        "job_id": job_id,
        "status": "running",
        "month": month,
        "total_count": len(text_files),
        "processed_count": 0,
        "success_count": 0,
        "failure_count": 0,
        "skip_count": 0,
        "concurrency": concurrency,
        "current_filename": "",
        "items": [],
        "error": "",
        "created_at": now,
        "updated_at": now,
    }
    with SESSION_IMPORT_JOBS_LOCK:
        SESSION_IMPORT_JOBS[job_id] = job

    model_name = os.getenv("OPENAI_MODEL", "deepseek-v4-pro")
    thread = threading.Thread(
        target=run_session_import_job,
        args=(job_id, month, text_files, model_name, concurrency),
        daemon=True,
    )
    thread.start()
    return import_job_snapshot(job_id)


@app.get("/api/session-stats/import-jobs/{job_id}")
def session_stats_import_job_get(job_id: str, _admin: None = Depends(require_admin)):
    return import_job_snapshot(job_id)


@app.post("/api/session-stats/dedupe")
def session_stats_dedupe(month: str, _admin: None = Depends(require_admin)):
    month = validate_stats_month(month)
    result = session_stats_service().cleanup_exact_duplicate_sessions_for_month(month)
    return {"ok": True, "month": month, **result}


@app.get("/api/session-stats/sessions/{session_id}")
def session_stats_session_detail(session_id: int):
    detail = session_stats_service().get_session_detail(session_id)
    if not detail:
        raise HTTPException(404, "团记录不存在")
    return detail


@app.patch("/api/session-stats/sessions/{session_id}")
def session_stats_update_session(session_id: int, body: SessionStatsSessionUpdateRequest, _admin: None = Depends(require_admin)):
    if body.title is None and body.duration_hours is None:
        raise HTTPException(400, "没有可更新字段")
    try:
        detail = session_stats_service().update_session(session_id, title=body.title, duration_hours=body.duration_hours)
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    if not detail:
        raise HTTPException(404, "团记录不存在")
    return {"ok": True, "session": detail}


@app.delete("/api/session-stats/sessions/{session_id}")
def session_stats_delete_session(session_id: int, _admin: None = Depends(require_admin)):
    deleted = session_stats_service().delete_session(session_id)
    if not deleted:
        raise HTTPException(404, "团记录不存在")
    return {"ok": True, "session_id": session_id}


@app.get("/api/normalization/reviews")
def list_normalization_reviews():
    """公开列出规范化审核任务。直达审核入口不依赖后台登录。"""
    if not NORMALIZATION_REVIEW_DIR.exists():
        return {"items": [], "count": 0, "required_signatures": NORMALIZATION_SIGNATURES_REQUIRED}
    items = []
    for path in sorted(NORMALIZATION_REVIEW_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True):
        try:
            items.append(review_summary(json.loads(path.read_text(encoding="utf-8"))))
        except Exception:
            continue
    return {"items": items, "count": len(items), "required_signatures": NORMALIZATION_SIGNATURES_REQUIRED}


@app.post("/api/normalization/reviews")
def create_normalization_review(req: NormalizationReviewCreateRequest):
    """创建规范化审核任务。无密码，供批量规范脚本生成直达审核链接。"""
    full = public_resource_path(req.resource_path, must_exist=True)
    original = req.original_content if req.original_content is not None else read_text(full)[0]
    review_id = hashlib.sha1(f"{req.resource_path}\0{time.time()}\0{secrets.token_hex(8)}".encode("utf-8")).hexdigest()[:16]
    review = {
        "id": review_id,
        "resource_path": rel_posix(full),
        "title": req.title or title_for(full),
        "original_content": original,
        "normalized_content": req.normalized_content,
        "note": req.note or "",
        "created_at": now_iso(),
        "updated_at": now_iso(),
        "required_signatures": NORMALIZATION_SIGNATURES_REQUIRED,
        "signatures": [],
        "status": "pending",
    }
    save_review(review)
    return {**review_summary(review), "review_url": f"/normalize-review?id={review_id}"}


@app.get("/api/normalization/reviews/{review_id}")
def get_normalization_review(review_id: str):
    review = load_review(review_id)
    review_summary(review)
    return review


@app.post("/api/normalization/reviews/{review_id}/sign")
def sign_normalization_review(review_id: str, req: NormalizationReviewSignRequest):
    review = load_review(review_id)
    signer = req.signer.strip()
    if not signer:
        raise HTTPException(400, "签名不能为空")
    signatures = review.setdefault("signatures", [])
    now = now_iso()
    existing = next((s for s in signatures if str(s.get("signer", "")).strip() == signer), None)
    if existing:
        existing["note"] = req.note or ""
        existing["signed_at"] = now
    else:
        signatures.append({"signer": signer, "note": req.note or "", "signed_at": now})
    review["updated_at"] = now
    review_summary(review)
    save_review(review)
    return review


@app.post("/api/admin/login")
def login(req: LoginRequest, response: Response):
    pw = admin_password()
    if not pw:
        raise HTTPException(503, "未配置 ADMIN_PASSWORD，无法登录后台")
    if not secrets.compare_digest(req.password, pw):
        raise HTTPException(401, "密码错误")
    token = sign_payload({"role": "admin", "iat": int(time.time()), "exp": int(time.time()) + 86400})
    response.set_cookie(SESSION_COOKIE, token, httponly=True, samesite="lax", max_age=86400)
    return {"ok": True}


@app.post("/api/admin/logout")
def logout(response: Response):
    response.delete_cookie(SESSION_COOKIE)
    return {"ok": True}


@app.get("/api/admin/me")
def me(token: str | None = Cookie(default=None, alias=SESSION_COOKIE)):
    configured = bool(admin_password())
    return {"admin_configured": configured, "authenticated": configured and verify_token(token)}


@app.get("/api/resources")
def list_resources(
    q: str = "",
    root: Literal["", "序列库", "荣誉室"] = "",
    category: str = "",
    include_content: bool = False,
    kinds: Annotated[list[str] | None, Query()] = None,
    sides: Annotated[list[str] | None, Query()] = None,
    authors: Annotated[list[str] | None, Query()] = None,
    limit: int = 200,
    offset: int = 0,
):
    """资源列表/智能搜索。支持拼音、繁简、子序列、多 token 评分、分面。"""
    if root and root not in PUBLIC_ROOTS:
        raise HTTPException(403, "该资源根目录不在前台公开范围")
    roots = (root,) if root else PUBLIC_ROOTS
    return SEARCH_INDEX.search(
        q,
        roots=roots,
        category=category,
        kinds=kinds,
        sides=sides,
        authors=authors,
        limit=max(1, min(limit, 500)),
        offset=max(0, offset),
        include_content=include_content,
    )


@app.get("/api/facets")
def list_facets():
    """直接拿索引里的 facet 计数（不带任何查询过滤），便于前端首屏渲染分面初始态。"""
    SEARCH_INDEX.refresh()
    return {"facets": SEARCH_INDEX.facets(), "total": len(SEARCH_INDEX.all_entries())}


@app.get("/api/tree")
def tree():
    roots: dict[str, dict] = {r: {"name": r, "path": r, "count": 0, "children": {}} for r in PUBLIC_ROOTS}
    for path in iter_txt_files(PUBLIC_ROOTS):
        rel = Path(rel_posix(path))
        node = roots[rel.parts[0]]
        node["count"] += 1
        acc = rel.parts[0]
        for part in rel.parts[1:-1]:
            acc += "/" + part
            node = node["children"].setdefault(part, {"name": part, "path": acc, "count": 0, "children": {}})
            node["count"] += 1

    def compact(n: dict) -> dict:
        return {**{k: v for k, v in n.items() if k != "children"}, "children": [compact(c) for c in n["children"].values()]}

    return {"items": [compact(v) for v in roots.values()]}


@app.get("/api/resources/{path:path}")
def get_resource(path: str):
    full = public_resource_path(path, must_exist=True)
    content, encoding = read_text(full)
    return {**resource_entry(full), "content": content, "encoding": encoding}


@app.get("/api/raw/{path:path}", response_class=PlainTextResponse)
def get_raw(path: str):
    full = public_resource_path(path, must_exist=True)
    return read_text(full)[0]



PACKAGE_ROOT_EXTENSIONS = {".txt", ".html", ".htm", ".docx", ".doc", ".xlsx"}
PACKAGE_EDIT_EXTENSIONS = {".txt"}
PACKAGE_EXCLUDE_NAMES = {"README.txt"}


def safe_package_file_path(name: str, *, must_exist: bool = False, require_editable: bool = False) -> Path:
    if "\x00" in name or "/" in name or "\\" in name:
        raise HTTPException(400, "发行文件必须是仓库根目录单文件名")
    rel = Path(name)
    if rel.is_absolute() or ".." in rel.parts or len(rel.parts) != 1:
        raise HTTPException(400, "发行文件路径越界")
    if rel.name.startswith(".") or rel.name in PACKAGE_EXCLUDE_NAMES:
        raise HTTPException(400, "该文件不是发行说明文件")
    suffix = rel.suffix.lower()
    if suffix not in PACKAGE_ROOT_EXTENSIONS:
        raise HTTPException(400, "只允许发行包支持的根目录说明文件类型")
    if require_editable and suffix not in PACKAGE_EDIT_EXTENSIONS:
        raise HTTPException(400, "本轮只支持编辑根目录 .txt 发行文件")
    full = (REPO_ROOT / rel).resolve()
    if full.parent != REPO_ROOT.resolve():
        raise HTTPException(400, "发行文件必须位于仓库根目录")
    if must_exist and not full.is_file():
        raise HTTPException(404, "发行文件不存在")
    return full


def package_file_entry(path: Path) -> dict:
    stat = path.stat()
    title = path.stem
    if path.suffix.lower() == ".txt":
        try:
            title = title_for(path)
        except Exception:
            title = path.stem
    return {
        "path": path.name,
        "filename": path.name,
        "title": title,
        "size": stat.st_size,
        "mtime": stat.st_mtime,
        "editable": path.suffix.lower() in PACKAGE_EDIT_EXTENSIONS,
        "extension": path.suffix.lower(),
    }


def iter_package_files() -> Iterable[Path]:
    for path in sorted(REPO_ROOT.iterdir(), key=lambda p: sort_key(p.name)):
        if not path.is_file() or path.suffix.lower() not in PACKAGE_ROOT_EXTENSIONS:
            continue
        if path.name.startswith(".") or path.name in PACKAGE_EXCLUDE_NAMES:
            continue
        yield path


@app.get("/api/admin/package-files")
def list_package_files(_admin: None = Depends(require_admin)):
    items = [package_file_entry(path) for path in iter_package_files()]
    return {"items": items, "count": len(items)}


@app.get("/api/admin/package-files/{path:path}")
def get_package_file(path: str, _admin: None = Depends(require_admin)):
    full = safe_package_file_path(path, must_exist=True)
    if full.suffix.lower() != ".txt":
        return {**package_file_entry(full), "content": None, "encoding": None}
    content, encoding = read_text(full)
    return {**package_file_entry(full), "content": content, "encoding": encoding}


@app.post("/api/admin/package-files")
def create_package_file(req: PackageFileRequest, _admin: None = Depends(require_admin)):
    full = safe_package_file_path(req.path, require_editable=True)
    if full.exists() and not req.overwrite:
        raise HTTPException(409, "目标已存在；如需覆盖请设置 overwrite=true")
    before = git("status", "--short", "--", req.path)
    write_text_utf8(full, req.content)
    return {"ok": True, "path": full.name, "git_status_before": before, "git_status_after": git("status", "--short", "--", full.name)}


@app.put("/api/admin/package-files/{path:path}")
def save_package_file(path: str, req: PackageFileRequest, _admin: None = Depends(require_admin)):
    full = safe_package_file_path(path, must_exist=True, require_editable=True)
    before = git("status", "--short", "--", full.name)
    write_text_utf8(full, req.content)
    return {"ok": True, "path": full.name, "git_status_before": before, "git_status_after": git("status", "--short", "--", full.name)}

@app.post("/api/admin/resources")
def save_resource(req: SaveRequest, _admin: None = Depends(require_admin)):
    full = safe_resource_path(req.path)
    existed = full.exists()
    if existed and not req.overwrite:
        raise HTTPException(409, "目标已存在；如需覆盖请设置 overwrite=true")
    before = git("status", "--short", "--", req.path)
    write_text_utf8(full, req.content)
    SEARCH_INDEX.invalidate(req.path)
    return {"ok": True, "existed": existed, "path": rel_posix(full), "git_status_before": before, "git_status_after": git("status", "--short", "--", req.path)}


@app.put("/api/admin/resources/{path:path}")
def edit_resource(path: str, req: SaveRequest, _admin: None = Depends(require_admin)):
    full = safe_resource_path(path, must_exist=True)
    before = git("status", "--short", "--", path)
    write_text_utf8(full, req.content)
    SEARCH_INDEX.invalidate(path)
    return {"ok": True, "path": rel_posix(full), "git_status_before": before, "git_status_after": git("status", "--short", "--", path)}


@app.post("/api/admin/upload/{path:path}")
async def upload_txt(path: str, file: UploadFile = File(...), _admin: None = Depends(require_admin)):
    full = safe_resource_path(path, must_exist=True)
    if not (file.filename or "").lower().endswith(".txt"):
        raise HTTPException(400, "只支持上传 .txt")
    data = await file.read()
    if len(data) > 2_000_000:
        raise HTTPException(413, "TXT 过大")
    text = None
    used = ""
    for enc in TEXT_ENCODINGS:
        try:
            text = data.decode(enc)
            used = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = data.decode("utf-8", errors="replace")
        used = "utf-8-replace"
    before = git("status", "--short", "--", path)
    write_text_utf8(full, text)
    SEARCH_INDEX.invalidate(path)
    return {"ok": True, "path": rel_posix(full), "source_encoding": used, "git_status_before": before, "git_status_after": git("status", "--short", "--", path)}


@app.post("/api/admin/move")
def move_resource(req: MoveRequest, _admin: None = Depends(require_admin)):
    old = safe_resource_path(req.old_path, must_exist=True)
    new = safe_resource_path(req.new_path)
    if new.exists() and not req.overwrite:
        raise HTTPException(409, "目标已存在")
    before = git("status", "--short", "--", req.old_path, req.new_path)
    new.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(old), str(new))
    SEARCH_INDEX.invalidate(req.old_path)
    SEARCH_INDEX.invalidate(rel_posix(new))
    return {"ok": True, "old_path": req.old_path, "new_path": rel_posix(new), "git_status_before": before, "git_status_after": git("status", "--short", "--", req.old_path, req.new_path)}


@app.post("/api/admin/delete")
def delete_resource(req: DeleteRequest, _admin: None = Depends(require_admin)):
    full = safe_resource_path(req.path, must_exist=True)
    before = git("status", "--short", "--", req.path)
    full.unlink()
    SEARCH_INDEX.invalidate(req.path)
    return {"ok": True, "path": req.path, "git_status_before": before, "git_status_after": git("status", "--short", "--", req.path)}


HONOR_FILENAME_RE = re.compile(r"^(\d{3,})】.+\.txt$", re.IGNORECASE)


HONOR_CATEGORY_MAP = {
    "特质改造": "特质",
    "职业": "职业",
    "技能表": "技能表",
    "能量池": "能量池",
    "魔药列表": "魔药列表",
    "成就": "成就",
}


def suggest_honor_category_for_path(path: str) -> str:
    try:
        parts = Path(path).parts
        if len(parts) >= 2 and parts[0] == "序列库":
            return HONOR_CATEGORY_MAP.get(parts[1], "其他")
    except Exception:
        pass
    return "其他"

def clean_filename_title(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r"^\d+】", "", stem).strip() or stem


def validate_honor_category(category: str) -> str:
    cat = category.strip()
    if not cat or "\x00" in cat or "/" in cat or "\\" in cat or cat in {".", ".."}:
        raise HTTPException(400, "荣誉室大类只能是一级目录名称")
    if Path(cat).is_absolute() or ".." in Path(cat).parts:
        raise HTTPException(400, "荣誉室大类不能越界")
    return cat


def validate_honor_target_path(path: str) -> Path:
    rel = Path(path)
    if rel.is_absolute() or ".." in rel.parts or len(rel.parts) != 3 or rel.parts[0] != "荣誉室":
        raise HTTPException(400, "target_path 必须形如 荣誉室/<一级大类>/<NNN】标题.txt>")
    validate_honor_category(rel.parts[1])
    if rel.suffix.lower() != ".txt" or not HONOR_FILENAME_RE.match(rel.name):
        raise HTTPException(400, "荣誉室文件名必须形如 NNN】标题.txt")
    return safe_resource_path(rel.as_posix())


def honor_category_info(category: str) -> dict:
    cat = validate_honor_category(category)
    base = REPO_ROOT / "荣誉室" / cat
    max_no = 0
    count = 0
    if base.exists():
        for path in base.iterdir():
            if not path.is_file() or path.suffix.lower() != ".txt":
                continue
            count += 1
            m = re.match(r"^(\d+)】", path.name)
            if m:
                max_no = max(max_no, int(m.group(1)))
    next_no = max_no + 1
    return {"category": cat, "count": count, "next_number": next_no, "next_prefix": f"{next_no:03d}】"}


def cleanup_empty_dirs_after_move(old_file: Path) -> None:
    for parent in old_file.parents:
        if parent == REPO_ROOT / "序列库" or parent == REPO_ROOT / "荣誉室" or parent == REPO_ROOT:
            break
        try:
            parent.rmdir()
        except OSError:
            break


@app.get("/api/admin/honor-categories")
def honor_categories(path: str | None = None, _admin: None = Depends(require_admin)):
    base = REPO_ROOT / "荣誉室"
    names = set()
    if base.exists():
        names.update(p.name for p in base.iterdir() if p.is_dir())
    names.update(["其他", "成就", "技能表", "特质", "职业", "能量池", "魔药列表"])
    suggested = suggest_honor_category_for_path(path or "")
    return {"items": [honor_category_info(name) for name in sorted(names)], "suggested_category": suggested}


@app.post("/api/admin/move-to-honor")
def move_to_honor(req: MoveToHonorRequest, _admin: None = Depends(require_admin)):
    old = safe_resource_path(req.path, must_exist=True)
    old_rel = Path(req.path)
    if old_rel.parts[0] == "荣誉室":
        raise HTTPException(400, "资源已在荣誉室")

    if req.target_path:
        new = validate_honor_target_path(req.target_path)
    else:
        cat = validate_honor_category(req.category or suggest_honor_category_for_path(req.path))
        title_source = req.filename or req.title or clean_filename_title(old.name)
        title = clean_filename_title(title_source).replace("/", "／").replace("\\", "＼").strip()
        if title.lower().endswith(".txt"):
            title = title[:-4].strip()
        if not title:
            raise HTTPException(400, "标题不能为空")
        info = honor_category_info(cat)
        new = safe_resource_path(Path("荣誉室", cat, f"{info['next_prefix']}{title}.txt").as_posix())

    if new.exists() and not req.overwrite:
        raise HTTPException(409, "目标已存在")
    before = git("status", "--short", "--", req.path, rel_posix(new))
    new.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(old), str(new))
    cleanup_empty_dirs_after_move(old)
    SEARCH_INDEX.invalidate(req.path)
    new_path = rel_posix(new)
    return {
        "ok": True,
        "old_path": req.path,
        "new_path": new_path,
        "category": Path(new_path).parts[1],
        "title": clean_filename_title(new.name),
        "git_status_before": before,
        "git_status_after": git("status", "--short", "--", req.path, new_path),
    }


@app.get("/api/git/info")
def git_info():
    tag = git("describe", "--tags", "--abbrev=0")
    head_full = git("rev-parse", "HEAD")
    head_short = git("rev-parse", "--short", "HEAD")
    branch = git("branch", "--show-current")
    status_short = git("status", "--short")
    status_branch = git("status", "--short", "--branch")
    upstream = git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}")
    ahead_behind = None
    tracking = upstream["stdout"].strip() if upstream["returncode"] == 0 else None
    if tracking:
        counts = git("rev-list", "--left-right", "--count", f"HEAD...{tracking}")
        if counts["returncode"] == 0:
            parts = counts["stdout"].strip().split()
            if len(parts) == 2:
                ahead_behind = {"ahead": int(parts[0]), "behind": int(parts[1]), "tracking": tracking}
    latest_tag = tag["stdout"].strip() if tag["returncode"] == 0 else None
    head_tag = git("tag", "--points-at", "HEAD")
    remote_main = git("rev-parse", "--short", "origin/main")
    is_dirty = bool(status_short["stdout"].strip())
    return {
        "latest_tag": latest_tag,
        "head": head_short,  # 兼容旧前端：短 hash
        "head_full": head_full["stdout"].strip() if head_full["returncode"] == 0 else None,
        "head_short": head_short["stdout"].strip() if head_short["returncode"] == 0 else None,
        "head_tags": [t for t in head_tag["stdout"].splitlines() if t.strip()] if head_tag["returncode"] == 0 else [],
        "branch": branch,
        "branch_name": branch["stdout"].strip() if branch["returncode"] == 0 else None,
        "status": status_short,
        "status_short": status_short["stdout"],
        "status_branch": status_branch["stdout"],
        "is_dirty": is_dirty,
        "tracking": tracking,
        "ahead_behind": ahead_behind,
        "remote_main": remote_main["stdout"].strip() if remote_main["returncode"] == 0 else None,
        "admin_configured": bool(admin_password()),
        "wiki_configured": wiki_credentials_configured(),
    }


def parse_name_status(output: str) -> dict:
    changes = {"added": [], "modified": [], "deleted": [], "renamed": []}
    for line in output.splitlines():
        if not line.strip():
            continue
        parts = line.split("\t")
        code = parts[0]
        status = code[0]
        if status == "A" and len(parts) >= 2:
            changes["added"].append(parts[1])
        elif status == "M" and len(parts) >= 2:
            changes["modified"].append(parts[1])
        elif status == "D" and len(parts) >= 2:
            changes["deleted"].append(parts[1])
        elif status == "R" and len(parts) >= 3:
            changes["renamed"].append({"old": parts[1], "new": parts[2], "score": code[1:]})
    return changes


def normalized_resource_path_key(path: str) -> str:
    """用于未提交工作区的改名兜底：展示规范化括号不应让旧版对比丢失。"""
    return path.replace("[", "【").replace("]", "】")


def pair_normalized_worktree_renames(summary: dict) -> None:
    added_by_key: dict[str, list[str]] = {}
    for path in summary["added"]:
        added_by_key.setdefault(normalized_resource_path_key(path), []).append(path)

    paired_added: set[str] = set()
    paired_deleted: set[str] = set()
    inferred: list[dict] = []
    for old_path in summary["deleted"]:
        key = normalized_resource_path_key(old_path)
        candidates = [path for path in added_by_key.get(key, []) if path not in paired_added]
        if len(candidates) != 1:
            continue
        new_path = candidates[0]
        if old_path == new_path:
            continue
        paired_deleted.add(old_path)
        paired_added.add(new_path)
        inferred.append({"old": old_path, "new": new_path, "score": "规范化"})

    if not inferred:
        return
    summary["added"] = [path for path in summary["added"] if path not in paired_added]
    summary["deleted"] = [path for path in summary["deleted"] if path not in paired_deleted]
    existing_pairs = {(item["old"], item["new"]) for item in summary["renamed"]}
    for item in inferred:
        if (item["old"], item["new"]) not in existing_pairs:
            summary["renamed"].append(item)


def category_for_path(path: str) -> str:
    parts = Path(path).parts
    if len(parts) == 1:
        return "发行文件"
    if len(parts) <= 2:
        return "根目录"
    return "/".join(parts[1:-1]) or "根目录"


def root_for_path(path: str) -> str:
    parts = Path(path).parts
    if len(parts) == 1:
        return "发行包根目录"
    return parts[0] if parts else ""


def inferred_title_for_path(path: str) -> str:
    return strip_number_prefix(Path(path).name)


def change_item(
    path: str,
    *,
    old_path: str | None = None,
    score: str | None = None,
    allowed_roots: tuple[str, ...] = ALLOWED_ROOTS,
) -> dict:
    full = (REPO_ROOT / Path(path)).resolve()
    exists = full.is_file() and root_for_path(path) in allowed_roots
    entry = resource_entry(full) if exists else None
    return {
        "title": entry["title"] if entry else inferred_title_for_path(path),
        "path": path,
        "old_path": old_path,
        "category": entry["category"] if entry else category_for_path(path),
        "root": entry["root"] if entry else root_for_path(path),
        "size": entry["size"] if entry else None,
        "exists": exists,
        "score": score,
    }


def dedupe_keep_order(items: list[str]) -> list[str]:
    seen = set()
    out = []
    for item in items:
        if item not in seen:
            seen.add(item)
            out.append(item)
    return out


def default_from_ref() -> str:
    tag = git("describe", "--tags", "--abbrev=0")
    return tag["stdout"].strip() if tag["returncode"] == 0 and tag["stdout"].strip() else "HEAD"


def safe_ref(ref: str) -> str:
    if not re.fullmatch(r"[A-Za-z0-9._/\-]+", ref):
        raise HTTPException(400, "版本引用非法")
    exists = git("rev-parse", "--verify", f"{ref}^{{commit}}")
    if exists["returncode"] != 0:
        raise HTTPException(404, f"版本引用不存在: {ref}")
    return ref


def safe_git_path(rel_path: str, *, allowed_roots: tuple[str, ...]) -> str:
    if "\x00" in rel_path:
        raise HTTPException(400, "路径包含非法字符")
    normalized = rel_path.replace("\\", "/").strip("/")
    rel = Path(normalized)
    if rel.is_absolute() or ".." in rel.parts:
        raise HTTPException(400, "路径越界")
    if not rel.parts or rel.parts[0] not in allowed_roots:
        root_hint = " 或 ".join(f"{root}/" for root in allowed_roots)
        raise HTTPException(400, f"路径必须位于 {root_hint} 内")
    if rel.suffix.lower() != ".txt":
        raise HTTPException(400, "只允许查看 .txt 资源差异")
    full = (REPO_ROOT / rel).resolve()
    roots = [(REPO_ROOT / root).resolve() for root in allowed_roots]
    if not any(full == root or root in full.parents for root in roots):
        raise HTTPException(400, "路径越界")
    return rel.as_posix()


def decode_blob(data: bytes) -> tuple[str, str]:
    for enc in TEXT_ENCODINGS:
        try:
            return data.decode(enc), enc
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def git_text_at(ref: str, rel_path: str) -> tuple[str, str, bool]:
    result = git_bytes("show", f"{ref}:{rel_path}")
    if result["returncode"] != 0:
        return "", "", False
    text, enc = decode_blob(result["stdout"])
    return text, enc, True


def current_text_at(rel_path: str) -> tuple[str, str, bool]:
    full = (REPO_ROOT / Path(rel_path)).resolve()
    if not full.is_file():
        return "", "", False
    text, enc = read_text(full)
    return text, enc, True


def build_line_diff(old_text: str, new_text: str, *, context: int = 3, max_rows: int = 900) -> dict:
    old_lines = old_text.splitlines()
    new_lines = new_text.splitlines()
    matcher = difflib.SequenceMatcher(None, old_lines, new_lines)
    rows: list[dict] = []
    additions = 0
    deletions = 0

    for group in matcher.get_grouped_opcodes(n=context):
        if rows:
            rows.append({"type": "gap"})
        for tag, i1, i2, j1, j2 in group:
            if tag == "equal":
                for offset, line in enumerate(old_lines[i1:i2]):
                    rows.append({"type": "context", "old_no": i1 + offset + 1, "new_no": j1 + offset + 1, "text": line})
            elif tag == "delete":
                for offset, line in enumerate(old_lines[i1:i2]):
                    deletions += 1
                    rows.append({"type": "removed", "old_no": i1 + offset + 1, "new_no": None, "text": line})
            elif tag == "insert":
                for offset, line in enumerate(new_lines[j1:j2]):
                    additions += 1
                    rows.append({"type": "added", "old_no": None, "new_no": j1 + offset + 1, "text": line})
            elif tag == "replace":
                for offset, line in enumerate(old_lines[i1:i2]):
                    deletions += 1
                    rows.append({"type": "removed", "old_no": i1 + offset + 1, "new_no": None, "text": line})
                for offset, line in enumerate(new_lines[j1:j2]):
                    additions += 1
                    rows.append({"type": "added", "old_no": None, "new_no": j1 + offset + 1, "text": line})
            if len(rows) >= max_rows:
                return {"rows": rows[:max_rows], "truncated": True, "additions": additions, "deletions": deletions}

    if not rows and old_text == new_text:
        preview = old_lines[: min(len(old_lines), 60)]
        rows = [{"type": "context", "old_no": i + 1, "new_no": i + 1, "text": line} for i, line in enumerate(preview)]
    return {"rows": rows, "truncated": False, "additions": additions, "deletions": deletions}


def build_readable_changes(summary: dict, from_ref: str, *, allowed_roots: tuple[str, ...] = ALLOWED_ROOTS) -> dict:
    labels = {"added": "新增", "modified": "修改", "deleted": "删除", "renamed": "移动/改名"}
    readable = {
        "added": [change_item(p, allowed_roots=allowed_roots) for p in dedupe_keep_order(summary["added"])],
        "modified": [change_item(p, allowed_roots=allowed_roots) for p in dedupe_keep_order(summary["modified"])],
        "deleted": [change_item(p, allowed_roots=allowed_roots) for p in dedupe_keep_order(summary["deleted"])],
        "renamed": [change_item(r["new"], old_path=r["old"], score=r.get("score"), allowed_roots=allowed_roots) for r in summary["renamed"]],
    }
    for items in readable.values():
        items.sort(key=lambda i: (i["root"], i["category"], i["title"], i["path"]))

    grouped: dict[str, list[dict]] = {}
    for kind, items in readable.items():
        for item in items:
            grouped.setdefault(item["category"] or "根目录", []).append({"type": kind, "type_label": labels[kind], **item})

    stats = {
        "added": len(readable["added"]),
        "modified": len(readable["modified"]),
        "deleted": len(readable["deleted"]),
        "renamed": len(readable["renamed"]),
    }
    stats["total"] = sum(stats.values())

    lines = [
        f"从 {from_ref} 到 latest，共 {stats['total']} 项变更：新增 {stats['added']}，修改 {stats['modified']}，删除 {stats['deleted']}，移动/改名 {stats['renamed']}。"
    ]
    for kind in ("added", "modified", "deleted", "renamed"):
        lines.append("")
        lines.append(f"【{labels[kind]}】")
        if not readable[kind]:
            lines.append("- 无")
            continue
        for item in readable[kind]:
            cat = item["category"] or "根目录"
            if kind == "renamed":
                lines.append(f"- {cat}：{item['title']}（{item['old_path']} -> {item['path']}）")
            else:
                lines.append(f"- {cat}：{item['title']}（{item['path']}）")

    return {"stats": stats, "readable": readable, "groups": grouped, "text": "\n".join(lines), "markdown": "\n".join(lines)}


@app.get("/api/git/changes")
def git_changes(from_ref: str | None = None, public_only: bool = False):
    if not from_ref:
        from_ref = default_from_ref()
    else:
        from_ref = safe_ref(from_ref)
    roots = PUBLIC_ROOTS if public_only else ALLOWED_ROOTS
    pathspecs = [*roots] if public_only else [*roots, ":(top)*.txt"]
    committed = git("diff", "--name-status", "--find-renames", "--diff-filter=AMDR", from_ref, "HEAD", "--", *pathspecs)
    worktree = git("diff", "--name-status", "--find-renames", "--diff-filter=AMDR", "HEAD", "--", *pathspecs)
    untracked = git("ls-files", "--others", "--exclude-standard", "--", *pathspecs)
    summary = parse_name_status((committed["stdout"] if committed["returncode"] == 0 else "") + "\n" + (worktree["stdout"] if worktree["returncode"] == 0 else ""))
    for p in untracked["stdout"].splitlines():
        if p.endswith(".txt") and p not in summary["added"]:
            summary["added"].append(p)
    pair_normalized_worktree_renames(summary)
    friendly = build_readable_changes(summary, from_ref, allowed_roots=roots)
    return {
        "from_ref": from_ref,
        "to": "working-tree/latest",
        "public_only": public_only,
        "summary": summary,
        **friendly,
        "raw": {"committed": committed, "worktree": worktree, "untracked": untracked},
    }


@app.get("/api/git/change-detail/{path:path}")
def git_change_detail(
    path: str,
    kind: Literal["added", "modified", "deleted", "renamed"] = "modified",
    old_path: str | None = None,
    from_ref: str | None = None,
    public_only: bool = False,
):
    from_ref = safe_ref(from_ref) if from_ref else default_from_ref()
    roots = PUBLIC_ROOTS if public_only else ALLOWED_ROOTS
    rel_path = safe_git_path(path, allowed_roots=roots)
    old_rel_path = safe_git_path(old_path, allowed_roots=roots) if old_path else rel_path

    old_text = ""
    old_encoding = ""
    old_exists = False
    new_text = ""
    new_encoding = ""
    new_exists = False

    if kind != "added":
        old_text, old_encoding, old_exists = git_text_at(from_ref, old_rel_path)
    if kind != "deleted":
        new_text, new_encoding, new_exists = current_text_at(rel_path)
        if not new_exists:
            new_text, new_encoding, new_exists = git_text_at("HEAD", rel_path)

    if kind == "added" and not new_exists:
        raise HTTPException(404, "新增资源当前不存在，无法显示内容")
    if kind == "deleted" and not old_exists:
        raise HTTPException(404, "旧版本中找不到该资源，无法显示删除内容")
    if kind in ("modified", "renamed") and not old_exists and not new_exists:
        raise HTTPException(404, "找不到可比较的资源内容")

    diff = build_line_diff(old_text, new_text)
    title_path = rel_path if new_exists else old_rel_path
    return {
        "from_ref": from_ref,
        "to": "working-tree/latest",
        "kind": kind,
        "path": rel_path,
        "old_path": old_path,
        "title": inferred_title_for_path(title_path),
        "old_exists": old_exists,
        "new_exists": new_exists,
        "old_encoding": old_encoding,
        "new_encoding": new_encoding,
        "old_line_count": len(old_text.splitlines()),
        "new_line_count": len(new_text.splitlines()),
        **diff,
    }


@app.post("/api/admin/wiki/sync")
def wiki_sync(dry_run: bool = True, skip_honor: bool = False, _admin: None = Depends(require_admin)):
    user = os.getenv("WIKI_USER") or os.getenv("FANDOM_USER")
    password = os.getenv("WIKI_PASSWORD") or os.getenv("FANDOM_PASSWORD")
    if not dry_run and (not user or not password):
        raise HTTPException(503, "未配置 WIKI_USER/WIKI_PASSWORD，无法执行真实同步")
    args = [sys.executable, "wiki/sync_to_wiki.py", "--user", user or "DRY_RUN_USER", "--password", password or "DRY_RUN_PASSWORD"]
    if dry_run:
        args.append("--dry-run")
    if skip_honor:
        args.append("--skip-honor")
    result = run_cmd(args, timeout=3600)
    return result


@app.post("/api/admin/publish")
def publish(req: PublishRequest, _admin: None = Depends(require_admin)):
    version = req.version if req.version.startswith("v") else f"v{req.version}"
    steps = []
    # 只加入实际维护/发布需要的范围，避免把被 .gitignore 忽略的构建产物、venv、wiki 目录等混入。
    # 如果后续确实要发布 wiki 脚本，应单独做明确入口，而不是 publish 时默认 add。
    publish_paths = [
        "序列库",
        "荣誉室",
        ":(top)*.txt",
        ".gitignore",
        "web/README.md",
        "web/backend/app",
        "web/backend/*.py",
        "web/backend/requirements.txt",
        "web/normalization_reviews/README.md",
        "web/frontend/index.html",
        "web/frontend/package.json",
        "web/frontend/package-lock.json",
        "web/frontend/src",
        "web/frontend/tsconfig.json",
        "web/frontend/vite.config.ts",
    ]
    steps.append(git("status", "--short"))
    add = git("add", *publish_paths)
    steps.append(add)
    if add["returncode"] != 0:
        return {"ok": False, "version": version, "failed_at": "git add", "steps": steps}
    commit = git("commit", "-m", req.message)
    steps.append(commit)
    if commit["returncode"] != 0 and "nothing to commit" not in (commit["stdout"] + commit["stderr"]).lower():
        return {"ok": False, "version": version, "steps": steps}
    tag = git("tag", "-a", version, "-m", req.message)
    steps.append(tag)
    if tag["returncode"] != 0:
        return {"ok": False, "version": version, "steps": steps}
    if req.push:
        steps.append(git("push", "origin", req.branch, "--tags"))
    return {"ok": all(s["returncode"] == 0 for s in steps[1:]), "version": version, "steps": steps}
