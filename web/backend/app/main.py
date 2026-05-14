from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
import shutil
import sys
import subprocess
import time
from pathlib import Path
from typing import Iterable, Literal

from fastapi import Cookie, Depends, FastAPI, File, HTTPException, Response, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import PlainTextResponse
from pydantic import BaseModel, Field

REPO_ROOT = Path(__file__).resolve().parents[3]
ALLOWED_ROOTS = ("序列库", "荣誉室")
PUBLIC_ROOTS = ("序列库",)
TEXT_ENCODINGS = ("utf-8", "utf-8-sig", "gbk", "gb2312", "big5")
SESSION_COOKIE = "seqlib_admin"

app = FastAPI(title="宏观界域强化序列库 Web API", version="0.1.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=os.getenv("WEB_CORS_ORIGINS", "http://localhost:5173").split(","),
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


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


def strip_number_prefix(name: str) -> str:
    stem = Path(name).stem
    return re.sub(r"^\d+】", "", stem).strip() or stem


def title_for(path: Path) -> str:
    try:
        text, _ = read_text(path)
        first = next((line.strip() for line in text.splitlines() if line.strip()), "")
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


def run_cmd(args: list[str], timeout: int = 120) -> dict:
    started = time.time()
    try:
        p = subprocess.run(
            args,
            cwd=REPO_ROOT,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=timeout,
        )
        return {
            "cmd": args,
            "returncode": p.returncode,
            "stdout": p.stdout,
            "stderr": p.stderr,
            "seconds": round(time.time() - started, 3),
        }
    except subprocess.TimeoutExpired as e:
        return {"cmd": args, "returncode": 124, "stdout": e.stdout or "", "stderr": f"超时: {e}", "seconds": timeout}


def git(*args: str) -> dict:
    return run_cmd(["git", "-c", "core.quotepath=false", *args])


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


class MoveToHonorRequest(BaseModel):
    path: str
    category: str | None = None
    title: str | None = None
    filename: str | None = None
    target_path: str | None = None
    overwrite: bool = False


class PublishRequest(BaseModel):
    version: str = Field(pattern=r"^v?\d+(?:\.\d+)*(?:[-\w.]*)?$")
    message: str = "更新序列库 latest"
    branch: str = Field(default="main", pattern=r"^[A-Za-z0-9._/-]+$")
    push: bool = True


def wiki_credentials_configured() -> bool:
    return bool((os.getenv("WIKI_USER") or os.getenv("FANDOM_USER")) and (os.getenv("WIKI_PASSWORD") or os.getenv("FANDOM_PASSWORD")))



@app.get("/api/health")
def health():
    return {"ok": True, "repo": str(REPO_ROOT), "admin_configured": bool(admin_password()), "wiki_configured": wiki_credentials_configured()}


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
def list_resources(q: str = "", root: Literal["", "序列库", "荣誉室"] = "", category: str = "", include_content: bool = True):
    query = q.strip().lower()
    rows = []
    public_roots = PUBLIC_ROOTS if root != "荣誉室" else ()
    for path in iter_txt_files(public_roots):
        entry = resource_entry(path)
        if root and entry["root"] != root:
            continue
        if category and not entry["category"].startswith(category.strip("/")):
            continue
        if query:
            hay = f"{entry['path']}\n{entry['filename']}\n{entry['title']}".lower()
            if include_content:
                try:
                    hay += "\n" + read_text(path)[0].lower()
                except Exception:
                    pass
            if query not in hay:
                continue
        rows.append(entry)
    return {"items": rows, "count": len(rows)}


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


@app.post("/api/admin/resources")
def save_resource(req: SaveRequest, _admin: None = Depends(require_admin)):
    full = safe_resource_path(req.path)
    existed = full.exists()
    if existed and not req.overwrite:
        raise HTTPException(409, "目标已存在；如需覆盖请设置 overwrite=true")
    before = git("status", "--short", "--", req.path)
    write_text_utf8(full, req.content)
    return {"ok": True, "existed": existed, "path": rel_posix(full), "git_status_before": before, "git_status_after": git("status", "--short", "--", req.path)}


@app.put("/api/admin/resources/{path:path}")
def edit_resource(path: str, req: SaveRequest, _admin: None = Depends(require_admin)):
    full = safe_resource_path(path, must_exist=True)
    before = git("status", "--short", "--", path)
    write_text_utf8(full, req.content)
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
    return {"ok": True, "old_path": req.old_path, "new_path": rel_posix(new), "git_status_before": before, "git_status_after": git("status", "--short", "--", req.old_path, req.new_path)}


@app.post("/api/admin/delete")
def delete_resource(req: DeleteRequest, _admin: None = Depends(require_admin)):
    full = safe_resource_path(req.path, must_exist=True)
    before = git("status", "--short", "--", req.path)
    full.unlink()
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


def category_for_path(path: str) -> str:
    parts = Path(path).parts
    if len(parts) <= 2:
        return "根目录"
    return "/".join(parts[1:-1]) or "根目录"


def root_for_path(path: str) -> str:
    parts = Path(path).parts
    return parts[0] if parts else ""


def inferred_title_for_path(path: str) -> str:
    return strip_number_prefix(Path(path).name)


def change_item(path: str, *, old_path: str | None = None, score: str | None = None) -> dict:
    full = (REPO_ROOT / Path(path)).resolve()
    exists = full.is_file() and root_for_path(path) in ALLOWED_ROOTS
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


def build_readable_changes(summary: dict, from_ref: str) -> dict:
    labels = {"added": "新增", "modified": "修改", "deleted": "删除", "renamed": "移动/改名"}
    readable = {
        "added": [change_item(p) for p in dedupe_keep_order(summary["added"])],
        "modified": [change_item(p) for p in dedupe_keep_order(summary["modified"])],
        "deleted": [change_item(p) for p in dedupe_keep_order(summary["deleted"])],
        "renamed": [change_item(r["new"], old_path=r["old"], score=r.get("score")) for r in summary["renamed"]],
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
def git_changes(from_ref: str | None = None):
    if not from_ref:
        tag = git("describe", "--tags", "--abbrev=0")
        from_ref = tag["stdout"].strip() if tag["returncode"] == 0 and tag["stdout"].strip() else "HEAD"
    committed = git("diff", "--name-status", "--find-renames", "--diff-filter=AMDR", from_ref, "HEAD", "--", *ALLOWED_ROOTS)
    worktree = git("diff", "--name-status", "--find-renames", "--diff-filter=AMDR", "HEAD", "--", *ALLOWED_ROOTS)
    untracked = git("ls-files", "--others", "--exclude-standard", "--", *ALLOWED_ROOTS)
    summary = parse_name_status((committed["stdout"] if committed["returncode"] == 0 else "") + "\n" + (worktree["stdout"] if worktree["returncode"] == 0 else ""))
    for p in untracked["stdout"].splitlines():
        if p.endswith(".txt") and p not in summary["added"]:
            summary["added"].append(p)
    friendly = build_readable_changes(summary, from_ref)
    return {
        "from_ref": from_ref,
        "to": "working-tree/latest",
        "summary": summary,
        **friendly,
        "raw": {"committed": committed, "worktree": worktree, "untracked": untracked},
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
