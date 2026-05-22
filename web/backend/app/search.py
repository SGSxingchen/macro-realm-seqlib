"""智能搜索索引：在内存里维护 TXT 资源的归一化文本、拼音、字符集，
按 mtime 做增量缓存，支持多 token + 子序列 + 拼音 + 紧凑度评分。

设计要点：
- 不依赖外部数据库，进程内字典 + RLock；
- pypinyin / opencc 可选，缺失时拼音/繁简归一自动降级；
- 子序列匹配用于 "强制驱散" → "强驱散"、"百夫长" → "夫长" 这类少字、跨字模糊；
- 多 token AND 语义：query 拆 token 后每个都得在某个字段命中，避免噪音命中。
"""

from __future__ import annotations

import re
import threading
import time
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable

try:
    from pypinyin import lazy_pinyin, Style  # type: ignore
    _PINYIN_OK = True
except Exception:
    _PINYIN_OK = False

try:
    from opencc import OpenCC  # type: ignore
    _OCC = OpenCC("t2s")
    _OPENCC_OK = True
    def _t2s(text: str) -> str:
        try:
            return _OCC.convert(text)
        except Exception:
            return text
except Exception:
    _OPENCC_OK = False
    def _t2s(text: str) -> str:
        return text


TOKEN_SPLIT_RE = re.compile(
    r"[\s,.;:!?，。；：！？、“”‘’《》<>()\[\]{}【】「」『』|/\\\-_+~`@#$%^&*=]+"
)
AUTHOR_LINE_RE = re.compile(
    r"[（(](?:制作人|作者|原作者|审核人|修改人|调整人|重置人|复查人|策划)\s*[：:]\s*([^)）]+?)[)）]"
)
TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "gbk", "gb2312", "big5")
SIDE_NAMES = ("战技侧", "神秘侧", "科技侧", "特殊侧")


def _read_text(path: Path) -> str:
    for enc in TEXT_ENCODINGS:
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def normalize(text: str) -> str:
    """全角→半角 + 繁→简 + 小写。无 opencc 时只做 NFKC + 小写。"""
    if not text:
        return ""
    return _t2s(unicodedata.normalize("NFKC", text)).lower()


def _clean_title_for_pinyin(title: str) -> str:
    """去掉编号前缀、英文/数字/标点，留中文做拼音键。"""
    cleaned = re.sub(r"^\d+】", "", title)
    cleaned = re.sub(r"[《》〈〉「」『』【】\[\]()（）<>{}…—\-_/\\\s\d]", "", cleaned)
    return cleaned


def pinyin_of(text: str) -> tuple[str, str]:
    """返回 (拼音连写, 拼音首字母)。无 pypinyin 时返回空串。"""
    if not _PINYIN_OK or not text:
        return ("", "")
    cleaned = _clean_title_for_pinyin(text)
    if not cleaned:
        return ("", "")
    pieces = lazy_pinyin(cleaned, style=Style.NORMAL)
    full = "".join(pieces)
    initials = "".join(p[0] for p in pieces if p)
    return (full, initials)


def tokenize_query(q: str) -> list[str]:
    """把一句话拆成多个查询 token。空白/标点都切分；保留单字 token（中文里"驱"也算）。"""
    if not q:
        return []
    parts = [p for p in TOKEN_SPLIT_RE.split(q.strip()) if p]
    seen: set[str] = set()
    out: list[str] = []
    for p in parts:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out


def subseq_match(token: str, text: str) -> tuple[bool, float, int, int]:
    """token 字符是否按顺序出现在 text 里。返回 (是否命中, 紧凑度0~1, 起始, 结束含)。
    紧凑度 = len(token)/span，1.0 表示连续。"""
    n = len(token)
    if n == 0 or not text:
        return (False, 0.0, 0, 0)
    pos = 0
    first = -1
    last = -1
    for i, ch in enumerate(text):
        if ch == token[pos]:
            if first < 0:
                first = i
            last = i
            pos += 1
            if pos == n:
                span = last - first + 1
                return (True, n / span if span else 1.0, first, last)
    return (False, 0.0, 0, 0)


def best_subseq_window(token: str, text: str) -> tuple[int, int, float] | None:
    """在 text 中寻找 token 的所有子序列出现，返回紧凑度最高那次的 (start, end, compact)。
    单纯贪心 subseq_match 的窗口可能很长；用动态搜索找最佳起点。"""
    n = len(token)
    if n == 0 or not text:
        return None
    best: tuple[int, int, float] | None = None
    text_len = len(text)
    first_char = token[0]
    i = 0
    while i < text_len:
        if text[i] != first_char:
            i += 1
            continue
        # 从 i 开始试一次
        pos = 1
        last = i
        for j in range(i + 1, text_len):
            if text[j] == token[pos]:
                last = j
                pos += 1
                if pos == n:
                    break
        if pos == n:
            span = last - i + 1
            compact = n / span
            if best is None or compact > best[2]:
                best = (i, last, compact)
            if compact >= 1.0:
                return best
        i += 1
    return best


@dataclass
class IndexEntry:
    path: str
    filename: str
    title: str
    root: str
    category: str
    size: int
    mtime: float

    side: str = ""           # 战技侧/神秘侧/科技侧/特殊侧
    top_kind: str = ""       # 职业/特质改造/技能表/能量池/公共建筑
    authors: tuple[str, ...] = ()

    body: str = ""
    title_norm: str = ""
    filename_norm: str = ""
    path_norm: str = ""
    body_norm: str = ""
    title_pinyin: str = ""
    title_pyinit: str = ""
    body_charset: frozenset = field(default_factory=frozenset)

    def to_meta(self) -> dict:
        return {
            "path": self.path,
            "filename": self.filename,
            "title": self.title,
            "root": self.root,
            "category": self.category,
            "size": self.size,
            "mtime": self.mtime,
            "side": self.side,
            "top_kind": self.top_kind,
            "authors": list(self.authors),
        }


class SearchIndex:
    """进程内倒排索引。线程安全，惰性按 mtime 增量刷新。"""

    REFRESH_MIN_INTERVAL = 2.0  # 秒；高频查询时跳过重复 stat

    def __init__(self, repo_root: Path, roots: tuple[str, ...]):
        self.repo_root = repo_root
        self.roots = roots
        self._entries: dict[str, IndexEntry] = {}
        self._lock = threading.RLock()
        self._last_refresh: float = 0.0

    # ---------- 构建 ----------

    def _make_entry(self, abs_path: Path, rel_path: str) -> IndexEntry | None:
        try:
            stat = abs_path.stat()
            content = _read_text(abs_path)
        except Exception:
            return None
        rel = Path(rel_path)
        parts = rel.parts
        root = parts[0] if parts else ""
        category = "/".join(parts[1:-1])
        first_line = next((ln.strip().lstrip("\ufeff") for ln in content.splitlines() if ln.strip()), "")
        title = first_line or re.sub(r"^\d+】", "", rel.stem).strip() or rel.stem
        side = next((p for p in parts if p in SIDE_NAMES), "")
        top_kind = parts[1] if len(parts) >= 2 else ""
        authors = tuple(sorted({m.group(1).strip() for m in AUTHOR_LINE_RE.finditer(content) if m.group(1).strip()}))

        title_norm = normalize(title)
        body_norm = normalize(content)
        py_full, py_init = pinyin_of(title)

        return IndexEntry(
            path=rel_path,
            filename=abs_path.name,
            title=title,
            root=root,
            category=category,
            size=stat.st_size,
            mtime=stat.st_mtime,
            side=side,
            top_kind=top_kind,
            authors=authors,
            body=content,
            title_norm=title_norm,
            filename_norm=normalize(abs_path.name),
            path_norm=normalize(rel_path),
            body_norm=body_norm,
            title_pinyin=py_full,
            title_pyinit=py_init,
            body_charset=frozenset(body_norm),
        )

    def refresh(self, force: bool = False) -> int:
        """扫一遍磁盘，按 mtime 增量更新。短间隔内重复调用会被跳过。返回当前条目总数。"""
        with self._lock:
            now = time.time()
            if not force and self._entries and now - self._last_refresh < self.REFRESH_MIN_INTERVAL:
                return len(self._entries)
            seen: set[str] = set()
            for root in self.roots:
                base = self.repo_root / root
                if not base.exists():
                    continue
                for abs_path in base.rglob("*.txt"):
                    if not abs_path.is_file():
                        continue
                    rel = abs_path.relative_to(self.repo_root).as_posix()
                    seen.add(rel)
                    cached = self._entries.get(rel)
                    try:
                        mtime = abs_path.stat().st_mtime
                        size = abs_path.stat().st_size
                    except OSError:
                        continue
                    if cached and cached.mtime == mtime and cached.size == size:
                        continue
                    entry = self._make_entry(abs_path, rel)
                    if entry:
                        self._entries[rel] = entry
            for stale in set(self._entries.keys()) - seen:
                self._entries.pop(stale, None)
            self._last_refresh = now
            return len(self._entries)

    def invalidate(self, rel_path: str) -> None:
        with self._lock:
            self._entries.pop(rel_path, None)
            self._last_refresh = 0.0  # 写操作后下次 search 强制重扫

    def all_entries(self) -> list[IndexEntry]:
        with self._lock:
            return list(self._entries.values())

    def facets(self) -> dict:
        return self._build_facets(self.all_entries())

    # ---------- 查询 ----------

    def _score_entry(self, entry: IndexEntry, tokens_norm: list[str]) -> tuple[float, list[tuple[str, int, int]]]:
        """对一个文档按所有 token 评分。返回 (总分, [(token, snippet 起, 止)])。
        多 token 默认 AND；但当其它 token 已强命中（≥600）时容忍最多 1 个 token 完全没命中，
        以兼容用户输入英文别名/同义词 + 主中文名的混合查询。"""
        per_token: list[tuple[float, tuple[int, int] | None]] = []
        for tk in tokens_norm:
            best = 0.0
            snippet_pos: tuple[int, int] | None = None

            if tk == entry.title_norm:
                best = max(best, 1000)
            elif entry.title_norm.startswith(tk):
                best = max(best, 520)
            elif tk in entry.title_norm:
                best = max(best, 300)
            elif len(tk) >= 2:
                win = best_subseq_window(tk, entry.title_norm)
                # "强驱散" → "强制驱散" 紧凑度 0.75，可命中；散落字不会
                if win and win[2] >= 0.55:
                    best = max(best, 220 * win[2])

            if entry.title_pyinit and (tk == entry.title_pyinit or entry.title_pyinit.startswith(tk)):
                best = max(best, 240)
            if entry.title_pinyin and tk in entry.title_pinyin:
                best = max(best, 200)

            if tk in entry.filename_norm:
                best = max(best, 130)
            if tk in entry.path_norm:
                best = max(best, 90)

            for author in entry.authors:
                if tk in normalize(author):
                    best = max(best, 110)
                    break

            # 正文：先 charset 粗过滤，再连续/子序列
            if all(c in entry.body_charset for c in tk):
                if tk in entry.body_norm:
                    pos = entry.body_norm.find(tk)
                    snippet_pos = (pos, pos + len(tk))
                    best = max(best, 80 + min(20, 200 / max(1, pos / 100)))
                elif len(tk) >= 3:
                    # 仅长 token 才走正文子序列模糊。短 token 散落容易误命中。
                    win = best_subseq_window(tk, entry.body_norm)
                    if win and win[2] >= 0.7:
                        s, e, compact = win
                        snippet_pos = (s, e + 1)
                        best = max(best, 35 * compact)

            per_token.append((best, snippet_pos))

        # AND 语义：默认每个 token 都得 >0。但允许 1 个失配 token，前提是其它 token 至少一个 ≥250
        # （≥250 ≈ 标题包含命中，足以代表用户意图）。
        misses = sum(1 for sc, _ in per_token if sc <= 0)
        if misses > 0:
            strong = any(sc >= 250 for sc, _ in per_token)
            if not (misses == 1 and strong and len(per_token) >= 2):
                return (0.0, [])

        total = sum(sc for sc, _ in per_token)
        snippet_hints: list[tuple[str, int, int]] = [
            (tokens_norm[i], pos[0], pos[1]) for i, (_, pos) in enumerate(per_token) if pos
        ]
        if entry.size > 0:
            total += min(20.0, 800.0 / max(200, entry.size))
        return (total, snippet_hints)

    def _make_snippet(self, entry: IndexEntry, hints: list[tuple[str, int, int]], radius: int = 36) -> str:
        if not hints:
            head = entry.body.strip().splitlines()
            for line in head[:6]:
                line = line.strip()
                if line and not line.startswith(("（", "(")):
                    return line[: radius * 2]
            return ""
        # 取第一个 hint 的窗口（已是 body_norm 上的下标，body 可能与 body_norm 长度有差距，但近似可用）
        _, s, e = hints[0]
        body = entry.body
        if e > len(body):
            e = len(body)
        start = max(0, s - radius)
        end = min(len(body), e + radius)
        prefix = "…" if start > 0 else ""
        suffix = "…" if end < len(body) else ""
        return f"{prefix}{body[start:end].strip()}{suffix}".replace("\n", " ")

    def search(
        self,
        q: str,
        *,
        roots: Iterable[str] | None = None,
        category: str = "",
        kinds: Iterable[str] | None = None,
        sides: Iterable[str] | None = None,
        authors: Iterable[str] | None = None,
        limit: int = 200,
        offset: int = 0,
        include_content: bool = False,
    ) -> dict:
        self.refresh()
        tokens_raw = tokenize_query(q)
        tokens_norm = [normalize(t) for t in tokens_raw if t]
        # 拼音 token：用户输入英文也接，按拼音查
        tokens_norm = [t for t in tokens_norm if t]

        kinds_set = {k for k in (kinds or []) if k}
        sides_set = {s for s in (sides or []) if s}
        authors_set = {a for a in (authors or []) if a}
        roots_set = {r for r in (roots or []) if r}
        cat = category.strip("/")

        scored: list[tuple[float, IndexEntry, list[tuple[str, int, int]]]] = []
        all_matched: list[IndexEntry] = []

        for entry in self.all_entries():
            if roots_set and entry.root not in roots_set:
                continue
            if cat and not entry.category.startswith(cat):
                continue
            if kinds_set and entry.top_kind not in kinds_set:
                continue
            if sides_set and entry.side not in sides_set:
                continue
            if authors_set and not (set(entry.authors) & authors_set):
                continue

            if tokens_norm:
                score, hints = self._score_entry(entry, tokens_norm)
                if score <= 0:
                    continue
                scored.append((score, entry, hints))
            else:
                all_matched.append(entry)

        if tokens_norm:
            scored.sort(key=lambda x: (-x[0], x[1].path))
            matched_entries = [s[1] for s in scored]
        else:
            all_matched.sort(key=lambda e: e.path)
            matched_entries = all_matched

        total = len(matched_entries)
        page = matched_entries[offset : offset + limit] if limit > 0 else matched_entries

        items = []
        if tokens_norm:
            score_map = {id(s[1]): s for s in scored}
            for entry in page:
                s = score_map.get(id(entry))
                hints = s[2] if s else []
                row = entry.to_meta()
                row["score"] = round(s[0], 2) if s else 0
                row["snippet"] = self._make_snippet(entry, hints)
                if include_content:
                    row["content"] = entry.body
                items.append(row)
        else:
            for entry in page:
                row = entry.to_meta()
                row["score"] = 0
                row["snippet"] = ""
                if include_content:
                    row["content"] = entry.body
                items.append(row)

        # facets：在当前过滤集合上聚合
        facets = self._build_facets(matched_entries)
        return {
            "items": items,
            "count": total,
            "total": total,
            "limit": limit,
            "offset": offset,
            "tokens": tokens_raw,
            "facets": facets,
            "engine": {
                "pinyin": _PINYIN_OK,
                "opencc": _OPENCC_OK,
            },
        }

    def _build_facets(self, entries: list[IndexEntry]) -> dict:
        kinds: dict[str, int] = {}
        sides: dict[str, int] = {}
        authors: dict[str, int] = {}
        for e in entries:
            if e.top_kind:
                kinds[e.top_kind] = kinds.get(e.top_kind, 0) + 1
            if e.side:
                sides[e.side] = sides.get(e.side, 0) + 1
            for a in e.authors:
                authors[a] = authors.get(a, 0) + 1

        def to_list(d: dict[str, int]) -> list[dict]:
            return [{"name": k, "count": v} for k, v in sorted(d.items(), key=lambda kv: (-kv[1], kv[0]))]

        return {
            "kinds": to_list(kinds),
            "sides": to_list(sides),
            "authors": to_list(authors)[:80],  # 作者列表截断，避免长尾
        }

    def get(self, rel_path: str) -> IndexEntry | None:
        with self._lock:
            return self._entries.get(rel_path)

    def stats(self) -> dict:
        with self._lock:
            return {
                "count": len(self._entries),
                "pinyin": _PINYIN_OK,
                "opencc": _OPENCC_OK,
            }
