#!/usr/bin/env python3
"""Export the formal sequence library as native SealDice HelpDoc JSON."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import sys
import tarfile
import tempfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from text_encoding import read_text

GROUPS = ("公共建筑", "技能表", "能量池", "特质改造", "职业")
EXCLUDED = {"新过审序列", "荣誉室"}
MANIFEST = ".seqlib-helpdoc-manifest.json"
DIRECTORY = "【收录词条目录】"

# Align with lint_resources.py: paired brackets, including bare [名称].
CARD_STARTER = re.compile(
    r"^[ \t]*(?:\[[^\]\r\n]{0,12}名称\]|【[^】\r\n]{0,12}名称】|［[^］\r\n]{0,12}名称］|"
    r"〖[^〗\r\n]{0,12}名称〗|〔[^〕\r\n]{0,12}名称〕)[ \t]*[：:](.*)$"
)
CARD_NOTE = re.compile(r"^[ \t]*【\s*(?:能力修改|说明)\s*[：:].+】[ \t]*$")
SECTION = re.compile(r"^[ \t]*【([^】：:\r\n]+)】[ \t]*$")

# Match the search index's production roles and nested credit lists, but retain
# source spellings (the Web-specific alias dictionary is not a build dependency).
ROLES = "制作人|作者|原作者|审核人|修改人|调整人|重置人|复查人|策划|文本优化|制作者|投稿人|审核者|审核"
NON_CREATORS = {"审核人", "复查人", "审核者", "审核"}
CREDIT = re.compile(rf"({ROLES})\s*[：:]\s*(.*)")
PAREN_CREDIT = re.compile(rf"[（(]\s*({ROLES})\s*[：:]\s*(.*?)\s*[)）](?=\s*(?:[（(]\s*(?:{ROLES})\s*[：:]|$))")
NON_NAMES = {"未知", "佚名", "无名", "不知道", "不愿透露姓名", "?", "??", "？", "？？"}


@dataclass
class Resource:
    path: Path  # relative TXT path
    name: str
    author: str
    overview: str
    cards: list[tuple[str, str]]


def credit_names(text: str) -> list[str]:
    depth = start = 0
    names = []
    for i, char in enumerate(text):
        if char in "（([【":
            depth += 1
        elif char in "）)]】":
            depth = max(0, depth - 1)
        elif depth == 0 and char in "、,，;；":
            names.append(text[start:i])
            start = i + 1
    return names + [text[start:]]


def author_summary(header: str) -> str:
    authors = set()
    for raw in header.splitlines():
        line = raw.strip()
        matches = PAREN_CREDIT.findall(line)
        if not matches:
            bare = CREDIT.fullmatch(line)
            matches = [bare.groups()] if bare else []
        for role, names in matches:
            for name in credit_names(names):
                name = name.strip()
                inline = CREDIT.fullmatch(name)
                if inline:
                    role, name = inline.groups()
                    name = name.strip()
                if role not in NON_CREATORS and name and name not in NON_NAMES:
                    authors.add(name)
    return "、".join(sorted(authors)) or "宏观界域开发组"


def original_slice(lines: list[str], start: int, end: int) -> str:
    """Drop only outer blank lines and the last line delimiter, not indentation."""
    while start < end and not lines[start].strip():
        start += 1
    while start < end and not lines[end - 1].strip():
        end -= 1
    return "".join(lines[start:end]).rstrip("\r\n")


def parse_resource(path: Path, text: str) -> Resource:
    name = re.sub(r"^\d+】", "", path.stem).strip()
    if not name or not text.strip():
        raise ValueError(f"资源名或原文为空：{path}")
    lines = text.splitlines(keepends=True)
    card_starts = []
    sections = []
    for i, raw in enumerate(lines):
        line = raw.rstrip("\r\n")
        match = CARD_STARTER.match(line)
        if match:
            top = i
            while top and CARD_NOTE.fullmatch(lines[top - 1].rstrip("\r\n")):
                top -= 1
            card_starts.append((top, match[1].strip() or f"{name}·未命名条目"))
        elif SECTION.fullmatch(line):
            sections.append(i)
    if not card_starts:
        return Resource(path, name, author_summary(text), text.strip(), [])

    header = original_slice(lines, 0, card_starts[0][0])
    prose = [header] if header else []
    cards = []
    # A standalone section ends the preceding card even when it has its own
    # explanatory paragraphs. Keep that whole interval in the overview.
    boundaries = iter(sections)
    section = next(boundaries, len(lines))
    for i, (start, title) in enumerate(card_starts):
        following = card_starts[i + 1][0] if i + 1 < len(card_starts) else len(lines)
        while section <= start and section < len(lines):
            section = next(boundaries, len(lines))
        end = min(section, following)
        cards.append((title, original_slice(lines, start, end)))
        if end < following:
            prose.append(original_slice(lines, end, following))
    return Resource(path, name, author_summary(header), "\n\n".join(prose), cards)


def walk_error(error):
    raise error


def iter_resources(source: Path):
    for group in GROUPS:
        base = source / group
        if base.is_symlink() or not base.is_dir():
            raise ValueError(f"正式分类目录缺失或为符号链接：{base}")
        for parent, dirs, names in os.walk(base, onerror=walk_error):
            root = Path(parent)
            dirs[:] = sorted(d for d in dirs if d not in EXCLUDED and not d.startswith(".") and not (root / d).is_symlink())
            for filename in sorted(names):
                path = root / filename
                if not filename.startswith(".") and path.suffix.lower() == ".txt" and path.is_file() and not path.is_symlink():
                    text, encoding = read_text(path)
                    if encoding.endswith("-replace"):
                        raise ValueError(f"无法无损解码：{path}")
                    yield parse_resource(path.relative_to(source), text)


def title_sets(resources: list[Resource]) -> list[list[str]]:
    """Resolve all titles globally; generated suffixes never steal natural titles."""
    owners = defaultdict(set)
    originals = []
    for resource in resources:
        titles = [resource.name, *(title for title, _ in resource.cards)]
        originals.append(titles)
        for title in titles:
            owners[title].add(resource.path)

    plans = []
    for resource, titles in zip(resources, originals):
        counts = Counter()
        plan = []
        for i, title in enumerate(titles):
            candidate = f"{title}〔{resource.name}〕" if len(owners[title]) > 1 else title
            if i == 0 and title in titles[1:]:
                candidate += "（总览）"
            else:
                counts[title] += 1
                if counts[title] > 1:
                    candidate += f"（#{counts[title]}）"
            plan.append((title, candidate, len(owners[title]) == 1 and title == candidate))
        plans.append(plan)

    reserved = {title for plan in plans for title, _, natural in plan if natural}
    preferred = Counter(candidate for plan in plans for _, candidate, _ in plan)
    used = set(reserved)
    result = []
    for resource, plan in zip(resources, plans):
        assigned = []
        for title, candidate, natural in plan:
            if natural:
                assigned.append(title)
                continue
            if preferred[candidate] > 1 or candidate in used:
                base = f"{candidate}〔{resource.path.parent.as_posix()}〕"
                candidate = base
                sequence = 2
                while candidate in used or candidate in preferred:
                    candidate = f"{base}（#{sequence}）"
                    sequence += 1
            used.add(candidate)
            assigned.append(candidate)
        result.append(assigned)
    return result


def document(resource: Resource, titles: list[str], commit: str) -> dict:
    overview = resource.overview
    if resource.cards:
        directory = "\n".join(f"- {title}" for title in titles[1:])
        overview = "\n\n".join(part for part in (overview, f"{DIRECTORY}\n{directory}") if part)
    entries = dict(zip(titles, [overview, *(content for _, content in resource.cards)]))
    if len(entries) != len(resource.cards) + 1 or any(not title.strip() or not content.strip() for title, content in entries.items()):
        raise ValueError(f"词条覆盖或存在空词条：{resource.path}")
    return {
        "mod": resource.name,
        "author": resource.author,
        "brief": f"{resource.name}{resource.path.parts[0]}查询",
        "comment": f"宏观界域强化序列库自动生成 ({commit})",
        "helpdoc": entries,
    }


def json_bytes(data: dict) -> bytes:
    return (json.dumps(data, ensure_ascii=False, indent=2) + "\n").encode("utf-8")


def unique_object(items):
    obj = dict(items)
    if len(obj) != len(items):
        raise ValueError("JSON 包含重复 key")
    return obj


def build(source: Path, output: Path, commit: str | None = None, check: bool = False) -> dict:
    source, output = source.resolve(), output.resolve()
    if output == source or source in output.parents or output in source.parents:
        raise ValueError("输出目录与源目录不能重叠")
    if output.exists() and (not output.is_dir() or any(output.iterdir())):
        raise ValueError(f"输出目录必须不存在或为空：{output}")
    if commit is None:
        commit = subprocess.check_output(["git", "rev-parse", "--short", "HEAD"], cwd=ROOT, text=True).strip()
    if not commit.strip():
        raise ValueError("commit 不能为空")
    resources = list(iter_resources(source))
    titles = title_sets(resources)
    groups = {group: {"files": 0, "entries": 0} for group in GROUPS}
    files = {}
    output.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".seqlib-build-", dir=output.parent) as tmp:
        staging = Path(tmp) / "export"
        staging.mkdir()
        for group in GROUPS:
            (staging / group).mkdir()
        for resource, names in zip(resources, titles):
            rel = resource.path.with_suffix(".json")
            target = staging / rel
            if target.exists():
                raise ValueError(f"资源输出路径冲突：{rel}")
            doc = document(resource, names, commit)
            data = json_bytes(doc)
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(data)
            if check and json.loads(target.read_bytes(), object_pairs_hook=unique_object) != doc:
                raise ValueError(f"生成文件与原始切片不一致：{rel}")
            count = len(doc["helpdoc"])
            files[rel.as_posix()] = {"entries": count, "sha256": hashlib.sha256(data).hexdigest()}
            groups[rel.parts[0]]["files"] += 1
            groups[rel.parts[0]]["entries"] += count
        entry_count = sum(group["entries"] for group in groups.values())
        if check:
            actual = {p.relative_to(staging).as_posix() for p in staging.rglob("*.json")}
            if actual != set(files) or len(files) != len(resources) or not all(g["files"] for g in groups.values()):
                raise ValueError("资源或分类未完整导出")
            if len(files) < 650 or entry_count < 5000:
                raise ValueError(f"低于门禁要求：{len(files)} 文件 / {entry_count} 词条（至少 650 / 5000）")
        manifest = {"schema_version": 1, "commit": commit, "groups": groups, "files": files,
                    "file_count": len(files), "entry_count": entry_count}
        (staging / MANIFEST).write_bytes(json_bytes(manifest))
        staging.replace(output)
    return manifest


def archive_export(output: Path, archive: Path) -> None:
    archive.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory(prefix=".seqlib-archive-", dir=archive.parent) as tmp:
        temporary_archive = Path(tmp) / "bundle.tar.gz"
        with tarfile.open(temporary_archive, "w:gz") as bundle:
            for name in (*GROUPS, MANIFEST):
                bundle.add(output / name, arcname=name)
        temporary_archive.replace(archive)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-dir", type=Path, default=ROOT / "序列库")
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--commit")
    parser.add_argument("--archive", type=Path)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    output, source = args.output_dir.resolve(), args.source_dir.resolve()
    archive = args.archive.resolve() if args.archive else None
    if archive and (not archive.name.endswith(".tar.gz") or archive == output or output in archive.parents
                    or archive == source or source in archive.parents):
        parser.error("--archive 必须是位于源目录及输出目录之外的 .tar.gz 文件")
    try:
        manifest = build(source, output, args.commit, args.check)
        if archive:
            archive_export(output, archive)
    except (OSError, ValueError, subprocess.CalledProcessError, tarfile.TarError) as exc:
        parser.exit(1, f"构建失败：{exc}\n")
    print(f"### SealDice HelpDoc · `{manifest['commit']}`\n")
    print("| 分组 | 文件 | 词条 |\n| --- | ---: | ---: |")
    for group, counts in manifest["groups"].items():
        print(f"| {group} | {counts['files']} | {counts['entries']} |")
    print(f"| **合计** | **{manifest['file_count']}** | **{manifest['entry_count']}** |")
    return 0


if __name__ == "__main__":
    sys.exit(main())
