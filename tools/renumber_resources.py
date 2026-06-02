#!/usr/bin/env python3
"""
资源序号重排工具。

默认只输出 dry run 计划；只有传入 --apply 才会实际改名和替换 txt 首行标题。
"""

from __future__ import annotations

import argparse
import re
import uuid
from dataclasses import dataclass
from pathlib import Path


NUMBERED_NAME_RE = re.compile(r"^(?P<number>\d+)】(?P<title>.+)$")
RESOURCE_EXTENSIONS = {".txt", ".html", ".htm", ".docx", ".doc", ".xlsx"}
TEXT_EXTENSIONS = {".txt"}


@dataclass(frozen=True)
class RenameItem:
    old_path: Path
    new_path: Path


@dataclass(frozen=True)
class TitleUpdate:
    path: Path
    title: str


@dataclass(frozen=True)
class ResourceUpdatePlan:
    repo_root: Path
    roots: tuple[Path, ...]
    renames: tuple[RenameItem, ...]
    title_updates: tuple[TitleUpdate, ...]


def numbered_sort_key(path: Path) -> tuple[int, str, str]:
    match = NUMBERED_NAME_RE.match(path.stem if path.is_file() else path.name)
    if match:
        return (int(match.group("number")), match.group("title"), path.name)
    return (10**9, path.name, path.name)


def numbered_title(path: Path) -> str | None:
    name = path.stem if path.is_file() else path.name
    match = NUMBERED_NAME_RE.match(name)
    if not match:
        return None
    return match.group("title")


def is_resource_file(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in RESOURCE_EXTENSIONS


def is_numbered_resource(path: Path) -> bool:
    if path.is_dir():
        return NUMBERED_NAME_RE.match(path.name) is not None
    return is_resource_file(path) and numbered_title(path) is not None


def resolve_roots(repo_root: Path, roots: list[str | Path]) -> tuple[Path, ...]:
    resolved = []
    for root in roots:
        path = Path(root)
        if not path.is_absolute():
            path = repo_root / path
        resolved.append(path.resolve())
    return tuple(resolved)


def iter_directories(root: Path) -> list[Path]:
    directories = [root]
    directories.extend(path for path in root.rglob("*") if path.is_dir())
    return directories


def final_path_for(path: Path, renames: list[RenameItem]) -> Path:
    result = path
    for item in sorted(renames, key=lambda rename: len(rename.old_path.parts), reverse=True):
        if result == item.old_path:
            result = item.new_path
        else:
            try:
                rel = result.relative_to(item.old_path)
            except ValueError:
                continue
            result = item.new_path / rel
    return result


def read_text_file(path: Path) -> str:
    for encoding in ("utf-8", "utf-8-sig", "gbk", "gb2312", "big5"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def text_first_line(text: str) -> str:
    if not text:
        return ""
    return text.splitlines()[0] if text.splitlines() else ""


def plan_resource_updates(
    repo_root: Path,
    roots: list[str | Path],
    *,
    include_titles: bool = True,
) -> ResourceUpdatePlan:
    repo_root = repo_root.resolve()
    resolved_roots = resolve_roots(repo_root, roots)
    renames: list[RenameItem] = []

    for root in resolved_roots:
        if not root.exists():
            raise FileNotFoundError(f"资源根目录不存在: {root}")

        for directory in iter_directories(root):
            children = [child for child in directory.iterdir() if is_numbered_resource(child)]
            children.sort(key=numbered_sort_key)

            for index, child in enumerate(children, start=1):
                title = numbered_title(child)
                if title is None:
                    continue

                suffix = child.suffix if child.is_file() else ""
                new_name = f"{index:03d}】{title}{suffix}"
                new_path = child.with_name(new_name)
                if new_path != child:
                    renames.append(RenameItem(child, new_path))

    title_updates: list[TitleUpdate] = []
    if include_titles:
        for root in resolved_roots:
            for path in root.rglob("*"):
                if not path.is_file() or path.suffix.lower() not in TEXT_EXTENSIONS:
                    continue
                final_path = final_path_for(path, renames)
                title = final_path.stem
                current_title = text_first_line(read_text_file(path))
                if current_title != title:
                    title_updates.append(TitleUpdate(final_path, title))

    return ResourceUpdatePlan(
        repo_root=repo_root,
        roots=resolved_roots,
        renames=tuple(renames),
        title_updates=tuple(title_updates),
    )


def validate_plan(plan: ResourceUpdatePlan) -> None:
    by_parent: dict[Path, list[RenameItem]] = {}
    for item in plan.renames:
        by_parent.setdefault(item.old_path.parent, []).append(item)

    for parent, items in by_parent.items():
        final_names = [item.new_path.name for item in items]
        if len(final_names) != len(set(final_names)):
            raise RuntimeError(f"重排后存在重名: {parent}")

        moving_old_names = {item.old_path.name for item in items}
        for item in items:
            if item.new_path.exists() and item.new_path.name not in moving_old_names:
                raise RuntimeError(f"目标路径已存在，无法覆盖: {item.new_path}")


def replace_first_line(path: Path, title: str) -> None:
    text = read_text_file(path)
    if text == "":
        path.write_text(f"{title}\n", encoding="utf-8")
        return

    lines = text.splitlines(keepends=True)
    if not lines:
        path.write_text(f"{title}\n", encoding="utf-8")
        return

    newline = "\n"
    if lines[0].endswith("\r\n"):
        newline = "\r\n"
    elif lines[0].endswith("\n"):
        newline = "\n"
    elif lines[0].endswith("\r"):
        newline = "\r"
    else:
        newline = ""

    lines[0] = f"{title}{newline}"
    path.write_text("".join(lines), encoding="utf-8")


def apply_renames(renames: tuple[RenameItem, ...]) -> None:
    by_parent: dict[Path, list[RenameItem]] = {}
    for item in renames:
        by_parent.setdefault(item.old_path.parent, []).append(item)

    for parent in sorted(by_parent, key=lambda path: len(path.parts), reverse=True):
        items = by_parent[parent]
        temp_pairs: list[tuple[Path, Path]] = []

        for item in items:
            temp_path = item.old_path.with_name(f".renumber-tmp-{uuid.uuid4().hex}-{item.old_path.name}")
            item.old_path.rename(temp_path)
            temp_pairs.append((temp_path, item.new_path))

        for temp_path, final_path in temp_pairs:
            temp_path.rename(final_path)


def apply_resource_updates(plan: ResourceUpdatePlan) -> None:
    validate_plan(plan)
    apply_renames(plan.renames)
    for update in plan.title_updates:
        replace_first_line(update.path, update.title)


def format_path(path: Path, repo_root: Path) -> str:
    try:
        return str(path.relative_to(repo_root))
    except ValueError:
        return str(path)


def print_plan(plan: ResourceUpdatePlan) -> None:
    print("资源重排计划")
    print(f"根目录: {', '.join(format_path(root, plan.repo_root) for root in plan.roots)}")
    print(f"改名: {len(plan.renames)} 项")
    for item in plan.renames:
        print(f"  R {format_path(item.old_path, plan.repo_root)} -> {format_path(item.new_path, plan.repo_root)}")

    print(f"首行标题更新: {len(plan.title_updates)} 项")
    for item in plan.title_updates:
        print(f"  T {format_path(item.path, plan.repo_root)} => {item.title}")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="序列库/荣誉室资源序号重排工具")
    parser.add_argument(
        "--root",
        action="append",
        dest="roots",
        help="资源根目录，可重复传入。默认处理 序列库 和 荣誉室。",
    )
    parser.add_argument("--repo-root", default=".", help="仓库根目录，默认当前目录。")
    parser.add_argument("--apply", action="store_true", help="实际执行改名和首行标题更新。")
    parser.add_argument("--skip-title-update", action="store_true", help="只重排序号，不替换 txt 首行。")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    repo_root = Path(args.repo_root).resolve()
    roots = args.roots or ["序列库", "荣誉室"]
    plan = plan_resource_updates(repo_root, roots, include_titles=not args.skip_title_update)
    validate_plan(plan)
    print_plan(plan)

    if not args.apply:
        print("\n当前是 dry run，没有修改文件。确认无误后加 --apply 执行。")
        return 0

    apply_resource_updates(plan)
    print("\n已执行重排和首行标题更新。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
