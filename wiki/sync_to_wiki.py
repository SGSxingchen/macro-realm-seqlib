#!/usr/bin/env python3
"""
序列库 Fandom Wiki 同步脚本

将序列库/荣誉室的 txt/docx 文件自动同步到 Fandom Wiki。
每次同步采用全量清理 + 重建策略，避免文件名修改导致残留页面。

用法:
    # 同步全部（序列库 + 荣誉室）
    python sync_to_wiki.py --user BotName --password xxx

    # 增量同步：只同步最近一次提交的改动
    python sync_to_wiki.py --user BotName --password xxx --incremental

    # 增量同步：同步从指定提交以来的改动
    python sync_to_wiki.py --user BotName --password xxx --diff-from v6.3.1

    # 增量同步：同步工作区未提交的修改
    python sync_to_wiki.py --user BotName --password xxx --diff-unstaged

    # 只同步序列库
    python sync_to_wiki.py --user BotName --password xxx --skip-honor

    # 试运行（不推送，只在本地预览）
    python sync_to_wiki.py --user BotName --password xxx --dry-run

    # 只同步某个子目录
    python sync_to_wiki.py --user BotName --password xxx --filter 职业/战技侧

依赖:
    pip install mwclient
    pandoc (用于 docx 转换，已有)
"""

import argparse
import re
import subprocess
import sys
import time
from pathlib import Path

try:
    import mwclient
except ImportError:
    print("错误: 请先安装 mwclient")
    print("  pip install mwclient")
    sys.exit(1)

# ============================================================
# 配置
# ============================================================

WIKI_SITE = "macro-realm.fandom.com"
WIKI_PATH = "/zh/"

# 同步的内容目录
CONTENT_DIRS = ["序列库"]
HONOR_DIR = "荣誉室"

# 支持的文件类型
SUPPORTED_EXTENSIONS = {".txt", ".html", ".htm", ".docx", ".doc"}

# 自动同步管理的分类（清理时只删除这些分类下的页面）
# 手动创建的页面（发展历史、管理组等）不要加入这些分类，就不会被误删
AUTO_SYNC_CATEGORIES = [
    "职业", "战技侧", "神秘侧", "科技侧", "特殊侧",
    "技能表", "能量池", "公共建筑",
    "特质改造", "生化改造类", "特化改造类", "特殊特质", "异化改造类",
    "荣誉室",
]

# API 请求间隔（秒），避免被限流
# Fandom 对非 Bot 账号限流较严，建议 5 秒以上
REQUEST_DELAY = 5

# 被限流时的重试配置
RATE_LIMIT_WAIT = 30   # 被限流后等待秒数
RATE_LIMIT_RETRIES = 3  # 最大重试次数


# ============================================================
# 工具函数
# ============================================================


def read_text_file(path: Path) -> str:
    """读取文本文件，自动检测编码（与 build_chm.py 逻辑一致）"""
    for enc in ("utf-8", "utf-8-sig", "gbk", "gb2312", "big5"):
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def sort_key(name: str):
    """排序键：按前导编号排序，无编号的排后面"""
    match = re.match(r"(\d+)", name)
    if match:
        return (0, int(match.group(1)), name)
    return (1, 0, name)


def strip_number_prefix(name: str) -> str:
    """
    去掉编号前缀：'001】百夫长《荣耀战魂》' → '百夫长《荣耀战魂》'
    如果没有编号前缀，返回原名。
    """
    match = re.match(r"\d+】(.+)", name)
    if match:
        return match.group(1)
    return name


def sanitize_page_name(name: str) -> str:
    """
    清理 MediaWiki 不允许的页面名字符。
    非法字符: # < > [ ] | { }
    替换为对应的全角字符，保持可读性。
    """
    replacements = {
        "[": "【", "]": "】",
        "{": "（", "}": "）",
        "#": "＃", "<": "＜", ">": "＞", "|": "｜",
    }
    for old, new in replacements.items():
        name = name.replace(old, new)
    return name.strip()


# ============================================================
# Git 增量检测
# ============================================================


def get_git_root() -> Path:
    """获取 Git 仓库根目录"""
    result = subprocess.run(
        ["git", "rev-parse", "--show-toplevel"],
        capture_output=True, text=True, check=True, encoding='utf-8'
    )
    return Path(result.stdout.strip())


def get_git_changes(diff_from: str = None, unstaged: bool = False) -> dict:
    """
    通过 Git diff 获取文件变更。

    Args:
        diff_from: 起始提交（如 'HEAD~1', 'v6.3.1'），None 则为最近一次提交
        unstaged: 是否检测未暂存的工作区修改

    Returns:
        {
            'added': [Path, ...],      # 新增文件
            'modified': [Path, ...],   # 修改文件
            'deleted': [Path, ...],    # 删除文件
            'renamed': [(old_path, new_path), ...]  # 重命名文件
        }
    """
    changes = {
        'added': [],
        'modified': [],
        'deleted': [],
        'renamed': []
    }

    try:
        # 构建 git diff 命令
        # 使用 -z 选项输出 NUL 分隔的路径，避免路径转义问题
        # 使用 -c core.quotepath=false 禁用路径引号转义
        if unstaged:
            # 工作区改动（未暂存）
            cmd = ["git", "-c", "core.quotepath=false", "diff", "--name-status", "--diff-filter=AMDR"]
        elif diff_from:
            # 从指定提交到 HEAD
            cmd = ["git", "-c", "core.quotepath=false", "diff", "--name-status", "--diff-filter=AMDR", diff_from, "HEAD"]
        else:
            # 最近一次提交
            cmd = ["git", "-c", "core.quotepath=false", "diff", "--name-status", "--diff-filter=AMDR", "HEAD~1", "HEAD"]

        # Git for Windows 默认输出 UTF-8（与系统 ACP/cp936 不同）
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            encoding='utf-8',
            errors='replace'
        )

        for line in result.stdout.strip().split('\n'):
            if not line:
                continue

            parts = line.split('\t')
            status = parts[0][0]  # A/M/D/R

            if status == 'A':
                # 新增
                file_path = Path(parts[1])
                if _is_supported_file(file_path):
                    changes['added'].append(file_path)

            elif status == 'M':
                # 修改
                file_path = Path(parts[1])
                if _is_supported_file(file_path):
                    changes['modified'].append(file_path)

            elif status == 'D':
                # 删除
                file_path = Path(parts[1])
                if _is_supported_file(file_path):
                    changes['deleted'].append(file_path)

            elif status == 'R':
                # 重命名 (格式: R<similarity>\told_path\tnew_path)
                # 处理可能的相似度数字 (如 R100)
                old_path = Path(parts[1])
                new_path = Path(parts[2])
                if _is_supported_file(old_path) or _is_supported_file(new_path):
                    changes['renamed'].append((old_path, new_path))

    except subprocess.CalledProcessError as e:
        print(f"  [!] Git diff 失败: {e}")
        print(f"  [!] stderr: {e.stderr}")
    except Exception as e:
        print(f"  [!] Git diff 执行错误: {e}")

    return changes


def _is_supported_file(path: Path) -> bool:
    """检查文件是否在支持的目录和扩展名范围内"""
    parts = path.parts
    if not parts:
        return False

    # 检查是否在序列库或荣誉室目录
    top_dir = parts[0]
    if top_dir not in CONTENT_DIRS and top_dir != HONOR_DIR:
        return False

    # 检查扩展名
    return path.suffix.lower() in SUPPORTED_EXTENSIONS


# ============================================================
# 文件扫描
# ============================================================


def scan_files(source_dir: Path, content_dirs: list, include_honor: bool = True) -> dict:
    """
    扫描源目录，返回 {相对路径: 绝对路径}。
    相对路径相对于 source_dir。
    """
    files = {}

    all_dirs = list(content_dirs)
    if include_honor:
        all_dirs.append(HONOR_DIR)

    for dir_name in all_dirs:
        content_dir = source_dir / dir_name
        if not content_dir.is_dir():
            print(f"  [!] 目录不存在，跳过: {dir_name}")
            continue
        for f in sorted(content_dir.rglob("*"), key=lambda p: sort_key(p.name)):
            if f.is_file() and f.suffix.lower() in SUPPORTED_EXTENSIONS:
                files[f.relative_to(source_dir)] = f

    return files


# ============================================================
# 内容转换
# ============================================================


def txt_to_wikitext(file_path: Path) -> tuple:
    """
    将 txt 文件转为 WikiText（最小转换策略）。
    返回 (页面名, WikiText内容)。

    策略：
    - 页面名始终用文件名（去掉编号前缀），因为文件格式不统一
    - 文件全部内容用 <pre> 包裹，原样保留
    """
    content = read_text_file(file_path)

    # 页面名用文件名，去掉编号前缀
    page_name = strip_number_prefix(file_path.stem)

    # 全部内容用 <pre> 包裹
    body = content.rstrip()

    if body:
        wiki_text = f"<pre>\n{body}\n</pre>"
    else:
        wiki_text = ""

    return page_name, wiki_text


def docx_to_wikitext(file_path: Path) -> tuple:
    """
    将 docx 文件通过 pandoc 转为 WikiText。
    返回 (页面名, WikiText内容)。
    """
    page_name = strip_number_prefix(file_path.stem)

    try:
        result = subprocess.run(
            ["pandoc", str(file_path), "-t", "mediawiki", "-o", "-"],
            check=True, capture_output=True, timeout=60,
        )
        wiki_text = result.stdout.decode("utf-8", errors="replace").rstrip()
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"  [!] pandoc 转换失败: {file_path.name} ({e})")
        # 回退：读取原始内容用 <pre> 包裹
        try:
            raw = read_text_file(file_path)
            wiki_text = f"<pre>\n{raw.rstrip()}\n</pre>"
        except Exception:
            wiki_text = f"''此文件原为 {file_path.suffix} 格式，自动转换失败。''"

    return page_name, wiki_text


def html_to_wikitext(file_path: Path) -> tuple:
    """
    将 html 文件转为 WikiText。
    尝试用 pandoc 转换，失败则原样包裹。
    返回 (页面名, WikiText内容)。
    """
    page_name = strip_number_prefix(file_path.stem)

    try:
        result = subprocess.run(
            ["pandoc", str(file_path), "-f", "html", "-t", "mediawiki", "-o", "-"],
            check=True, capture_output=True, timeout=60,
        )
        wiki_text = result.stdout.decode("utf-8", errors="replace").rstrip()
    except Exception as e:
        print(f"  [!] html 转换失败: {file_path.name} ({e})")
        raw = read_text_file(file_path)
        wiki_text = raw.rstrip()

    return page_name, wiki_text


def convert_file(file_path: Path) -> tuple:
    """
    根据文件类型选择转换方法。
    返回 (页面名, WikiText内容)。
    """
    ext = file_path.suffix.lower()
    if ext == ".txt":
        return txt_to_wikitext(file_path)
    elif ext in (".docx", ".doc"):
        return docx_to_wikitext(file_path)
    elif ext in (".html", ".htm"):
        return html_to_wikitext(file_path)
    else:
        return file_path.stem, ""


# ============================================================
# 分类映射
# ============================================================


def get_categories(rel_path: Path) -> list:
    """
    根据文件的相对路径生成分类列表。
    跳过顶层目录名（序列库/荣誉室），取中间的目录层级作为分类。

    例：
    序列库/职业/战技侧/xxx.txt → ['职业', '战技侧']
    序列库/能量池/xxx.txt → ['能量池']
    荣誉室/职业/xxx.txt → ['荣誉室', '职业']
    """
    parts = rel_path.parent.parts  # 如 ('序列库', '职业', '战技侧')
    if not parts:
        return []

    categories = []
    top = parts[0]

    # 荣誉室的页面加上荣誉室分类
    if top == HONOR_DIR:
        categories.append("荣誉室")
        # 取荣誉室下的子目录作为分类
        for part in parts[1:]:
            categories.append(part)
    else:
        # 序列库：跳过顶层，取子目录
        for part in parts[1:]:
            categories.append(part)

    return categories


def is_honor_hall(rel_path: Path) -> bool:
    """判断文件是否属于荣誉室"""
    parts = rel_path.parts
    return len(parts) > 0 and parts[0] == HONOR_DIR


def build_category_tags(categories: list) -> str:
    """将分类列表转为 WikiText 分类标签"""
    if not categories:
        return ""
    return "\n" + "\n".join(f"[[分类:{cat}]]" for cat in categories)


# ============================================================
# Wiki 同步
# ============================================================


def delete_wiki_page(site, page_name: str, delay: float, dry_run: bool = False) -> bool:
    """
    删除单个 Wiki 页面。

    Args:
        site: mwclient.Site 实例
        page_name: 页面名称
        delay: 删除后的等待时间
        dry_run: 是否试运行

    Returns:
        是否成功删除
    """
    if dry_run:
        print(f"  [dry-run] 将删除页面: {page_name}")
        return True

    try:
        page = site.pages[page_name]
        if page.exists:
            page.delete(reason="自动同步：源文件已删除")
            print(f"  已删除页面: {page_name}")
            time.sleep(delay)
            return True
        else:
            print(f"  [!] 页面不存在，跳过删除: {page_name}")
            return False
    except Exception as e:
        print(f"  [!] 删除失败: {page_name} ({e})")
        return False


def cleanup_wiki_pages(site, categories: list, delay: float, dry_run: bool = False):
    """
    清理自动同步分类下的所有页面。
    只删除属于 AUTO_SYNC_CATEGORIES 中分类的页面。
    """
    deleted = set()  # 避免重复删除（一个页面可能属于多个分类）

    for cat_name in categories:
        print(f"  清理分类 [{cat_name}] ...")
        try:
            cat = site.categories[cat_name]
            for page in cat:
                if page.name in deleted:
                    continue
                if dry_run:
                    print(f"    [dry-run] 将删除: {page.name}")
                else:
                    try:
                        page.delete(reason="自动同步：全量清理重建")
                        print(f"    已删除: {page.name}")
                    except Exception as e:
                        print(f"    [!] 删除失败: {page.name} ({e})")
                deleted.add(page.name)
                time.sleep(delay)
        except Exception as e:
            print(f"  [!] 获取分类 [{cat_name}] 失败: {e}")

    print(f"  清理完成，共处理 {len(deleted)} 个页面\n")


def sync_files_to_wiki(
    source_dir: Path,
    files: dict,
    site,
    delay: float,
    dry_run: bool = False,
    filter_path: str = None,
):
    """
    将文件同步到 Wiki。

    Args:
        source_dir: 源文件根目录
        files: {相对路径: 绝对路径} 映射
        site: mwclient.Site 实例
        delay: 请求间隔（秒）
        dry_run: 是否试运行
        filter_path: 只同步包含此路径的文件
    """
    total = 0
    created = 0
    skipped = 0
    errors = 0

    for rel_path, abs_path in files.items():
        # 过滤
        if filter_path:
            rel_str = str(rel_path).replace("\\", "/")
            if filter_path not in rel_str:
                continue

        total += 1
        ext = abs_path.suffix.lower()
        ext_label = {
            ".txt": "txt", ".docx": "doc", ".doc": "doc",
            ".html": "htm", ".htm": "htm",
        }.get(ext, ext)

        # 转换
        try:
            page_name, wiki_text = convert_file(abs_path)
            page_name = sanitize_page_name(page_name)
        except Exception as e:
            print(f"  [{ext_label:>3}] [!] 转换失败: {rel_path} ({e})")
            errors += 1
            continue

        # 生成分类标签
        categories = get_categories(rel_path)
        category_tags = build_category_tags(categories)

        # 荣誉室页面加已下架标记
        if is_honor_hall(rel_path):
            honor_notice = "''本内容已从序列库下架，仅供历史参考。''\n----\n"
            wiki_text = honor_notice + wiki_text

        # 拼接最终内容
        full_text = wiki_text + category_tags

        if dry_run:
            print(f"  [{ext_label:>3}] {rel_path}")
            print(f"        → 页面名: {page_name}")
            print(f"        → 分类: {categories}")
            if total <= 3:
                # 前 3 个文件显示完整内容预览
                preview = full_text[:300]
                if len(full_text) > 300:
                    preview += f"\n... (共 {len(full_text)} 字符)"
                print(f"        → 内容预览:\n{preview}\n")
            created += 1
        else:
            success = False
            for attempt in range(RATE_LIMIT_RETRIES + 1):
                try:
                    page = site.pages[page_name]
                    page.save(full_text, summary="自动同步更新")
                    print(f"  [{ext_label:>3}] 已创建: {page_name} ({len(full_text)} 字符)")
                    created += 1
                    success = True
                    break
                except mwclient.errors.APIError as e:
                    if "ratelimited" in str(e) and attempt < RATE_LIMIT_RETRIES:
                        print(f"  [{ext_label:>3}] 被限流，等待 {RATE_LIMIT_WAIT} 秒后重试 ({attempt + 1}/{RATE_LIMIT_RETRIES})...")
                        time.sleep(RATE_LIMIT_WAIT)
                    else:
                        print(f"  [{ext_label:>3}] [!] 上传失败: {page_name} ({e})")
                        errors += 1
                        break
                except Exception as e:
                    print(f"  [{ext_label:>3}] [!] 上传失败: {page_name} ({e})")
                    errors += 1
                    break

            time.sleep(delay)

    print(f"\n  同步完成: 共 {total} 个文件，成功 {created}，失败 {errors}")


def incremental_sync(
    source_dir: Path,
    changes: dict,
    site,
    delay: float,
    dry_run: bool = False
):
    """
    增量同步：根据 Git diff 结果处理文件变更。

    Args:
        source_dir: 源文件根目录
        changes: get_git_changes() 返回的变更字典
        site: mwclient.Site 实例
        delay: 请求间隔（秒）
        dry_run: 是否试运行
    """
    total_added = len(changes['added'])
    total_modified = len(changes['modified'])
    total_deleted = len(changes['deleted'])
    total_renamed = len(changes['renamed'])

    print(f"  检测到变更:")
    print(f"    新增: {total_added} 个文件")
    print(f"    修改: {total_modified} 个文件")
    print(f"    删除: {total_deleted} 个文件")
    print(f"    重命名: {total_renamed} 个文件")
    print()

    success_count = 0
    error_count = 0

    # --- 处理删除 ---
    if changes['deleted']:
        print(f"[1/4] 处理删除的文件 ({total_deleted} 个)...")
        for rel_path in changes['deleted']:
            page_name = strip_number_prefix(rel_path.stem)
            page_name = sanitize_page_name(page_name)

            if delete_wiki_page(site, page_name, delay, dry_run):
                success_count += 1
            else:
                error_count += 1
        print()

    # --- 处理重命名 (删除旧页面) ---
    if changes['renamed']:
        print(f"[2/4] 处理重命名的文件 - 删除旧页面 ({total_renamed} 个)...")
        for old_path, new_path in changes['renamed']:
            old_page_name = strip_number_prefix(old_path.stem)
            old_page_name = sanitize_page_name(old_page_name)

            if delete_wiki_page(site, old_page_name, delay, dry_run):
                success_count += 1
            else:
                error_count += 1
        print()

    # --- 处理新增和修改 ---
    # Git diff 路径相对于仓库根目录，需要用 git root 来解析绝对路径
    git_root = get_git_root()
    to_sync = {}

    # 新增文件
    for rel_path in changes['added']:
        abs_path = git_root / rel_path
        if abs_path.exists():
            to_sync[rel_path] = abs_path

    # 修改文件
    for rel_path in changes['modified']:
        abs_path = git_root / rel_path
        if abs_path.exists():
            to_sync[rel_path] = abs_path

    # 重命名文件的新路径
    for old_path, new_path in changes['renamed']:
        abs_path = git_root / new_path
        if abs_path.exists():
            to_sync[new_path] = abs_path

    if to_sync:
        total_sync = len(to_sync)
        print(f"[3/4] 创建/更新页面 ({total_sync} 个)...")
        sync_files_to_wiki(
            source_dir=git_root,
            files=to_sync,
            site=site,
            delay=delay,
            dry_run=dry_run,
            filter_path=None
        )

    print(f"[4/4] 增量同步完成")
    print(f"  总计处理: {total_added + total_modified + total_deleted + total_renamed} 个变更")


# ============================================================
# 主流程
# ============================================================


def main():
    parser = argparse.ArgumentParser(
        description="序列库 Fandom Wiki 同步工具",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
示例:
  python sync_to_wiki.py --user MyBot --password xxx
  python sync_to_wiki.py --user MyBot --password xxx --dry-run
  python sync_to_wiki.py --user MyBot --password xxx --filter 公共建筑
  python sync_to_wiki.py --user MyBot --password xxx --skip-honor
""",
    )
    parser.add_argument("--user", required=True, help="Wiki 用户名")
    parser.add_argument("--password", required=True, help="Wiki 密码")
    parser.add_argument("--source-dir", default=".", help="源文件目录 (默认当前目录)")
    parser.add_argument("--dry-run", action="store_true", help="试运行，不实际推送")
    parser.add_argument("--skip-honor", action="store_true", help="跳过荣誉室")
    parser.add_argument("--skip-cleanup", action="store_true", help="跳过清理步骤（只创建/更新）")
    parser.add_argument("--filter", dest="filter_path", help="只同步包含此路径的文件 (如: 职业/战技侧)")
    parser.add_argument("--delay", type=float, default=REQUEST_DELAY, help=f"API 请求间隔秒数 (默认 {REQUEST_DELAY})")
    parser.add_argument("--wiki-site", default=WIKI_SITE, help=f"Wiki 站点 (默认 {WIKI_SITE})")

    # 增量同步选项
    incremental_group = parser.add_argument_group("增量同步选项")
    incremental_group.add_argument("--incremental", action="store_true",
                                   help="增量同步：只同步最近一次提交的改动")
    incremental_group.add_argument("--diff-from", metavar="COMMIT",
                                   help="增量同步：同步从指定提交/tag 到 HEAD 的改动 (如: v6.3.1, HEAD~5)")
    incremental_group.add_argument("--diff-unstaged", action="store_true",
                                   help="增量同步：同步工作区未暂存的修改")

    args = parser.parse_args()

    source_dir = Path(args.source_dir).resolve()
    include_honor = not args.skip_honor

    # 检测增量同步模式
    is_incremental = args.incremental or args.diff_from or args.diff_unstaged
    mode_desc = "增量同步" if is_incremental else "全量同步"
    if args.incremental:
        mode_desc += " (最近一次提交)"
    elif args.diff_from:
        mode_desc += f" (从 {args.diff_from})"
    elif args.diff_unstaged:
        mode_desc += " (工作区未暂存)"

    print("=" * 50)
    print("  序列库 Wiki 同步工具")
    print(f"  Wiki:     {args.wiki_site}")
    print(f"  源目录:   {source_dir}")
    print(f"  用户:     {args.user}")
    print(f"  荣誉室:   {'包含' if include_honor else '跳过'}")
    print(f"  过滤:     {args.filter_path or '无'}")
    print(f"  模式:     {mode_desc}")
    if args.dry_run:
        print(f"  试运行:   是")
    print("=" * 50)
    print()

    # --- 增量同步模式 ---
    if is_incremental:
        print("[1/3] 检测文件变更 (Git diff)...")

        # 获取变更
        if args.diff_unstaged:
            changes = get_git_changes(diff_from=None, unstaged=True)
        elif args.diff_from:
            changes = get_git_changes(diff_from=args.diff_from, unstaged=False)
        else:
            changes = get_git_changes(diff_from=None, unstaged=False)

        # 统计变更
        total_changes = (len(changes['added']) + len(changes['modified']) +
                        len(changes['deleted']) + len(changes['renamed']))

        if total_changes == 0:
            print("      没有检测到文件变更，退出。")
            return

        print(f"      检测到 {total_changes} 个变更\n")

        # 连接 Wiki
        site = None
        if not args.dry_run:
            print("[2/3] 连接 Wiki...")
            try:
                site = mwclient.Site(args.wiki_site, path=WIKI_PATH)
                site.login(args.user, args.password)
                print(f"      已登录: {args.user}\n")
            except Exception as e:
                print(f"      [!] 连接失败: {e}")
                sys.exit(1)
        else:
            print("[2/3] 试运行模式，跳过 Wiki 连接\n")

        # 执行增量同步
        print("[3/3] 增量同步...")
        incremental_sync(
            source_dir=source_dir,
            changes=changes,
            site=site,
            delay=args.delay,
            dry_run=args.dry_run
        )

        print()
        print("=" * 50)
        print("  增量同步完成!")
        print("=" * 50)
        return

    # --- 全量同步模式 ---
    print("[1/3] 扫描源文件...")
    files = scan_files(source_dir, CONTENT_DIRS, include_honor=include_honor)
    print(f"      找到 {len(files)} 个文件\n")

    if not files:
        print("没有找到需要同步的文件，退出。")
        return

    # --- 2. 连接 Wiki ---
    site = None
    if not args.dry_run:
        print("[2/3] 连接 Wiki...")
        try:
            site = mwclient.Site(args.wiki_site, path=args.wiki_path if hasattr(args, 'wiki_path') else WIKI_PATH)
            site.login(args.user, args.password)
            print(f"      已登录: {args.user}\n")
        except Exception as e:
            print(f"      [!] 连接失败: {e}")
            sys.exit(1)

        # --- 清理旧页面 ---
        if not args.skip_cleanup:
            print("[2.5/3] 清理旧页面...")
            cleanup_categories = list(AUTO_SYNC_CATEGORIES)
            if not include_honor and "荣誉室" in cleanup_categories:
                cleanup_categories.remove("荣誉室")
            cleanup_wiki_pages(site, cleanup_categories, args.delay, dry_run=False)
    else:
        print("[2/3] 试运行模式，跳过 Wiki 连接\n")

    # --- 3. 同步文件 ---
    print("[3/3] 同步文件到 Wiki...")
    sync_files_to_wiki(
        source_dir=source_dir,
        files=files,
        site=site,
        delay=args.delay,
        dry_run=args.dry_run,
        filter_path=args.filter_path,
    )

    print()
    print("=" * 50)
    print("  同步完成!")
    print("=" * 50)


if __name__ == "__main__":
    main()
