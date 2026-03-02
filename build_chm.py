#!/usr/bin/env python3
"""
序列库 CHM/ZIP 构建脚本

将序列库的 txt/html/docx/doc 文件转换为 HTML 并编译成 CHM，同时生成 ZIP 压缩包。

用法:
    python3 build_chm.py --version v6.2
    python3 build_chm.py --version v6.2 --skip-chm   # 只生成 ZIP
    python3 build_chm.py --version v6.2 --skip-zip   # 只生成 CHM

依赖:
    - pandoc (用于 docx/doc 转 HTML)
    - chmcmd (Linux, fp-utils 包) 或 hhc.exe (Windows, HTML Help Workshop)
"""

import argparse
import html
import os
import re
import shutil
import subprocess
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path

# ============================================================
# 配置
# ============================================================

# CHM 只包含序列库，不含荣誉室
CHM_CONTENT_DIRS = ["序列库"]
# ZIP 包含全部
ZIP_CONTENT_DIRS = ["序列库", "荣誉室"]
ROOT_EXTENSIONS = {".txt", ".html", ".htm", ".docx", ".doc"}
DEFAULT_TITLE = "序列库"

PAGE_STYLE = """\
body {
    font-family: "Microsoft YaHei", "SimSun", "PingFang SC", sans-serif;
    padding: 15px 20px;
    line-height: 1.8;
    color: #333;
    max-width: 900px;
}
h1 {
    color: #2c3e50;
    border-bottom: 2px solid #3498db;
    padding-bottom: 8px;
    font-size: 1.4em;
}
pre, code {
    background: #f5f5f5;
    padding: 2px 6px;
    border-radius: 4px;
}
pre { padding: 10px; overflow-x: auto; }
table { border-collapse: collapse; width: 100%; margin: 10px 0; }
th, td { border: 1px solid #ddd; padding: 6px 10px; text-align: left; }
th { background: #f0f0f0; }
"""

# CHM 内部项目文件的编码 —— 必须是 GBK 才能让 Windows CHM 查看器正确显示中文
CHM_PROJECT_ENCODING = "gbk"


# ============================================================
# 数据结构
# ============================================================


@dataclass
class FileEntry:
    """一个文件的转换记录"""
    original_rel: Path  # 原始相对路径 (如 序列库/职业/001】天师.txt)
    ascii_path: str     # CHM 内部路径 (纯 ASCII, 如 f/00001.html)
    display_name: str   # 显示名称 (如 001】天师)，GBK 编码用于 TOC/索引/全文检索
    dir_parts: tuple    # 目录层级 (如 ("序列库", "职业"))


# ============================================================
# 工具函数
# ============================================================


def sort_key(name: str):
    """排序键：按前导编号排序，无编号的排后面"""
    match = re.match(r"(\d+)", name)
    if match:
        return (0, int(match.group(1)), name)
    return (1, 0, name)


def read_text_file(path: Path) -> str:
    """读取文本文件，自动检测编码"""
    for enc in ("utf-8", "utf-8-sig", "gbk", "gb2312", "big5"):
        try:
            return path.read_text(encoding=enc)
        except (UnicodeDecodeError, UnicodeError):
            continue
    return path.read_text(encoding="utf-8", errors="replace")


def gbk_safe(text: str) -> str:
    """确保文本可以被 GBK 编码，替换不可编码的字符"""
    return text.encode("gbk", errors="replace").decode("gbk", errors="replace")


# ============================================================
# 文件转换（统一输出 GBK 编码 HTML）
# ============================================================


def write_gbk_html(path: Path, content: str):
    """
    将 HTML 内容写为 GBK 编码文件。
    GBK 不支持的字符自动转为 HTML 实体 (&#NNNN;)，查看器照样能渲染。
    """
    path.write_bytes(content.encode("gbk", errors="xmlcharrefreplace"))


def txt_to_html(txt_path: Path) -> str:
    """将 txt 文件内容转为 HTML 字符串（GBK charset）"""
    content = read_text_file(txt_path)
    lines = content.split("\n")

    # 按编者注规范：首行为标题，标题后空一行
    title = txt_path.stem
    content_start = 0
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped:
            title = stripped
            content_start = i + 1
            break

    while content_start < len(lines) and not lines[content_start].strip():
        content_start += 1

    body_text = "\n".join(lines[content_start:])
    escaped_body = html.escape(body_text).replace("\n", "<br>\n")

    return f"""<!DOCTYPE html>
<html>
<head>
<meta charset="gbk">
<meta http-equiv="Content-Type" content="text/html; charset=gbk">
<title>{html.escape(title)}</title>
<style>{PAGE_STYLE}</style>
</head>
<body>
<h1>{html.escape(title)}</h1>
<p>{escaped_body}</p>
</body>
</html>"""


def convert_with_pandoc(src_path: Path, dst_path: Path) -> bool:
    """用 pandoc 将 docx/doc 转为 HTML，然后转码为 GBK"""
    dst_path.parent.mkdir(parents=True, exist_ok=True)
    try:
        result = subprocess.run(
            [
                "pandoc", str(src_path),
                "-o", "-",  # 输出到 stdout
                "--standalone", "-t", "html5",
                "--metadata", f"title={src_path.stem}",
            ],
            check=True, capture_output=True, timeout=60,
        )
        # pandoc 输出 UTF-8，转为 GBK
        utf8_html = result.stdout.decode("utf-8", errors="replace")
        # 替换 charset 声明
        utf8_html = utf8_html.replace('charset="utf-8"', 'charset="gbk"')
        utf8_html = utf8_html.replace("charset=utf-8", "charset=gbk")
        write_gbk_html(dst_path, utf8_html)
        return True
    except (subprocess.CalledProcessError, FileNotFoundError, subprocess.TimeoutExpired) as e:
        print(f"  [!] pandoc 转换失败: {src_path.name} ({e})")
        return False


def make_fallback_html(name: str, ext: str) -> str:
    """为无法转换的文件生成占位 HTML"""
    return f"""<!DOCTYPE html>
<html>
<head><meta charset="gbk"><title>{html.escape(name)}</title>
<style>{PAGE_STYLE}</style></head>
<body><h1>{html.escape(name)}</h1>
<p><em>此文件原为 {ext} 格式，自动转换失败，请参阅原始文件。</em></p>
</body></html>"""


# ============================================================
# 扫描与转换
# ============================================================


def scan_source_files(source_dir: Path, content_dirs: list, include_root: bool = True) -> dict:
    """扫描源目录，返回 {相对路径: 绝对路径}"""
    supported = {".txt", ".html", ".htm", ".docx", ".doc"}
    files = {}

    if include_root:
        for f in sorted(source_dir.iterdir(), key=lambda p: sort_key(p.name)):
            if f.is_file() and f.suffix.lower() in ROOT_EXTENSIONS:
                files[f.relative_to(source_dir)] = f

    for dir_name in content_dirs:
        content_dir = source_dir / dir_name
        if not content_dir.is_dir():
            print(f"  [!] 目录不存在，跳过: {dir_name}")
            continue
        for f in sorted(content_dir.rglob("*"), key=lambda p: sort_key(p.name)):
            if f.is_file() and f.suffix.lower() in supported:
                files[f.relative_to(source_dir)] = f

    return files


def convert_all_to_html(source_dir: Path, build_dir: Path, files: dict) -> list:
    """
    将所有文件转为 GBK 编码 HTML，使用纯 ASCII 文件名存储到 build_dir。
    中文显示名称保留在 FileEntry 中，供 TOC/索引使用（GBK 编码）。
    返回 FileEntry 列表。
    """
    entries = []
    counter = 0

    for rel_path, abs_path in files.items():
        ext = rel_path.suffix.lower()
        counter += 1

        # 文件路径用纯 ASCII，确保 hhc.exe 在任何 Windows 语言环境下都能编译
        ascii_path = f"f/{counter:05d}.html"
        ascii_abs = build_dir / ascii_path
        ascii_abs.parent.mkdir(parents=True, exist_ok=True)

        display_name = rel_path.stem
        dir_parts = tuple(rel_path.parent.parts) if rel_path.parent != Path(".") else ()

        success = False
        if ext == ".txt":
            print(f"  [txt ] {rel_path}")
            write_gbk_html(ascii_abs, txt_to_html(abs_path))
            success = True
        elif ext in (".html", ".htm"):
            print(f"  [html] {rel_path}")
            try:
                content = read_text_file(abs_path)
                content = content.replace('charset="utf-8"', 'charset="gbk"')
                content = content.replace("charset=utf-8", "charset=gbk")
                write_gbk_html(ascii_abs, content)
            except Exception:
                shutil.copy2(abs_path, ascii_abs)
            success = True
        elif ext in (".docx", ".doc"):
            print(f"  [doc ] {rel_path}")
            if convert_with_pandoc(abs_path, ascii_abs):
                success = True
            else:
                write_gbk_html(ascii_abs, make_fallback_html(abs_path.stem, ext))
                success = True

        if success:
            entries.append(FileEntry(
                original_rel=rel_path,
                ascii_path=ascii_path,
                display_name=display_name,
                dir_parts=dir_parts,
            ))

    return entries


# ============================================================
# 首页
# ============================================================


def create_index_page(build_dir: Path, title: str, version: str):
    page = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="gbk">
<title>{html.escape(title)}</title>
<style>
{PAGE_STYLE}
.section {{ margin: 15px 0; }}
.section h2 {{ color: #2980b9; font-size: 1.2em; }}
.info {{ color: #666; font-size: 14px; }}
hr {{ border: none; border-top: 1px solid #ddd; margin: 20px 0; }}
</style>
</head>
<body>
<h1>{html.escape(title)}</h1>
<p class="info">版本: {html.escape(version)} | 使用左侧目录树浏览内容</p>
<div class="section">
<h2>序列库</h2>
<p>当前版本在用的所有资源，包括特质改造、职业、技能表、能量池、公共建筑等。</p>
</div>
<div class="section">
<h2>荣誉室</h2>
<p>已下架或归档的历史资源。</p>
</div>
<hr>
<p class="info">制作人：沧羽 | 发现错漏请联系QQ：853304398</p>
</body>
</html>"""
    write_gbk_html(build_dir / "index.html", page)


# ============================================================
# CHM 项目文件生成（统一 GBK 编码）
# ============================================================


def build_toc_tree(entries: list) -> dict:
    """
    从 FileEntry 列表构建目录树。
    目录节点为 dict，文件节点为 (display_name, ascii_path) 元组。
    """
    tree = {}
    for entry in entries:
        node = tree
        for part in entry.dir_parts:
            if part not in node:
                node[part] = {}
            elif not isinstance(node[part], dict):
                # 如果同名的已经是文件节点，跳过（不应该发生）
                break
            node = node[part]
        else:
            # 用原始文件名做 key（确保同目录下唯一），存 (显示名, ASCII路径) 做 value
            original_filename = entry.original_rel.name
            node[original_filename] = (entry.display_name, entry.ascii_path)
    return tree


def generate_hhc(tree: dict, output_path: Path):
    """
    生成 .hhc 目录文件（GBK 编码）。
    显示名称用 GBK 编码（CHM 查看器按 Language=0x804 解码），文件路径为纯 ASCII。
    """
    lines = [
        '<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">',
        "<HTML><HEAD>",
        '<meta http-equiv="Content-Type" content="text/html; charset=gbk">',
        '<meta name="GENERATOR" content="XuLieKu Builder">',
        "</HEAD><BODY>",
        '<OBJECT type="text/site properties">',
        '  <param name="ImageType" value="Folder">',
        "</OBJECT>",
        "<UL>",
        '<LI> <OBJECT type="text/sitemap">',
        '  <param name="Name" value="首页">',
        '  <param name="Local" value="index.html">',
        "  </OBJECT>",
    ]

    def _write_node(node, indent=0):
        pfx = "  " * indent
        for key in sorted(node.keys(), key=sort_key):
            value = node[key]
            if isinstance(value, dict):
                safe_name = gbk_safe(key)
                lines.append(f'{pfx}<LI> <OBJECT type="text/sitemap">')
                lines.append(f'{pfx}  <param name="Name" value="{safe_name}">')
                lines.append(f"{pfx}  </OBJECT>")
                lines.append(f"{pfx}<UL>")
                _write_node(value, indent + 1)
                lines.append(f"{pfx}</UL>")
            elif isinstance(value, tuple):
                display_name, ascii_path = value
                safe_name = gbk_safe(display_name)
                lines.append(f'{pfx}<LI> <OBJECT type="text/sitemap">')
                lines.append(f'{pfx}  <param name="Name" value="{safe_name}">')
                lines.append(f'{pfx}  <param name="Local" value="{ascii_path}">')
                lines.append(f"{pfx}  </OBJECT>")

    _write_node(tree)
    lines.extend(["</UL>", "</BODY></HTML>"])

    output_path.write_bytes("\n".join(lines).encode("gbk", errors="replace"))


def generate_hhk(entries: list, output_path: Path):
    """生成 .hhk 索引文件（GBK 编码）"""
    lines = [
        '<!DOCTYPE HTML PUBLIC "-//IETF//DTD HTML//EN">',
        "<HTML><HEAD>",
        '<meta http-equiv="Content-Type" content="text/html; charset=gbk">',
        '<meta name="GENERATOR" content="XuLieKu Builder">',
        "</HEAD><BODY>",
        "<UL>",
    ]

    for entry in sorted(entries, key=lambda e: sort_key(e.display_name)):
        safe_name = gbk_safe(entry.display_name)
        lines.append(f'<LI> <OBJECT type="text/sitemap">')
        lines.append(f'  <param name="Name" value="{safe_name}">')
        lines.append(f'  <param name="Local" value="{entry.ascii_path}">')
        lines.append(f"  </OBJECT>")

    lines.extend(["</UL>", "</BODY></HTML>"])

    output_path.write_bytes("\n".join(lines).encode("gbk", errors="replace"))


def generate_hhp(entries: list, output_path: Path, chm_filename: str, title: str):
    """生成 .hhp 项目文件（GBK 编码，文件路径为 ASCII）"""
    file_list = "\n".join(e.ascii_path for e in entries)

    # .hhp 不是 HTML，不能用实体编码，标题用 GBK
    content = f"""[OPTIONS]
Compatibility=1.1 or later
Compiled file={chm_filename}
Contents file=toc.hhc
Default topic=index.html
Display compile progress=Yes
Full-text search=Yes
Index file=index.hhk
Language=0x804
Title={title}

[FILES]
index.html
{file_list}
"""
    output_path.write_bytes(content.encode(CHM_PROJECT_ENCODING, errors="replace"))


# ============================================================
# CHM 编译
# ============================================================


def find_chm_compiler():
    """查找 CHM 编译器，返回 (路径, 类型) 或 (None, None)"""
    if sys.platform == "win32":
        # 优先查找仓库内置的 hhc.exe（tools/hhw/）
        bundled = Path(__file__).parent / "tools" / "hhw" / "hhc.exe"
        if bundled.is_file():
            return str(bundled), "hhc"
        for candidate in [
            r"C:\Program Files (x86)\HTML Help Workshop\hhc.exe",
            r"C:\Program Files\HTML Help Workshop\hhc.exe",
        ]:
            if os.path.isfile(candidate):
                return candidate, "hhc"
        found = shutil.which("hhc.exe")
        if found:
            return found, "hhc"

    found = shutil.which("chmcmd")
    if found:
        return found, "chmcmd"

    return None, None


def compile_chm(build_dir: Path, hhp_file: str) -> bool:
    """编译 CHM"""
    compiler, compiler_type = find_chm_compiler()
    if not compiler:
        print("  [!] 未找到 CHM 编译器")
        print("      Linux/CI:  sudo apt-get install fp-utils")
        print("      Windows:   安装 HTML Help Workshop")
        return False

    print(f"  编译器: {compiler} ({compiler_type})")

    try:
        result = subprocess.run(
            [compiler, hhp_file],
            cwd=build_dir,
            capture_output=True,
            timeout=300,  # 5 分钟超时，防止 hhc.exe 卡死
        )
    except subprocess.TimeoutExpired:
        print("  [!] CHM 编译超时（5分钟），已终止")
        return False

    # hhc.exe 返回 1 表示成功，chmcmd 返回 0 表示成功
    success = (result.returncode == 1) if compiler_type == "hhc" else (result.returncode == 0)

    if success:
        print("  CHM 编译成功!")
        return True
    else:
        print(f"  CHM 编译失败 (返回码: {result.returncode})")
        for label, data in [("stdout", result.stdout), ("stderr", result.stderr)]:
            try:
                text = data.decode("gbk", errors="replace").strip()
            except Exception:
                text = str(data)
            if text:
                print(f"  {label}: {text[:500]}")
        return False


# ============================================================
# ZIP 打包
# ============================================================


def create_zip(source_dir: Path, output_path: Path):
    """创建 ZIP 压缩包（原始文件，包含全部目录）"""
    count = 0
    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for dir_name in ZIP_CONTENT_DIRS:
            content_dir = source_dir / dir_name
            if not content_dir.is_dir():
                continue
            for f in sorted(content_dir.rglob("*")):
                if f.is_file():
                    zf.write(f, str(f.relative_to(source_dir)))
                    count += 1

        for f in sorted(source_dir.iterdir(), key=lambda p: sort_key(p.name)):
            if f.is_file() and f.suffix.lower() in ROOT_EXTENSIONS:
                zf.write(f, f.name)
                count += 1

    print(f"  ZIP 已生成: {output_path.name} ({count} 个文件)")


# ============================================================
# 主流程
# ============================================================


def main():
    parser = argparse.ArgumentParser(description="序列库 CHM/ZIP 构建工具")
    parser.add_argument("--version", default="dev", help="版本号 (如 v6.2)")
    parser.add_argument("--source-dir", default=".", help="源文件目录")
    parser.add_argument("--output-dir", default="dist", help="输出目录")
    parser.add_argument("--skip-chm", action="store_true", help="跳过 CHM 编译")
    parser.add_argument("--skip-zip", action="store_true", help="跳过 ZIP 打包")
    args = parser.parse_args()

    source_dir = Path(args.source_dir).resolve()
    output_dir = Path(args.output_dir).resolve()
    build_dir = output_dir / "_chm_build"
    version = args.version

    # 统一命名格式：宏观界域强化序列库V版本号
    # 版本号处理：v6.2 → V6.2, 6.2 → V6.2
    if version != "dev":
        ver_num = version.lstrip("vV")
        display_version = f"V{ver_num}"
        product_name = f"宏观界域强化序列库{display_version}"
    else:
        display_version = "dev"
        product_name = "宏观界域强化序列库"

    title = product_name
    chm_internal = "output.chm"  # CHM 编译内部用 ASCII 名，编译后再改名
    chm_final_name = f"{product_name}.chm"
    zip_filename = f"{product_name}.zip"

    print("=" * 50)
    print(f"  序列库构建工具")
    print(f"  源目录:   {source_dir}")
    print(f"  输出目录: {output_dir}")
    print(f"  版本:     {version}")
    print("=" * 50)
    print()

    if build_dir.exists():
        shutil.rmtree(build_dir)
    build_dir.mkdir(parents=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    # --- 1. 扫描（CHM 只包含序列库，不含荣誉室）---
    print("[1/6] 扫描源文件...")
    files = scan_source_files(source_dir, CHM_CONTENT_DIRS, include_root=True)
    print(f"      找到 {len(files)} 个文件（CHM 用）\n")

    # --- 2. 转换 ---
    print("[2/6] 转换文件为 HTML（GBK 编码，ASCII 文件名）...")
    entries = convert_all_to_html(source_dir, build_dir, files)
    print(f"      生成 {len(entries)} 个 HTML 文件\n")

    # --- 3. 首页 ---
    print("[3/6] 创建首页...")
    create_index_page(build_dir, title, display_version)
    print()

    # --- 4 & 5. CHM ---
    if not args.skip_chm:
        print("[4/6] 生成 CHM 项目文件（GBK 编码）...")
        toc_tree = build_toc_tree(entries)
        generate_hhc(toc_tree, build_dir / "toc.hhc")
        generate_hhk(entries, build_dir / "index.hhk")
        generate_hhp(entries, build_dir / "project.hhp", chm_internal, title)
        print("      已生成: project.hhp, toc.hhc, index.hhk\n")

        print("[5/6] 编译 CHM...")
        if compile_chm(build_dir, "project.hhp"):
            chm_src = build_dir / chm_internal
            chm_dst = output_dir / chm_final_name
            if chm_src.exists():
                shutil.copy2(str(chm_src), str(chm_dst))
                size_mb = chm_dst.stat().st_size / (1024 * 1024)
                print(f"      输出: {chm_dst} ({size_mb:.1f} MB)\n")
            else:
                print("      [!] CHM 文件未生成\n")
        else:
            print("      CHM 编译失败\n")
    else:
        print("[4/6] 跳过 CHM\n[5/6] 跳过 CHM\n")

    # --- 6. ZIP ---
    if not args.skip_zip:
        print("[6/6] 创建 ZIP 压缩包...")
        create_zip(source_dir, output_dir / zip_filename)
        print()
    else:
        print("[6/6] 跳过 ZIP\n")

    # 清理
    if build_dir.exists():
        shutil.rmtree(build_dir)

    print("=" * 50)
    print("  构建完成!")
    print("=" * 50)


if __name__ == "__main__":
    main()
