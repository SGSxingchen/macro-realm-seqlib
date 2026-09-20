"""项目文本资源的统一解码入口。

资源正文最终统一按 UTF-8 保存；读取时兼容历史投稿可能使用的编码，
并优先识别带 BOM 的 UTF-16 文件，避免把 UTF-16 字节误当成 UTF-8/GBK。
"""

from __future__ import annotations

from pathlib import Path


TEXT_ENCODINGS = (
    "utf-8-sig",
    "utf-16",
    "utf-16-le",
    "utf-16-be",
    "utf-8",
    "gbk",
    "gb2312",
    "big5",
)


def decode_text_bytes(data: bytes) -> tuple[str, str]:
    """解码文本字节并返回 ``(文本, 实际编码)``。

    UTF-16 放在 UTF-8/GBK 兼容回退之前，带 BOM 的文件会由 ``utf-16``
    自动识别端序并去掉 BOM。最后的 replace 仅作为损坏文件的可读性兜底。
    """

    for encoding in TEXT_ENCODINGS:
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def read_text(path: Path) -> tuple[str, str]:
    """读取路径文本并返回 ``(文本, 实际编码)``。"""

    return decode_text_bytes(path.read_bytes())
