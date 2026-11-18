"""项目文本资源的统一解码入口。

资源正文最终统一按 UTF-8 保存；读取时兼容历史投稿可能使用的编码，
并优先识别带 BOM 的 UTF-16 文件，避免把 UTF-16 字节误当成 UTF-8/GBK。
"""

from __future__ import annotations

from codecs import BOM_UTF16_BE, BOM_UTF16_LE
from pathlib import Path


LEGACY_ENCODINGS = ("gbk", "gb2312", "big5")


def _encoding_candidates(data: bytes) -> tuple[str, ...]:
    if data.startswith((BOM_UTF16_BE, BOM_UTF16_LE)):
        return ("utf-16",)
    # Bare UTF-16 decoders accept many even-length GBK byte strings as gibberish.
    # For BOM-less resources, infer endianness from ASCII newlines, or NUL parity.
    if len(data) % 2 == 0 and b"\0" in data:
        even, odd = data[::2], data[1::2]
        le = sum(a in b"\r\n" and b == 0 for a, b in zip(even, odd))
        be = sum(a == 0 and b in b"\r\n" for a, b in zip(even, odd))
        if le != be or even.count(0) != odd.count(0):
            encoding = "utf-16-le" if (le, odd.count(0)) > (be, even.count(0)) else "utf-16-be"
            return (encoding, "utf-8-sig", *LEGACY_ENCODINGS)
    return ("utf-8-sig", *LEGACY_ENCODINGS)


def decode_text_bytes(data: bytes) -> tuple[str, str]:
    """解码文本字节并返回 ``(文本, 实际编码)``。

    先识别 UTF-16 BOM 或无 BOM 文本的换行特征，避免把 GBK 误读成
    UTF-16。最后的 replace 仅作为损坏文件的可读性兜底。
    """

    for encoding in _encoding_candidates(data):
        try:
            return data.decode(encoding), encoding
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace"), "utf-8-replace"


def read_text(path: Path) -> tuple[str, str]:
    """读取路径文本并返回 ``(文本, 实际编码)``。"""

    return decode_text_bytes(path.read_bytes())
