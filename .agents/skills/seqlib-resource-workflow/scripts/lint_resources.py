#!/usr/bin/env python3
"""宏观界域强化序列库 · 资源结构探针与格式门禁 v3（AI Harness 组件 1/2）。

定位：
- 改前当「病灶雷达」：精准定位 0 卡老表、空心括号 `〖...〗`、无冒号卡头、半全角混搭括号、
  行首 `【XXX】：` 孤儿字段隐患、卡前漂移字段、卡顶多层 `【...】` 语义歧义、首部日志吞背景、
  多人合集表内嵌作者、旧遗留 `【日期待复核】` 等全部深水区问题。
- 改后当「语法考官」：校验 AI 改完后的资源是否 100% 符合《资源标准模板.txt》与阅读器契约。

铁律：全程严格只读，绝不自动修改任何资源文件（一切文本修改均由 AI 结合上下文完成）。
"""

from __future__ import annotations

import argparse
from collections import Counter
import json
import re
import subprocess
import sys
from dataclasses import asdict, dataclass
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[4]
if str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))

try:
    from text_encoding import decode_text_bytes
except ImportError:  # pragma: no cover
    def decode_text_bytes(data: bytes) -> tuple[str, str]:
        for enc in ("utf-8-sig", "utf-8", "utf-16", "gbk", "gb2312", "big5"):
            try:
                return data.decode(enc), enc
            except UnicodeDecodeError:
                continue
        return data.decode("utf-8", errors="replace"), "utf-8-replace"


DIRTY_NAMES = {".DS_Store", "Thumbs.db", "desktop.ini"}
INVALID_FILENAME_CHARS_RE = re.compile(r'[/\\:*?"<>|]')
NUMBERED_PREFIX_RE = re.compile(r"^\d{3}】.+")

STANDARD_ROLES = {"制作人", "调整人", "修改人", "重置人", "审核人"}
COMPAT_ROLES = STANDARD_ROLES | {
    "制作者", "投稿人", "作者", "创作者", "码字者", "二次制作",
    "审核", "审核者", "审核员", "复查人", "文本优化", "优化人", "加强人", "更新人", "翻修人", "协力",
}
ROLE_GROUP = "|".join(sorted(COMPAT_ROLES, key=len, reverse=True))

PAREN_CREDIT_RE = re.compile(rf"^[ \t]*[（(]\s*({ROLE_GROUP})\s*[：:][^）)\r\n]+[）)][ \t]*$")
BARE_CREDIT_RE = re.compile(rf"^[ \t]*({ROLE_GROUP})\s*[：:]\s*\S+")
BRACKET_CREDIT_RE = re.compile(rf"^[ \t]*[\[［【〖〔]\s*({ROLE_GROUP})\s*[：:]([^\]］】〗〕\r\n]+)[\]］】〗〕][ \t]*$")

CHANGELOG_STD_RE = re.compile(
    r"^[ \t]*【\s*(\d{2,4}[./-]\d{1,2}(?:[./-]\d{1,2})?)\s*[：:](.+?)】(?:\s*[（(].*[）)])?[ \t]*$"
)
CHANGELOG_UNDATED_STD_RE = re.compile(
    r"^[ \t]*【\s*历史更新\s*[：:](.+?)】(?:\s*[（(].*[）)])?[ \t]*$"
)
CHANGELOG_PENDING_LEGACY_RE = re.compile(
    r"^[ \t]*【\s*日期待复核\s*[：:](.+?)】[ \t]*$"
)
CHANGELOG_ASCII_RE = re.compile(
    r"^[ \t]*\[\s*(\d{2,4}[./-]\d{1,2}(?:[./-]\d{1,2})?)\s*[：:](.+?)\][ \t]*$"
)
CHANGELOG_CN_OR_BARE_RE = re.compile(
    r"^[ \t]*【?\s*((?:20\d{2}|\d{2})年\d{1,2}月(?:\d{1,2}日)?|\d{2,4}[./]\d{1,2}[./]\d{1,2})\s*(?:更新|修改|调整|重置|补充)?\s*[：:](.+?)】?[ \t]*$"
)
CHANGELOG_UNDATED_RAW_RE = re.compile(
    r"^[ \t]*(?:\*?[ \t]*(?:更新|更新说明|修改内容|调整内容|重置内容)[：:]|\d{3,4}(?:更新|特别)[：:行动]*|【\d{3,4}[：:])(.+)$"
)

CARD_STARTER_STD_RE = re.compile(r"^[ \t]*\[([^\]\r\n]{0,12}名称)\][ \t]*[：:][ \t]*(.*)$")
CARD_STARTER_COMPAT_RE = re.compile(
    r"^[ \t]*(?:【([^】\r\n]{0,12}名称)】|［([^］\r\n]{0,12}名称)］|〖([^〗\r\n]{0,12}名称)〗|〔([^〕\r\n]{0,12}名称)〕)[ \t]*[：:][ \t]*(.*)$"
)

LENTICULAR_FIELD_RE = re.compile(r"^[ \t]*[〖〔]([^〗〕\r\n]{1,24})[〗〕][ \t]*[：:]?")
MIXED_BRACKET_FIELD_RE = re.compile(
    r"^[ \t]*(?:\[([^\]］】\r\n]{1,24})[］】]|[［【]([^\]］】\r\n]{1,24})\])[ \t]*[：:]"
)
BROKEN_LEFT_BRACKET_RE = re.compile(r"^[ \t]*([\u4e00-\u9fff]{0,10}名称)[\]】］〗][ \t]*[：:]")
COLON_INSIDE_BRACKET_RE = re.compile(r"^[ \t]*[\[【［〖]([\u4e00-\u9fff]{1,16})[：:][\]】］〗]")
MISSING_COLON_STARTER_RE = re.compile(
    r"^[ \t]*(?:\[([\u4e00-\u9fff]{1,10}名称)\]|【([\u4e00-\u9fff]{1,10}名称)】|［([\u4e00-\u9fff]{1,10}名称)］|〖([\u4e00-\u9fff]{1,10}名称)〗)[ \t]*([^\s：:\]】］〗].*)$"
)
FULLWIDTH_FIELD_RE = re.compile(r"^[ \t]*［([^］\r\n]{1,24})］[ \t]*[：:]?")
HEAVY_BRACKET_FIELD_RE = re.compile(r"^[ \t]*【([^】\r\n]{1,28})】[ \t]*[：:][ \t]*(.*)$")

EXPLICIT_FIELD_STD_RE = re.compile(r"^[ \t]*\[([^\]\r\n]{1,32})\][ \t]*[：:][ \t]*(.*)$")
JIYI_RULE_BAD_RE = re.compile(r"^[ \t]*(?:【技艺规则】[ \t]*[：:]?|\[技艺规则\][ \t]*$)")
SECTION_BRACKET_RE = re.compile(r"^[ \t]*【([^】\r\n]{1,160})】[ \t]*$")
DECORATIVE_DIVIDER_RE = re.compile(r"^[ \t]*[—―─\-=_=~～]{6,}[ \t]*$")
BARE_TRAIT_GROWTH_RE = re.compile(r"^[ \t]*(EX|[A-ES])级[：:][ \t]*\S+")
BARE_INLINE_POOL_RE = re.compile(r"^[ \t]*能量池[：:][ \t]*\S+")

KNOWN_SECTION_TITLES = {
    "通用规则", "购买规则", "升级规则", "技能表通用规则", "职业通用规则", "特质通用规则",
    "特质说明", "职业说明", "技能表说明", "能量池说明", "建筑简介", "建筑说明", "特殊规则",
    "称号规则", "道具规则", "能量池规则", "建筑规则", "限制", "升级树", "称号",
    "职业称号", "称号技能", "职业技能", "专属能量池", "能量池", "本事", "神通", "根器",
    "扩充技能组", "忠诚派", "混沌派", "改造手术",
    "基础性能", "前置能力", "前置技能", "初级能力", "中级能力", "高级能力", "终极能力", "特殊能力",
    "特殊技能", "基础技能", "进阶技能", "核心能力", "核心技能", "被动能力", "主动能力",
    "成长型能力", "成长性能力", "成长能力", "可成长能力", "特定高级能力", "特殊高级能力",
    "高级特质能力", "固有能力", "背景", "宝具", "历史记录", "历史更新记录", "说明", "备注",
    "商品列表", "食品列表", "饮品列表", "兑换列表",
}

MOD_NOTE_HINT_RE = re.compile(
    r"(修改|调整|加强|削弱|重置|重写|修复|修理|补充|新增|删除|移除|回调|下调|上调|降低|提高|提升|减少|增加|改为|→|->|明确|优化|合并)"
)
UNLOCK_OR_SCOPE_HINT_RE = re.compile(
    r"(需要|前置|由.+升级|升级而来|仅限|仅能|不可与|解锁|习得条件|兑换条件|前置要求|以下.{0,6}(?:技能|能力|法术|义体))"
)

PRE_CARD_FLOATING_KEYS = {
    "技艺条件", "技艺难度", "技艺规则", "解锁条件", "前置条件", "使用要求", "兑换条件", "学习条件", "消耗"
}
NON_STD_FIELD_ALIASES = {
    "初级能力效果": "初级效果",
    "中级能力效果": "中级效果",
    "高级能力效果": "高级效果",
    "终极能力效果": "终极效果",
    "终级效果": "终极效果",
    "终级能力效果": "终极效果",
    "技艺条件": "技艺规则",
    "技艺难度": "技艺规则",
    "简介": "能力简介",
    "技艺简介": "能力简介",
    "技艺效果": "能力效果",
    "技艺形容": "能力形容",
    "效果补充": "补充说明",
}
COMMON_CARD_FIELDS_IN_HEAVY_BRACKETS = {
    "简介", "能力简介", "能力介绍", "能力效果", "技能效果", "能力形容", "技能形容", "能力描述",
    "释放类型", "打击类型", "伤害类型", "最大距离", "攻击距离", "消耗能量", "能力消耗",
    "冷却时间", "技能冷却", "持续时间", "使用要求", "补充说明", "技艺规则", "能力备注",
    "初级效果", "中级效果", "高级效果", "终极效果",
}
CARD_TAIL_META_FIELDS = {
    "释放类型", "打击类型", "伤害类型", "最大距离", "攻击距离", "消耗能量",
    "能力消耗", "技能消耗", "冷却时间", "技能冷却", "能力冷却", "持续时间", "补充说明",
}
OPTIONAL_CARD_FIELDS = CARD_TAIL_META_FIELDS | {
    "技能类型", "能力类型", "使用要求", "能力形容", "能力简介", "占用部位", "单位限制", "孵化消耗",
}


@dataclass
class Issue:
    severity: str  # "ERROR" | "WARN"
    code: str
    path: str
    line: int
    message: str


def parse_date_tuple(raw_date: str) -> tuple[int, int, int] | None:
    cleaned = raw_date.strip()
    m_cn = re.match(r"^(\d{2,4})年(\d{1,2})月(?:(\d{1,2})日)?$", cleaned)
    if m_cn:
        year = int(m_cn.group(1))
        if year < 100:
            year += 2000
        month = int(m_cn.group(2))
        day = int(m_cn.group(3)) if m_cn.group(3) else 1
        return (year, month, day)
    parts = re.split(r"[./-]", cleaned)
    if len(parts) not in (2, 3):
        return None
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return None
    year = nums[0]
    if year < 100:
        year += 2000
    month = nums[1]
    day = nums[2] if len(nums) == 3 else 1
    return (year, month, day)


def is_prefixed_card_note(inner: str) -> bool:
    cleaned = inner.strip()
    return bool(re.match(r"^(?:能力修改|说明)[：:]", cleaned))


def is_known_section_title(inner: str) -> bool:
    cleaned = inner.strip()
    if cleaned in KNOWN_SECTION_TITLES:
        return True
    if is_prefixed_card_note(cleaned):
        return True
    if re.match(r"^(?:EX|[A-ES])(?:级)?(?:[/\-~～](?:EX|[A-ES])(?:级)?)?$", cleaned):
        return True
    if re.match(
        r"^(?:第[一二三四五六七八九十\d]+[阶段章节层篇]|.{1,12}(?:能力|技能|法术|奇迹|咒术|模块|形态|分支|列表|一览|说明|规则|称号|篇|技)(?:[（(].+[）)])?)$",
        cleaned,
    ):
        if not MOD_NOTE_HINT_RE.search(cleaned):
            return True
    return False


def match_any_card_starter(line: str) -> tuple[str, str] | None:
    m_std = CARD_STARTER_STD_RE.match(line)
    if m_std:
        return m_std.group(1).strip(), "std"
    m_compat = CARD_STARTER_COMPAT_RE.match(line)
    if m_compat:
        if m_compat.group(1) is not None:
            return m_compat.group(1).strip(), "heavy"
        if m_compat.group(2) is not None:
            return m_compat.group(2).strip(), "fullwidth"
        for g in m_compat.groups()[2:]:
            if g is not None:
                return g.strip(), "compat"
    return None


def is_any_field_or_starter(line: str) -> bool:
    return bool(
        match_any_card_starter(line)
        or EXPLICIT_FIELD_STD_RE.match(line)
        or HEAVY_BRACKET_FIELD_RE.match(line)
        or MIXED_BRACKET_FIELD_RE.match(line)
        or LENTICULAR_FIELD_RE.match(line)
    )


def _check_blank_line_hygiene(lines: list[str], first_non_empty_idx: int, rel_path: str) -> list[Issue]:
    """严格模式下的全文件空行与字段换行门禁（适配前端阅读器两列网格与 afterBlank 切分契约）。"""
    hygiene_issues: list[Issue] = []
    n = len(lines)
    if first_non_empty_idx > 0:
        hygiene_issues.append(Issue("WARN", "LEADING_BLANK_LINE", rel_path, 1, "文件顶部存在多余前导空行"))

    def _is_changelog_line(s: str) -> bool:
        return bool(
            CHANGELOG_STD_RE.match(s)
            or CHANGELOG_UNDATED_STD_RE.match(s)
            or CHANGELOG_PENDING_LEGACY_RE.match(s)
            or CHANGELOG_ASCII_RE.match(s)
        )

    def _is_standalone_section(s: str) -> str | None:
        m = SECTION_BRACKET_RE.match(s)
        if not m or _is_changelog_line(s):
            return None
        inner = m.group(1).strip()
        if is_prefixed_card_note(inner):
            return None
        if is_known_section_title(inner):
            return inner
        return None

    def _is_credit(s: str) -> bool:
        return bool(PAREN_CREDIT_RE.match(s) or BRACKET_CREDIT_RE.match(s) or BARE_CREDIT_RE.match(s))

    # 预标记所有卡片紧贴簇的顶行索引（含紧贴在 `[XX名称]` 正上方的 `【能力修改：...】` / `【说明：...】`）
    card_cluster_starts: set[int] = set()
    for i, raw in enumerate(lines):
        if match_any_card_starter(raw):
            top = i
            while top - 1 > first_non_empty_idx:
                m_prev = SECTION_BRACKET_RE.match(lines[top - 1].strip())
                if m_prev and is_prefixed_card_note(m_prev.group(1)):
                    top -= 1
                else:
                    break
            card_cluster_starts.add(top)
            if top > 0 and lines[top - 1].strip() != "":
                hygiene_issues.append(
                    Issue("WARN", "MISSING_BLANK_BEFORE_CARD", rel_path, top + 1, "卡片起始处上方缺少空行（相邻卡片之间或导语与首卡之间须保留 1 个空行）")
                )

    inside_card = False
    in_hdr = True
    prev_hdr_kind = ""  # "credit" | "log"

    for i in range(first_non_empty_idx + 1, n):
        raw = lines[i]
        stripped = raw.strip()
        lineno = i + 1

        # 1. 连续多空行检测
        if not stripped:
            if not lines[i - 1].strip() and (i - 2 < 0 or lines[i - 2].strip() != ""):
                hygiene_issues.append(
                    Issue("WARN", "CONSECUTIVE_BLANK_LINES", rel_path, lineno, "存在连续多余空行（任意结构块之间最多保留 1 个空行）")
                )
            # 卡片内部空行检测（只在该空行块的第一行触发一次）
            if inside_card and lines[i - 1].strip() != "":
                j = i + 1
                while j < n and not lines[j].strip():
                    j += 1
                if j < n:
                    if j in card_cluster_starts or _is_standalone_section(lines[j].strip()) is not None:
                        inside_card = False
                    else:
                        hygiene_issues.append(
                            Issue(
                                "WARN",
                                "BLANK_LINE_INSIDE_CARD",
                                rel_path,
                                lineno,
                                f"卡片内部存在空行（会导致下方 `{lines[j].strip()[:26]}` 脱离当前字段框或拆散卡片），请改为单换行连续书写",
                            )
                        )
            continue

        # 判断首部人员/日志块空行规整度
        if in_hdr:
            if _is_standalone_section(stripped) is not None or i in card_cluster_starts or match_any_card_starter(raw):
                in_hdr = False
            elif _is_credit(raw):
                if prev_hdr_kind == "credit" and not lines[i - 1].strip():
                    hygiene_issues.append(Issue("WARN", "BLANK_INSIDE_CREDITS", rel_path, lineno, "首部人员信息块内部不应插入空行"))
                prev_hdr_kind = "credit"
            elif _is_changelog_line(raw):
                if prev_hdr_kind == "credit" and lines[i - 1].strip() != "":
                    hygiene_issues.append(Issue("WARN", "NO_BLANK_BEFORE_CHANGELOG", rel_path, lineno, "首部人员块与更新记录块之间缺少空行"))
                elif prev_hdr_kind == "log" and not lines[i - 1].strip():
                    hygiene_issues.append(Issue("WARN", "BLANK_INSIDE_CHANGELOG", rel_path, lineno, "首部连续更新记录块内部不应插入空行"))
                prev_hdr_kind = "log"

        # 2. 单卡前缀备注与下方卡片脱节检测
        m_sec_any = SECTION_BRACKET_RE.match(stripped)
        if m_sec_any and is_prefixed_card_note(m_sec_any.group(1)):
            if i + 1 < n and not lines[i + 1].strip():
                hygiene_issues.append(
                    Issue(
                        "WARN",
                        "PREFIX_NOTE_DETACHED",
                        rel_path,
                        lineno,
                        f"单卡前缀备注 `{stripped[:32]}` 与下方卡片之间多出了空行（须紧贴所属 `[XX名称]` 无空行）",
                    )
                )

        # 3. 独立小节标题上下空行检测
        sec_title = _is_standalone_section(stripped)
        if sec_title is not None:
            inside_card = False
            if i > 0 and lines[i - 1].strip() != "":
                hygiene_issues.append(
                    Issue("WARN", "SECTION_MISSING_BLANK_BEFORE", rel_path, lineno, f"小节标题 `【{sec_title}】` 上方缺少空行")
                )
            if i + 1 < n and lines[i + 1].strip() != "":
                hygiene_issues.append(
                    Issue("WARN", "SECTION_MISSING_BLANK_AFTER", rel_path, lineno, f"小节标题 `【{sec_title}】` 下方缺少空行")
                )
            continue

        if match_any_card_starter(raw):
            inside_card = True

        # 4. 字段冒号后首行悬空换行检测（避免网页两列网格右侧首行错位留空）
        m_std_f = EXPLICIT_FIELD_STD_RE.match(raw)
        if m_std_f and not m_std_f.group(2).strip():
            f_key = m_std_f.group(1).strip()
            next_s = lines[i + 1].strip() if i + 1 < n else ""
            if next_s and not is_any_field_or_starter(lines[i + 1]) and not SECTION_BRACKET_RE.match(next_s):
                hygiene_issues.append(
                    Issue(
                        "WARN",
                        "FIELD_LEADING_NEWLINE",
                        rel_path,
                        lineno,
                        f"字段 `[{f_key}]：` 冒号后首行悬空换行（会导致网页端首行错位留白），请将下一行首条内容同行紧跟在冒号后",
                    )
                )

    return hygiene_issues


def check_txt_file(path: Path, rel_path: str, strict: bool, new_submission: bool = False) -> tuple[list[Issue], int]:
    issues: list[Issue] = []
    try:
        raw_bytes = path.read_bytes()
    except OSError as exc:
        return [Issue("ERROR", "IO_ERROR", rel_path, 0, f"读取文件失败: {exc}")], 0

    text, encoding = decode_text_bytes(raw_bytes)
    if encoding not in ("utf-8", "utf-8-sig"):
        issues.append(Issue("ERROR", "NON_UTF8_ENCODING", rel_path, 1, f"文件编码为 {encoding}，正式资源须保存为 UTF-8"))

    if INVALID_FILENAME_CHARS_RE.search(path.name):
        issues.append(Issue("ERROR", "INVALID_FILENAME", rel_path, 1, f"文件名包含路径非法字符: {path.name}"))
    if not NUMBERED_PREFIX_RE.match(path.stem):
        issues.append(Issue("WARN", "MISSING_NUMBER_PREFIX", rel_path, 1, f"文件名未以三位编号 `NNN】` 开头: {path.name}"))

    lines = text.splitlines()
    if not lines or not text.strip():
        issues.append(Issue("ERROR", "EMPTY_FILE", rel_path, 1, "文件内容为空"))
        return issues, 0

    is_trait_file = "特质改造" in path.parts
    is_energy_pool_file = "能量池" in path.parts
    is_public_building = "公共建筑" in path.parts
    parent_body = re.sub(r"^\d+】", "", path.parent.name)
    self_body = re.sub(r"^\d+】", "", path.stem)
    is_series_overview = bool(re.match(r"^\d+】", path.parent.name)) and (
        self_body == parent_body or any(kw in path.stem for kw in ("规则", "总纲", "简介", "目录"))
    )
    is_annotation_only = "注解" in path.stem or "详解" in path.stem or is_series_overview

    first_non_empty_idx = next((i for i, l in enumerate(lines) if l.strip()), 0)
    first_line = lines[first_non_empty_idx].strip()
    if first_line != path.stem:
        issues.append(
            Issue(
                "ERROR",
                "TITLE_MISMATCH",
                rel_path,
                first_non_empty_idx + 1,
                f"首行标题 `{first_line}` 与文件名 `{path.stem}` 不一致（若含历史尾注请拆入 `【历史更新：...】` 后对齐首行）",
            )
        )
    if strict and first_non_empty_idx + 1 < len(lines) and lines[first_non_empty_idx + 1].strip() != "":
        issues.append(Issue("WARN", "NO_BLANK_AFTER_TITLE", rel_path, first_non_empty_idx + 2, "首行标题后缺少空行"))

    card_count = 0
    in_header = True
    seen_any_credit = False
    seen_bracketed_changelog_in_header = False
    warned_log_swallows_intro = False
    last_changelog_date: tuple[int, int, int] | None = None

    for idx, line in enumerate(lines):
        lineno = idx + 1
        if idx <= first_non_empty_idx:
            continue
        stripped = line.strip()
        if not stripped:
            continue

        m_paren_credit = PAREN_CREDIT_RE.match(line)
        m_bracket_credit = BRACKET_CREDIT_RE.match(line) if not m_paren_credit else None
        m_bare_credit = BARE_CREDIT_RE.match(line) if not (m_paren_credit or m_bracket_credit) else None
        is_credit_line = bool(m_paren_credit or m_bracket_credit or m_bare_credit)
        if is_credit_line and in_header:
            seen_any_credit = True

        # 判断是否离开首部元信息区
        if in_header and not is_credit_line:
            if (
                match_any_card_starter(line)
                or EXPLICIT_FIELD_STD_RE.match(line)
                or MIXED_BRACKET_FIELD_RE.match(line)
                or LENTICULAR_FIELD_RE.match(line)
                or MISSING_COLON_STARTER_RE.match(line)
            ):
                m_sub = EXPLICIT_FIELD_STD_RE.match(line)
                if not m_sub and re.match(r"^\[[^\]：:]{2,12}\]$", stripped):
                    pass
                else:
                    in_header = False
            elif SECTION_BRACKET_RE.match(line) and not (
                CHANGELOG_STD_RE.match(line)
                or CHANGELOG_UNDATED_STD_RE.match(line)
                or CHANGELOG_PENDING_LEGACY_RE.match(line)
            ):
                inner = SECTION_BRACKET_RE.match(line).group(1)  # type: ignore[union-attr]
                if is_known_section_title(inner) or MOD_NOTE_HINT_RE.search(inner):
                    in_header = False

        if in_header:
            if is_credit_line:
                role = (m_paren_credit or m_bracket_credit or m_bare_credit).group(1)  # type: ignore[union-attr]
                if seen_bracketed_changelog_in_header and strict:
                    issues.append(
                        Issue("WARN", "HEADER_INTERLEAVED", rel_path, lineno, "首部人员字段出现在更新日志之后（标准要求人员块在前、空一行、更新日志在后）")
                    )
                if m_bracket_credit and strict:
                    issues.append(
                        Issue("WARN", "BRACKETED_CREDIT_LINE", rel_path, lineno, f"人员字段 `{stripped}` 使用了方括号/黑体括号，应统一为中文圆括号 `（{role}：...）`")
                    )
                elif m_bare_credit and strict:
                    issues.append(
                        Issue("WARN", "BARE_CREDIT_LINE", rel_path, lineno, f"人员字段 `{role}` 未加中文圆括号包裹，应统一为 `（{role}：...）`")
                    )
                if role not in STANDARD_ROLES and strict:
                    issues.append(
                        Issue("WARN", "NON_STANDARD_ROLE", rel_path, lineno, f"非标准人员称谓 `{role}`（规范化输出应归一为制作人/调整人/修改人/重置人/审核人）")
                    )
                continue

            if CHANGELOG_PENDING_LEGACY_RE.match(line):
                seen_bracketed_changelog_in_header = True
                if strict:
                    issues.append(
                        Issue("WARN", "LEGACY_PENDING_DATE_LOG", rel_path, lineno, "旧遗留 `【日期待复核：...】` 无法考证年份，应统一改为 `【历史更新：...】` 并置于带日期日志之前")
                    )
                continue
            if CHANGELOG_UNDATED_STD_RE.match(line):
                seen_bracketed_changelog_in_header = True
                if last_changelog_date is not None and strict:
                    issues.append(
                        Issue("WARN", "CHANGELOG_DATE_ORDER", rel_path, lineno, "`【历史更新：...】` 无日期旧记录应排在有明确日期的 `【YY.M.D：...】` 记录之前")
                    )
                continue
            if CHANGELOG_ASCII_RE.match(line):
                seen_bracketed_changelog_in_header = True
                issues.append(
                    Issue("WARN", "ASCII_CHANGELOG_BRACKET", rel_path, lineno, "更新记录使用了半角方括号 `[]`，应统一为 `【日期：说明】`")
                )
                continue
            if CHANGELOG_STD_RE.match(line):
                seen_bracketed_changelog_in_header = True
                dt = parse_date_tuple(CHANGELOG_STD_RE.match(line).group(1))  # type: ignore[union-attr]
                if dt and last_changelog_date and dt < last_changelog_date and strict:
                    issues.append(Issue("WARN", "CHANGELOG_DATE_ORDER", rel_path, lineno, "首部更新日志日期逆序（标准要求从早到晚升序排列）"))
                if dt:
                    last_changelog_date = dt
                continue
            if CHANGELOG_CN_OR_BARE_RE.match(line):
                seen_bracketed_changelog_in_header = True
                if strict:
                    issues.append(
                        Issue("WARN", "UNBRACKETED_CHANGELOG", rel_path, lineno, f"非标准日期更新记录 `{stripped[:42]}`，应转为首部标准块 `【YY.M.D：说明】`")
                    )
                continue
            if CHANGELOG_UNDATED_RAW_RE.match(line):
                seen_bracketed_changelog_in_header = True
                if strict:
                    issues.append(
                        Issue("WARN", "UNDATED_CHANGELOG_RAW", rel_path, lineno, f"无完整日期的首部旧更新说明 `{stripped[:42]}`，存量规范化应转为 `【历史更新：...】`")
                    )
                continue

            # 若首部已出现 `【...】` 更新日志，但紧接着出现了无小节标题包裹的背景介绍散文，阅读器会把最后一条更新日志当成统领整段背景的大章节标题！
            if (
                strict
                and seen_bracketed_changelog_in_header
                and not warned_log_swallows_intro
                and not re.match(r"^\[[^\]：:]{2,12}\]$", stripped)
                and not DECORATIVE_DIVIDER_RE.match(line)
            ):
                warned_log_swallows_intro = True
                sec_hint = "【特质说明】" if is_trait_file else ("【能量池说明】" if is_energy_pool_file else "【背景】（或【职业说明】/【技能表说明】）")
                issues.append(
                    Issue(
                        "WARN",
                        "HEADER_LOG_SWALLOWS_INTRO",
                        rel_path,
                        lineno,
                        f"首部 `【...】` 更新日志下方直接跟随无标题背景/总则段落，会导致阅读器把更新日志渲染成背景大标题；请在日志块后补一行 `{sec_hint}` 小节头",
                    )
                )

        # 正文中间出现单卡独立制作人（如合集表《006】技艺技能表》L717 `制作人：姜雷`）
        if not in_header and strict and is_credit_line:
            issues.append(
                Issue(
                    "WARN",
                    "INLINE_CARD_CREATOR",
                    rel_path,
                    lineno,
                    f"正文中间出现单卡人员行 `{stripped[:36]}`：切勿移至第 1 行冒充全表作者，请紧贴对应单卡上方写成 `【说明：{stripped.strip('（）()[]［］【】')}】`",
                )
            )
            continue

        if not in_header and strict and CHANGELOG_CN_OR_BARE_RE.match(line) and not CHANGELOG_STD_RE.match(line):
            issues.append(
                Issue("WARN", "SCATTERED_CHANGELOG_LINE", rel_path, lineno, f"正文区疑似散落更新记录 `{stripped[:42]}`，请核实是否归入首部日志块")
            )

        if strict and DECORATIVE_DIVIDER_RE.match(line):
            issues.append(Issue("WARN", "DECORATIVE_DIVIDER_LINE", rel_path, lineno, "纯装饰分割横线建议移除，改用标准空行与 `【小节标题】` 分隔"))
            continue

        if strict and is_trait_file and card_count == 0 and BARE_TRAIT_GROWTH_RE.match(line):
            lvl = BARE_TRAIT_GROWTH_RE.match(line).group(1)  # type: ignore[union-attr]
            issues.append(
                Issue("WARN", "BARE_TRAIT_GROWTH_LINE", rel_path, lineno, f"特质主成长树 `{lvl}级：` 建议规范为 `【{lvl}】级：...`")
            )

        if strict and BARE_INLINE_POOL_RE.match(line):
            issues.append(
                Issue("WARN", "BARE_INLINE_ENERGY_POOL", rel_path, lineno, f"内嵌裸写能量池 `{stripped[:36]}`，建议规范为 `[能量池名称]：` 专卡")
            )

        # 空心方头括号 `〖...〗` 拦截（如《009】血莲化生》）
        if LENTICULAR_FIELD_RE.match(line):
            l_key = LENTICULAR_FIELD_RE.match(line).group(1).strip()  # type: ignore[union-attr]
            issues.append(
                Issue("ERROR", "LENTICULAR_BRACKET_FIELD", rel_path, lineno, f"使用了空心括号 `〖{l_key}〗`（会导致 0 卡识别），须改为标准半角 `[{l_key}]：`")
            )
            continue

        m_mixed = MIXED_BRACKET_FIELD_RE.match(line)
        if m_mixed:
            key = (m_mixed.group(1) or m_mixed.group(2)).strip()
            issues.append(
                Issue("ERROR", "MIXED_BRACKET_FIELD", rel_path, lineno, f"字段左右括号半全角混搭 `{stripped[:32]}`，会导致字段解析断链，须改为半角 `[{key}]：`")
            )
            continue

        if BROKEN_LEFT_BRACKET_RE.match(line):
            issues.append(
                Issue("ERROR", "BROKEN_CARD_LEFT_BRACKET", rel_path, lineno, f"卡片起始字段漏写左括号，会导致整张卡被前文吞并: `{stripped[:40]}`")
            )
            starter_info: tuple[str, str] | None = (BROKEN_LEFT_BRACKET_RE.match(line).group(1).strip(), "std")  # type: ignore[union-attr]
        else:
            starter_info = match_any_card_starter(line)

        if COLON_INSIDE_BRACKET_RE.match(line):
            issues.append(
                Issue("ERROR", "COLON_INSIDE_BRACKET", rel_path, lineno, f"括号内部误含冒号 `{stripped[:40]}`（若为小节标题请去掉冒号写成 `【小节名】`）")
            )
            continue
        if MISSING_COLON_STARTER_RE.match(line):
            issues.append(
                Issue("ERROR", "MISSING_STARTER_COLON", rel_path, lineno, f"卡片起始字段后缺少冒号 `{stripped[:40]}`（会导致整张卡无法识别），须改为 `[能力名称]：...`")
            )
            continue
        if JIYI_RULE_BAD_RE.match(line):
            issues.append(
                Issue("ERROR", "INVALID_JIYI_RULE_FORMAT", rel_path, lineno, "`技艺规则` 必须固定写成结构字段 `[技艺规则]：`，不得使用 `【技艺规则】` 或漏写冒号")
            )
            continue

        if starter_info:
            starter_key, bracket_kind = starter_info
            card_count += 1

            if strict and bracket_kind != "std":
                issues.append(
                    Issue("WARN", "NON_ASCII_CARD_STARTER", rel_path, lineno, f"卡片起始头使用了非半角括号 `{stripped[:24]}`，应统一为 `[{starter_key}]：`")
                )
            if strict and starter_key == "名称":
                target_key = "能量池名称" if is_energy_pool_file else "能力名称（或对应专用 XX名称）"
                issues.append(
                    Issue("WARN", "BARE_NAME_STARTER", rel_path, lineno, f"使用了过于宽泛的 `[名称]：`，建议明确为 `[{target_key}]：`")
                )

            if idx > first_non_empty_idx + 1:
                prev_raw = lines[idx - 1]
                m_prev_field = EXPLICIT_FIELD_STD_RE.match(prev_raw) or HEAVY_BRACKET_FIELD_RE.match(prev_raw)
                if m_prev_field and not match_any_card_starter(prev_raw):
                    prev_key = m_prev_field.group(1).strip()
                    above_prev_blank = (idx - 2 < 0) or (not lines[idx - 2].strip())
                    if prev_key in PRE_CARD_FLOATING_KEYS or above_prev_blank:
                        issues.append(
                            Issue(
                                "ERROR" if strict else "WARN",
                                "PRE_CARD_FLOATING_FIELD",
                                rel_path,
                                lineno - 1,
                                f"属性字段 `[{prev_key}]` 写在卡片起始头 `[{starter_key}]` 正上方，会被上一张卡吞并；请移入本卡内部",
                            )
                        )

                if strict:
                    k = idx - 1
                    stacked_brackets: list[tuple[int, str]] = []
                    while k > first_non_empty_idx:
                        cand = lines[k].strip()
                        m_sec = SECTION_BRACKET_RE.match(cand)
                        if m_sec and not CHANGELOG_STD_RE.match(cand) and not CHANGELOG_UNDATED_STD_RE.match(cand):
                            stacked_brackets.append((k + 1, m_sec.group(1).strip()))
                            k -= 1
                        else:
                            break
                    stacked_brackets.reverse()

                    for pos_idx, (b_lineno, inner) in enumerate(stacked_brackets):
                        if is_prefixed_card_note(inner):
                            continue
                        if is_known_section_title(inner):
                            if len(stacked_brackets) >= 2 and pos_idx == 0:
                                issues.append(
                                    Issue(
                                        "WARN",
                                        "SECTION_NOT_SEPARATED",
                                        rel_path,
                                        b_lineno,
                                        f"小节标题 `【{inner}】` 与下方卡片/备注紧贴无空行，建议在小节标题后空一行",
                                    )
                                )
                            continue
                        if UNLOCK_OR_SCOPE_HINT_RE.search(inner) and not MOD_NOTE_HINT_RE.search(inner):
                            issues.append(
                                Issue(
                                    "WARN",
                                    "CARD_TOP_UNLOCK_OR_SCOPE_RULE",
                                    rel_path,
                                    b_lineno,
                                    f"卡顶 `【{inner[:36]}...】` 非版本改动：升级来源/普通说明请改为 `【说明：...】`；整节限制留在小节头下空一行；仅极其明确的硬性兑换门槛才可改为 `[使用要求]：`（涉及购买要求底层规则，改后须人工复核）",
                                )
                            )
                        elif MOD_NOTE_HINT_RE.search(inner):
                            issues.append(
                                Issue(
                                    "WARN",
                                    "UNPREFIXED_MOD_NOTE",
                                    rel_path,
                                    b_lineno,
                                    f"卡顶修改说明建议加前缀写成 `【能力修改：{inner}】`",
                                )
                            )
                        else:
                            issues.append(
                                Issue(
                                    "WARN",
                                    "AMBIGUOUS_CARD_TOP_BRACKET",
                                    rel_path,
                                    b_lineno,
                                    f"卡顶未分类 `【{inner[:36]}】`：请按语义区分小节标题（后空一行）、重影旧标题（清理）、版本改动（`【能力修改：...】`）或单卡说明（`【说明：...】`）",
                                )
                            )
            continue

        if SECTION_BRACKET_RE.match(line):
            continue

        if FULLWIDTH_FIELD_RE.match(line):
            fw_key = FULLWIDTH_FIELD_RE.match(line).group(1).strip()  # type: ignore[union-attr]
            issues.append(
                Issue("WARN", "FULLWIDTH_BRACKET_FIELD", rel_path, lineno, f"使用了全角方括号 `［{fw_key}］`，应规范为半角 `[{fw_key}]：`")
            )
            continue

        # 全面拦截一切行首 `【XXX】：...`（前端 adaptive.ts 会将一切行首 `【XXX】：` 拆分为独立卡片字段框！）
        m_heavy_field = HEAVY_BRACKET_FIELD_RE.match(line)
        if m_heavy_field:
            h_key = m_heavy_field.group(1).strip()
            h_val = m_heavy_field.group(2).strip()
            if strict:
                is_domain_card_field = bool(
                    h_key in COMMON_CARD_FIELDS_IN_HEAVY_BRACKETS
                    or h_key.startswith("能力形容")
                    or (
                        len(h_key) <= 8
                        and re.search(r"(?:简介|介绍|效果|形容|描述|类型|距离|消耗|冷却|时间|要求|限制|条件|规则|价格|部位|要素|详解|上限|奖励)$", h_key)
                    )
                )
                if is_domain_card_field:
                    target_key = NON_STD_FIELD_ALIASES.get(h_key, h_key)
                    issues.append(
                        Issue("WARN", "HEAVY_BRACKET_CARD_FIELD", rel_path, lineno, f"卡片字段使用了黑体括号 `【{h_key}】：`，应规范为半角 `[{target_key}]：`")
                    )
                elif is_known_section_title(h_key):
                    hint = f"小节标题 `【{h_key}】：` 多打了冒号（会被前卡当成字段吞并），请去掉冒号改为独立行 `【{h_key}】`"
                    if h_val:
                        hint += "，并将冒号后的导语移至下一行"
                    issues.append(Issue("WARN", "SECTION_WITH_TRAILING_COLON", rel_path, lineno, hint))
                else:
                    issues.append(
                        Issue(
                            "WARN",
                            "HEAVY_BRACKET_COLON_ORPHAN",
                            rel_path,
                            lineno,
                            f"行首 `【{h_key}】：` 会被前端拆成孤儿字段框！若属本卡衍生词条请收进所属卡 `[补充说明]：` 并改为 `「{h_key}」：`；若属子分项请改为 `{h_key}：`；若属重复旧卡头请清理",
                        )
                    )
            continue

        m_std_field = EXPLICIT_FIELD_STD_RE.match(line)
        if m_std_field:
            f_key = m_std_field.group(1).strip()
            f_val = m_std_field.group(2).strip()

            if strict and f_key in NON_STD_FIELD_ALIASES:
                next_is_starter = idx + 1 < len(lines) and bool(match_any_card_starter(lines[idx + 1]))
                if not next_is_starter:
                    issues.append(
                        Issue(
                            "WARN",
                            "NON_STANDARD_FIELD_ALIAS",
                            rel_path,
                            lineno,
                            f"非标字段别名 `[{f_key}]：`，建议统一为模板标准字段 `[{NON_STD_FIELD_ALIASES[f_key]}]：`",
                        )
                    )

            if f_key in OPTIONAL_CARD_FIELDS and not f_val:
                next_line = lines[idx + 1].strip() if idx + 1 < len(lines) else ""
                next_is_boundary = (
                    not next_line
                    or is_any_field_or_starter(next_line)
                    or bool(SECTION_BRACKET_RE.match(next_line))
                )
                if next_is_boundary:
                    issues.append(
                        Issue("WARN", "EMPTY_OPTIONAL_FIELD", rel_path, lineno, f"残留空的可选字段 `[{f_key}]：`，无内容时应直接删除该行")
                    )

    if strict and not is_public_building and not is_annotation_only and card_count == 0:
        issues.append(
            Issue(
                "ERROR",
                "ZERO_CARDS_IN_RESOURCE",
                rel_path,
                1,
                "全篇识别出 0 张 `[XX名称]` 卡片（常见于裸写字段、漏冒号、空心括号或 1.~N. 列表老档案），规范化时须重构为标准卡片",
            )
        )

    if strict:
        issues.extend(_check_blank_line_hygiene(lines, first_non_empty_idx, rel_path))

    if new_submission and not seen_any_credit:
        issues.append(
            Issue(
                "WARN",
                "MISSING_CREATOR_CREDIT",
                rel_path,
                1,
                "新投稿首部缺少制作人/审核人署名（注：存量老资源缺署名直接留空，不报此项）",
            )
        )

    return issues, card_count


def collect_git_changed_files() -> list[Path]:
    try:
        proc = subprocess.run(
            ["git", "status", "--porcelain", "-z", "--", "序列库", "荣誉室"],
            cwd=REPO_ROOT,
            capture_output=True,
            timeout=10,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return []
    raw_entries = proc.stdout.decode("utf-8", errors="replace").split("\0")
    paths: list[Path] = []
    i = 0
    while i < len(raw_entries):
        entry = raw_entries[i]
        if not entry or len(entry) < 4:
            i += 1
            continue
        status = entry[:2]
        rel = entry[3:]
        if "R" in status or "C" in status:
            i += 1
        candidate = REPO_ROOT / rel
        if candidate.is_file():
            paths.append(candidate)
        i += 1
    return sorted(set(paths))


def iter_target_files(targets: list[str], use_all: bool, git_changed: bool) -> list[Path]:
    if git_changed:
        return collect_git_changed_files()
    roots = [REPO_ROOT / "序列库"] if use_all or not targets else [Path(t).resolve() for t in targets]
    collected: list[Path] = []
    for root in roots:
        if root.is_file():
            collected.append(root)
        elif root.is_dir():
            collected.extend(sorted(p for p in root.rglob("*") if p.is_file()))
    return collected


def main() -> int:
    parser = argparse.ArgumentParser(description="宏观界域序列库资源只读探针与校验工具 v3（AI Harness 组件）")
    parser.add_argument("targets", nargs="*", help="待检查的文件或目录路径")
    parser.add_argument("--all", action="store_true", help="扫描整个正式 `序列库/` 目录")
    parser.add_argument("--git-changed", action="store_true", help="仅扫描当前 Git 工作区在 序列库/ 与 荣誉室/ 下有改动的文件")
    parser.add_argument("--strict", action="store_true", help="启用 AI 规范化探针 / 严格格式门禁模式")
    parser.add_argument("--new-submission", action="store_true", help="新过审投稿审稿模式（额外检查是否漏填制作人/审核人）")
    parser.add_argument("--summary", action="store_true", help="按问题代码输出分类统计摘要")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出结果")
    args = parser.parse_args()

    files = iter_target_files(args.targets, args.all, args.git_changed)
    all_issues: list[Issue] = []
    txt_checked = 0
    total_cards = 0

    for file_path in files:
        try:
            rel_path = str(file_path.relative_to(REPO_ROOT))
        except ValueError:
            rel_path = str(file_path)

        if ".obsidian" in file_path.parts:
            continue
        if file_path.name in DIRTY_NAMES:
            all_issues.append(Issue("ERROR", "DIRTY_SYSTEM_FILE", rel_path, 0, f"系统脏文件混入资源目录: {file_path.name}"))
            continue
        if file_path.suffix.lower() in (".doc", ".docx"):
            all_issues.append(
                Issue("WARN", "NON_TXT_RESOURCE", rel_path, 0, f"非 TXT 资源 ({file_path.suffix})：Web 在线站检索暂不解析二进制 Word")
            )
            continue
        if file_path.suffix.lower() != ".txt":
            continue

        txt_checked += 1
        file_issues, cards = check_txt_file(file_path, rel_path, strict=args.strict, new_submission=args.new_submission)
        all_issues.extend(file_issues)
        total_cards += cards

    errors = [i for i in all_issues if i.severity == "ERROR"]
    warns = [i for i in all_issues if i.severity == "WARN"]
    by_code = dict(Counter(i.code for i in all_issues).most_common())

    if args.json:
        print(
            json.dumps(
                {
                    "checked_txt_files": txt_checked,
                    "matched_cards": total_cards,
                    "error_count": len(errors),
                    "warn_count": len(warns),
                    "counts_by_code": by_code,
                    "issues": [asdict(i) for i in all_issues],
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        print(f"扫描完成：{txt_checked} 份 TXT 资源，泛匹配识别 {total_cards} 张 `[XX名称]` 卡片 | ERROR: {len(errors)} | WARN: {len(warns)}")
        if args.summary and by_code:
            print("按问题类型统计：")
            for code, cnt in by_code.items():
                print(f"  - {code}: {cnt}")
        for item in all_issues[:80]:
            loc = f"{item.path}:{item.line}" if item.line else item.path
            print(f"[{item.severity}] {item.code} @ {loc} -> {item.message}")
        if len(all_issues) > 80:
            print(f"... 其余 {len(all_issues) - 80} 条提示已省略（可传 --json 或缩小目标范围查看完整列表）")

    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
