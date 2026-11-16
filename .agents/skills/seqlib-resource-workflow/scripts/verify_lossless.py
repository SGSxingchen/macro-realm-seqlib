#!/usr/bin/env python3
"""宏观界域强化序列库 · AI 改稿无损对账与语义升格守门员 v3（AI Harness 组件 2/2）。

核心三级门禁：
1. 红牌直接拦截（`passed = False`）：
   - 数值/百分比/骰子公式被篡改或彻底删没（`count >= 1 -> 0` 或无中生有新数值）；
   - 原稿实质规则词句丢失或被改字；
   - 全文署名者（含合集表第 700+ 行的单卡内嵌制作人）丢失或凭空编造新署名。
2. 黄牌语义/去重复核预警（`needs_semantic_review = True`）：
   - 「字段语义桶看门狗」：如将非 `[使用要求]` 文本升格进 `[使用要求]`（涉及购买门槛底层规则）、
     跨字段搬移风味简介与规则效果、或从正文拆出新的强结算元数据（如 `[冷却时间]`）；
   - 「重影清理去重核验」：如《018】役亡师》清理了卡顶与卡内 100% 重复的旧版重影摘要
     （原稿出现 >=2 次、改后完整保留 >=1 次），自动挂黄牌生成签审单供人工一眼确认。
3. 绿牌纯格式通过（`passed = True` 且 `needs_semantic_review = False`）：
   - 空心括号/无冒号/裸写字段成卡、1.~19. 列表成卡、`「子实体」：` 收归 `[补充说明]`、
     `【历史更新：...】` / `【能力修改：...】` / `【说明：...】` 前缀化、首部顺序归位等零规则风险清洗。
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
from datetime import datetime, timezone
import hashlib
import json
import re
import secrets
import subprocess
import sys
import time
from dataclasses import asdict, dataclass, field
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


REVIEW_DIR = REPO_ROOT / "web" / "normalization_reviews"

COMPAT_ROLES = (
    "制作人", "制作者", "投稿人", "作者", "二次制作", "调整人", "修改人", "重置人",
    "审核人", "审核者", "审核", "复查人", "文本优化", "更新人", "翻修人", "协力",
)
ROLE_GROUP = "|".join(sorted(COMPAT_ROLES, key=len, reverse=True))

# 支持首部及正文中间的单卡作者（包括 `【说明：制作人：姜雷】`）
CREDIT_LINE_RE = re.compile(
    rf"^[ \t]*(?:【说明[：:]\s*)?[（(\[［【〖]?\s*(?:{ROLE_GROUP})\s*[：:]\s*([^）)\]］】〗\r\n]+?)[）)\]］】〗]*[ \t]*$"
)
DECORATIVE_DIVIDER_RE = re.compile(r"^[ \t]*[—―─\-=_=~～]{6,}[ \t]*$")

CHANGELOG_PREFIX_RE = re.compile(
    r"^[ \t]*[\[【]?\s*((?:20\d{2}|\d{2})年\d{1,2}月(?:\d{1,2}日)?|\d{2,4}[./-]\d{1,2}(?:[./-]\d{1,2})?)\s*(?:更新|修改|调整|重置|补充)?\s*[：:](.+?)(?:[\]】])?([ \t]*[（(].*[）)])?[ \t]*$"
)

KNOWN_BARE_FIELDS = (
    "能力名称", "技能名称", "称号名称", "道具名称", "能量池名称", "商品名称", "奖励名称",
    "义体名称", "模块名称", "单位名称", "符文名称", "天赋名称", "仪式名称", "职阶名称",
    "能力简介", "技能简介", "称号简介", "能力介绍", "技能介绍", "介绍", "称号描述", "称号形容",
    "道具简介", "道具介绍", "道具描述", "能量池简介", "能量池介绍", "商品简介", "奖励简介", "模块简介", "单位简介", "符文简介",
    "能力效果", "技能效果", "称号效果", "商品效果", "奖励效果", "道具效果", "能量池效果", "模块效果", "单位效果", "符文效果", "天赋效果", "仪式效果", "升级效果", "进阶效果", "强化效果", "效果",
    "初级效果", "中级效果", "高级效果", "终极效果",
    "能力形容", "技能形容", "释放类型", "打击类型", "伤害类型", "最大距离", "攻击距离",
    "消耗能量", "能量消耗", "能力消耗", "技能消耗", "冷却时间", "技能冷却", "能力冷却",
    "持续时间", "技能持续", "使用要求", "能力要求", "技能要求",
    "称号解锁条件", "称号条件", "称号要求", "获得条件", "获取条件", "解锁条件", "激活条件", "进阶条件",
    "补充说明", "效果补充", "能力备注", "能力注释", "备注",
    "技艺规则", "冥想难度", "习得难度", "传授难度", "学习难度", "冥想/传授难度", "习得/传授难度", "学习/传授难度",
    "价格", "兑换价格", "特殊", "能量池",
    "更新", "更新说明", "修改内容", "调整内容", "重置内容", "日期待复核",
)
BARE_FIELD_GROUP = "|".join(sorted(KNOWN_BARE_FIELDS, key=len, reverse=True))

# 兼容 `[]` / `【】` / `［］` / `〖〗`（有冒号或无冒号卡头）、破损左括号及裸写字段
EXPLICIT_OR_BROKEN_FIELD_RE = re.compile(
    rf"^[ \t]*(?:"
    rf"[\[［【〖]([^\]］】〗\r\n]{{1,32}})[\]］】〗][ \t]*[：:]"
    rf"|[\[［【〖]((?:{BARE_FIELD_GROUP}|[A-ES]级能力效果))[\]］】〗][ \t]*"
    rf"|([\u4e00-\u9fff]{{1,12}}名称)[\]］】〗][ \t]*[：:]"
    rf"|({BARE_FIELD_GROUP})[：:]"
    rf")[ \t]*(.*)$"
)
# 同卡内 `能力效果1：` / `技能效果2：` 与标准 `效果1：` / `效果2：` 等价归一
SUB_EFFECT_PREFIX_RE = re.compile(r"^(?:能力|技能)?效果(\d+[：:])[ \t]*")
TIER_SHORT_PREFIX_RE = re.compile(r"^(初级|中级|高级|终极)[：:][ \t]*")
TRAIT_LEVEL_HEAD_RE = re.compile(r"^[ \t]*【?(EX|[A-ES])】?级(?:别)?[：:\s][ \t]*")
SINGLE_HEAVY_WRAP_RE = re.compile(r"^[ \t]*【(?:(?:能力修改|说明|历史更新|日期待复核|调整|修改|更新)[：:])?\s*(.+?)】[ \t]*$")
CLAUSE_SPLIT_RE = re.compile(r"[,，.．。;；:：!！?？（）()\[\]【】\s]+")

ALLOWED_ADDED_SECTION_HEADERS = {
    "建筑简介", "建筑说明", "特殊规则", "特质说明", "职业说明", "技能表说明", "能量池说明",
    "背景", "升级树", "限制", "称号", "职业称号", "称号技能", "职业技能", "专属能量池", "能量池",
    "本事", "神通", "根器", "扩充技能组", "忠诚派", "混沌派", "改造手术",
    "E级", "D级", "C级", "B级", "A级", "S级", "EX级",
    "衍生状态说明", "衍生词条说明", "词条说明", "说明", "机制说明", "基础性能", "可成长能力",
    "特定高级能力", "特殊高级能力", "初级能力", "中级能力", "高级能力", "终极能力", "通用规则",
    "食品列表", "饮品列表", "商品列表", "兑换列表",
}

GATE_FIELDS = {
    "使用要求", "使用限制", "能力要求", "技能要求",
    "兑换条件", "解锁条件", "称号解锁条件", "称号条件", "称号要求",
    "获得条件", "获取条件", "激活条件", "进阶条件", "前置条件", "学习条件", "购买条件",
}
JIYI_FIELDS = {
    "技艺规则", "技艺条件", "技艺难度",
    "冥想难度", "习得难度", "传授难度", "学习难度", "冥想/传授难度", "习得/传授难度", "学习/传授难度",
}
MECHANIC_META_FIELDS = {
    "释放类型", "打击类型", "伤害类型", "最大距离", "攻击距离", "消耗能量", "能力消耗", "技能消耗", "能量消耗",
    "冷却时间", "技能冷却", "能力冷却", "持续时间", "技能持续", "技能类型", "能力类型", "占用部位", "单位限制", "孵化消耗",
}
PRICE_FIELDS = {"价格", "兑换价格", "兑换消耗", "售价"}
FLAVOR_FIELDS = {
    "简介", "能力简介", "技能简介", "技艺简介", "称号简介", "能力介绍", "技能介绍", "介绍",
    "称号描述", "称号形容", "道具简介", "道具介绍", "道具描述", "能量池简介", "能量池介绍",
    "商品简介", "奖励简介", "模块简介", "单位简介", "符文简介",
    "能力形容", "技能形容", "技艺形容", "形容", "能力描述",
}

NUMERIC_TOKEN_RE = re.compile(r"\d+(?:\.\d+)?(?:%|‰|[dD]\d+|[ⁿ]+)?")
PUNCT_AND_SPACE_RE = re.compile(r"[\s,，.．。;；:：!！?？、·`'\"“”‘’「」『』()（）\[\]［］【】〖〗{}<>《》/\\—―\-]+")


@dataclass
class SemanticUnit:
    lineno: int
    raw_payload: str
    compact_key: str
    bucket: str
    field_key: str


@dataclass
class LosslessReport:
    path: str
    passed: bool
    needs_semantic_review: bool = False
    missing_clauses: list[str] = field(default_factory=list)
    numeric_mismatches: list[str] = field(default_factory=list)
    credit_mismatches: list[str] = field(default_factory=list)
    semantic_changes: list[str] = field(default_factory=list)
    dedup_changes: list[str] = field(default_factory=list)
    added_clauses: list[str] = field(default_factory=list)
    added_sections: list[str] = field(default_factory=list)


def classify_field_bucket(field_key: str) -> str:
    fk = field_key.strip()
    if not fk:
        return "BODY_OR_NOTE"
    if fk == "能量池" or fk.endswith("名称") or fk == "名称":
        return "CARD_NAME"
    if fk in GATE_FIELDS:
        return "GATE"
    if fk in JIYI_FIELDS:
        return "JIYI"
    if fk in MECHANIC_META_FIELDS:
        return "MECHANIC_META"
    if fk in PRICE_FIELDS:
        return "PRICE"
    if fk in FLAVOR_FIELDS or fk.endswith("简介") or fk.startswith("能力形容") or fk.endswith("形容"):
        return "FLAVOR"
    return "BODY_OR_NOTE"


def normalize_date_key(raw_date: str) -> str:
    m_cn = re.match(r"^(\d{2,4})年(\d{1,2})月(?:(\d{1,2})日)?$", raw_date.strip())
    if m_cn:
        y = int(m_cn.group(1))
        y = y + 2000 if y < 100 else y
        m = int(m_cn.group(2))
        d = int(m_cn.group(3)) if m_cn.group(3) else 1
        return f"DATE:{y:04d}-{m:02d}-{d:02d}"
    parts = re.split(r"[./-]", raw_date.strip())
    try:
        nums = [int(p) for p in parts]
    except ValueError:
        return raw_date.strip()
    y = nums[0] + 2000 if nums[0] < 100 else nums[0]
    m = nums[1] if len(nums) > 1 else 1
    d = nums[2] if len(nums) > 2 else 1
    return f"DATE:{y:04d}-{m:02d}-{d:02d}"


def compact_text(s: str) -> str:
    return PUNCT_AND_SPACE_RE.sub("", s).lower()


def extract_semantic_units(text: str) -> tuple[list[SemanticUnit], Counter[str], set[str]]:
    lines = text.splitlines()
    units: list[SemanticUnit] = []
    num_counter: Counter[str] = Counter()
    authors: set[str] = set()

    first_non_empty = next((i for i, l in enumerate(lines) if l.strip()), -1)
    title_stem_without_num = ""
    if first_non_empty >= 0:
        first_line = lines[first_non_empty].strip()
        title_stem_without_num = re.sub(r"^\d{3}】", "", first_line).strip()

    current_bucket = "BODY_OR_NOTE"
    current_field_key = ""

    for idx, raw_line in enumerate(lines):
        lineno = idx + 1
        stripped = raw_line.strip()
        if not stripped or DECORATIVE_DIVIDER_RE.match(stripped):
            continue

        # 跳过首部完全重复首行标题的冗余副标题（含 `能量池（灵气）`、`合欢弟子【职业】（神秘侧）` 等复读）
        if idx > first_non_empty and idx <= first_non_empty + 6 and title_stem_without_num:
            if stripped.strip("【】[]（）() ") == title_stem_without_num.strip("【】[]（）() "):
                continue
            unwrapped_sub = re.sub(
                r"[（(\[【]\s*(?:职业|技能表|特质|特质改造|能量池|公共建筑|战技侧|神秘侧|科技侧|特殊侧)[）)\]】]",
                "",
                stripped,
            ).strip("【】[]（）() ")
            if unwrapped_sub == title_stem_without_num.strip("【】[]（）() "):
                continue

        # 1. 全文人员署名行扫描（覆盖首部与合集表第 700+ 行单卡署名）
        m_credit = CREDIT_LINE_RE.match(stripped)
        if m_credit:
            names_raw = m_credit.group(1)
            for name in re.split(r"[、,，/／&\s]+", names_raw):
                cleaned_name = name.strip().strip("（([［【〖）)]］】〗").strip()
                if cleaned_name:
                    authors.add(cleaned_name)
            continue

        # 2. 更新日志行
        m_log = CHANGELOG_PREFIX_RE.match(stripped)
        if m_log and idx <= max(first_non_empty + 35, 40):
            date_token = normalize_date_key(m_log.group(1))
            num_counter[date_token] += 1
            body = m_log.group(2).strip()
            tail_note = (m_log.group(3) or "").strip()
            full_log_payload = f"{body}{tail_note}"
            for tok in NUMERIC_TOKEN_RE.findall(full_log_payload):
                num_counter[tok.lower()] += 1
            for sub in re.split(r"[;；]", full_log_payload):
                ck = compact_text(sub)
                if ck:
                    units.append(SemanticUnit(lineno, sub.strip(), ck, "CHANGELOG", "更新记录"))
            continue

        # 3. 常规行
        payload = stripped
        unit_bucket = current_bucket
        unit_field_key = current_field_key

        m_trait = TRAIT_LEVEL_HEAD_RE.match(payload)
        if m_trait:
            lvl = m_trait.group(1)
            rest = payload[m_trait.end():].strip()
            payload = f"{lvl}级：{rest}"
            unit_bucket = "BODY_OR_NOTE"
            unit_field_key = "成长树"
            current_bucket = "BODY_OR_NOTE"
            current_field_key = ""
        else:
            m_field = EXPLICIT_OR_BROKEN_FIELD_RE.match(payload)
            if m_field:
                fkey = (m_field.group(1) or m_field.group(2) or m_field.group(3) or m_field.group(4) or "").strip()
                rest = (m_field.group(5) or "").strip()
                # 注意：若行首是 `【子实体】：buff...`（非已知结构字段名）转为 `「子实体」：buff...`，保留 `子实体` 名本身！
                if (
                    m_field.group(1)
                    and stripped.startswith("【")
                    and classify_field_bucket(fkey) == "BODY_OR_NOTE"
                    and fkey not in KNOWN_BARE_FIELDS
                    and not fkey.startswith("效果")
                    and not fkey.startswith("能力效果")
                ):
                    payload = f"{fkey}：{rest}" if rest else fkey
                    unit_bucket = current_bucket
                    unit_field_key = current_field_key
                else:
                    unit_bucket = classify_field_bucket(fkey)
                    unit_field_key = fkey
                    current_bucket = "BODY_OR_NOTE" if unit_bucket == "CARD_NAME" else unit_bucket
                    current_field_key = "" if unit_bucket == "CARD_NAME" else fkey
                    payload = rest
            else:
                m_wrap = SINGLE_HEAVY_WRAP_RE.match(payload)
                if m_wrap:
                    payload = m_wrap.group(1).strip()
                    # 若整段 `【设定背景...】` 被重构进了 `[能力简介]：`（FLAVOR），允许记录原桶与后续对账
                    unit_bucket = "BODY_OR_NOTE"
                    unit_field_key = ""
                    current_bucket = "BODY_OR_NOTE"
                    current_field_key = ""

        payload = SUB_EFFECT_PREFIX_RE.sub(r"效果\1", payload)
        m_tier_short = TIER_SHORT_PREFIX_RE.match(payload)
        if m_tier_short:
            payload = payload[m_tier_short.end():].strip()
        if not payload:
            continue

        for tok in NUMERIC_TOKEN_RE.findall(payload):
            num_counter[tok.lower()] += 1

        ck = compact_text(payload)
        if ck:
            units.append(SemanticUnit(lineno, payload, ck, unit_bucket, unit_field_key))

    return units, num_counter, authors


def compare_lossless(rel_path: str, before_text: str, after_text: str) -> LosslessReport:
    before_units, before_nums, before_authors = extract_semantic_units(before_text)
    after_units, after_nums, after_authors = extract_semantic_units(after_text)

    after_key_pool = Counter(u.compact_key for u in after_units)
    before_key_pool = Counter(u.compact_key for u in before_units)

    after_full_compact = "".join(u.compact_key for u in after_units)
    before_full_compact = "".join(u.compact_key for u in before_units)

    before_buckets_by_key: dict[str, set[str]] = defaultdict(set)
    for u in before_units:
        before_buckets_by_key[u.compact_key].add(u.bucket)

    # 1. 实质词句零丢失对账（精准区分：1:1拆行重排 vs 重影清理 `N>=2 -> M>=1` vs 真丢失 `-> 0`）
    missing_clauses: list[str] = []
    dedup_changes: list[str] = []
    dedup_dropped_nums: Counter[str] = Counter()
    temp_after_pool = after_key_pool.copy()

    def _check_subclauses_preserved(raw_line: str) -> tuple[bool, bool]:
        """返回 (all_preserved_in_after, any_reduced_duplicate)。"""
        parts = [p for p in CLAUSE_SPLIT_RE.split(raw_line) if compact_text(p)]
        if not parts:
            return True, False
        any_dup_reduced = False
        for part in parts:
            ck = compact_text(part)
            a_occ = after_full_compact.count(ck)
            if a_occ == 0:
                return False, False
            b_occ = before_full_compact.count(ck)
            if b_occ >= 2 and a_occ < b_occ:
                any_dup_reduced = True
        return True, any_dup_reduced

    for u in before_units:
        if temp_after_pool[u.compact_key] > 0:
            temp_after_pool[u.compact_key] -= 1
        elif after_key_pool[u.compact_key] >= 1 and before_key_pool[u.compact_key] >= 2:
            dedup_changes.append(f"原稿 L{u.lineno} 重复重影已去重（改后仍完整保留）: `{u.raw_payload[:52]}`")
            for tok in NUMERIC_TOKEN_RE.findall(u.raw_payload):
                dedup_dropped_nums[tok.lower()] += 1
        else:
            all_ok, is_dup_reduction = _check_subclauses_preserved(u.raw_payload)
            if all_ok and is_dup_reduction:
                dedup_changes.append(f"原稿 L{u.lineno} 跨字段重复重影已去重（各子句改后均完整保留）: `{u.raw_payload[:52]}`")
                for tok in NUMERIC_TOKEN_RE.findall(u.raw_payload):
                    dedup_dropped_nums[tok.lower()] += 1
            elif all_ok or u.compact_key in after_full_compact:
                continue
            else:
                missing_clauses.append(f"原稿 L{u.lineno} 内容疑似丢失或被改字: `{u.raw_payload[:60]}`")

    # 2. 新增词句与「字段语义桶看门狗」
    added_clauses: list[str] = []
    added_sections: list[str] = []
    semantic_changes: list[str] = []
    temp_before_pool = before_key_pool.copy()

    for u in after_units:
        orig_buckets = before_buckets_by_key.get(u.compact_key, set())
        if orig_buckets and u.bucket not in orig_buckets:
            if u.bucket == "GATE":
                semantic_changes.append(
                    f"改后 L{u.lineno} 语义升格为 `[{u.field_key}]`（涉及购买要求底层规则）: `{u.raw_payload[:50]}`"
                )
            elif u.bucket == "MECHANIC_META":
                semantic_changes.append(
                    f"改后 L{u.lineno} 拆入强结算元数据 `[{u.field_key}]`: `{u.raw_payload[:50]}`"
                )
            elif "FLAVOR" in orig_buckets or u.bucket == "FLAVOR":
                semantic_changes.append(
                    f"改后 L{u.lineno} 设定背景与规则正文跨字段归类（`[{u.field_key or '正文'}]`）: `{u.raw_payload[:50]}`"
                )
        elif not orig_buckets and u.bucket in ("GATE", "MECHANIC_META") and u.compact_key in before_full_compact:
            semantic_changes.append(
                f"改后 L{u.lineno} 从原正文中拆分新建了强规则字段 `[{u.field_key}]`: `{u.raw_payload[:50]}`"
            )

        if temp_before_pool[u.compact_key] > 0:
            temp_before_pool[u.compact_key] -= 1
        elif u.compact_key in before_full_compact:
            continue
        elif u.raw_payload in ALLOWED_ADDED_SECTION_HEADERS:
            added_sections.append(f"改后 L{u.lineno} 新增结构小节头: `【{u.raw_payload}】`")
        else:
            all_in_before = all(
                compact_text(p) in before_full_compact
                for p in CLAUSE_SPLIT_RE.split(u.raw_payload)
                if compact_text(p)
            )
            if all_in_before and compact_text(u.raw_payload):
                continue
            added_clauses.append(f"改后 L{u.lineno} 出现原稿不存在的新增词句: `{u.raw_payload[:60]}`")

    # 3. 数值/百分比/骰子多重集对账（仅允许被证实去重的重影行所携带的精确数字差额转为黄牌，其余一切数字增减一律红牌拦截！）
    numeric_mismatches: list[str] = []
    all_num_keys = sorted(set(before_nums.keys()) | set(after_nums.keys()))
    for nk in all_num_keys:
        b_cnt = before_nums[nk]
        a_cnt = after_nums[nk]
        if b_cnt != a_cnt:
            diff = b_cnt - a_cnt
            if diff > 0 and a_cnt >= 1 and dedup_dropped_nums[nk] >= diff:
                dedup_changes.append(f"随重影清理同步去重的数值 `{nk}`: 原稿 {b_cnt} 处 -> 改后 {a_cnt} 处")
            else:
                numeric_mismatches.append(f"数值/日期指纹 `{nk}` 数量变化: 原稿 {b_cnt} 处 -> 改后 {a_cnt} 处")

    # 4. 署名者集合对账
    credit_mismatches: list[str] = []
    lost_authors = sorted(before_authors - after_authors)
    invented_authors = sorted(after_authors - before_authors)
    if lost_authors:
        credit_mismatches.append(f"丢失原稿署名人员: {', '.join(lost_authors)}")
    if invented_authors:
        credit_mismatches.append(f"出现原稿不存在的署名人员: {', '.join(invented_authors)}")

    passed = not (missing_clauses or numeric_mismatches or credit_mismatches)
    needs_semantic_review = bool(semantic_changes or dedup_changes or added_clauses)
    return LosslessReport(
        path=rel_path,
        passed=passed,
        needs_semantic_review=needs_semantic_review,
        missing_clauses=missing_clauses,
        numeric_mismatches=numeric_mismatches,
        credit_mismatches=credit_mismatches,
        semantic_changes=semantic_changes,
        dedup_changes=dedup_changes,
        added_clauses=added_clauses,
        added_sections=added_sections,
    )


def git_head_text(rel_path: str) -> str | None:
    try:
        proc = subprocess.run(
            ["git", "show", f"HEAD:{rel_path}"],
            cwd=REPO_ROOT,
            capture_output=True,
            timeout=10,
            check=False,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if proc.returncode != 0:
        return None
    text, _ = decode_text_bytes(proc.stdout)
    return text


def collect_git_modified_txts() -> list[Path]:
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
    entries = proc.stdout.decode("utf-8", errors="replace").split("\0")
    out: list[Path] = []
    i = 0
    while i < len(entries):
        e = entries[i]
        if not e or len(e) < 4:
            i += 1
            continue
        status = e[:2]
        rel = e[3:]
        if "R" in status or "C" in status:
            i += 1
        p = REPO_ROOT / rel
        if p.is_file() and p.suffix.lower() == ".txt":
            out.append(p)
        i += 1
    return sorted(set(out))


def emit_review_json(rel_path: str, title: str, before_text: str, after_text: str, note: str) -> Path:
    REVIEW_DIR.mkdir(parents=True, exist_ok=True)
    norm_rel = rel_path.replace("\\", "/")
    review_id = hashlib.sha1(norm_rel.encode("utf-8")).hexdigest()[:16]
    now = datetime.now(timezone.utc).isoformat()
    payload = {
        "id": review_id,
        "resource_path": norm_rel,
        "title": title,
        "original_content": before_text,
        "normalized_content": after_text,
        "note": note,
        "created_at": now,
        "updated_at": now,
        "required_signatures": 1,
        "signatures": [],
        "status": "pending",
    }
    target = REVIEW_DIR / f"{review_id}.json"
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description="AI 规范化改稿无损对账与语义升格守门员 v3（AI Harness 组件）")
    parser.add_argument("targets", nargs="*", help="待与 Git HEAD 对账的资源文件路径")
    parser.add_argument("--git-changed", action="store_true", help="对当前工作区已修改的 序列库/ TXT 逐一对账 Git HEAD")
    parser.add_argument("--before", help="显式指定修改前文件路径（配合 --after 单文件沙盒对账）")
    parser.add_argument("--after", help="显式指定 AI 修改后文件路径（配合 --before 单文件沙盒对账）")
    parser.add_argument("--emit-review", action="store_true", help="强制为对账目标生成 web/normalization_reviews/*.json 人工签审单")
    parser.add_argument("--auto-review-on-semantic", action="store_true", help="当检测到语义变更或重影去重时自动生成人工签审单")
    parser.add_argument("--note", default="", help="附加在人工签审单上的说明")
    parser.add_argument("--json", action="store_true", help="输出 JSON 格式对账结果")
    args = parser.parse_args()

    reports: list[LosslessReport] = []
    emitted_reviews: list[str] = []

    def _maybe_emit(rep: LosslessReport, rel_label: str, stem: str, b_txt: str, a_txt: str) -> None:
        if args.emit_review or (args.auto_review_on_semantic and rep.needs_semantic_review and rep.passed):
            auto_note = args.note or "; ".join((rep.semantic_changes + rep.dedup_changes + rep.added_clauses)[:6])
            r_file = emit_review_json(rel_label, stem, b_txt, a_txt, auto_note)
            emitted_reviews.append(str(r_file.relative_to(REPO_ROOT)))

    if args.before and args.after:
        b_path = Path(args.before)
        a_path = Path(args.after)
        b_text, _ = decode_text_bytes(b_path.read_bytes())
        a_text, _ = decode_text_bytes(a_path.read_bytes())
        label = str(b_path)
        try:
            label = str(b_path.resolve().relative_to(REPO_ROOT))
        except ValueError:
            pass
        rep = compare_lossless(label, b_text, a_text)
        reports.append(rep)
        _maybe_emit(rep, label, b_path.stem, b_text, a_text)
    else:
        files = collect_git_modified_txts() if args.git_changed or not args.targets else [Path(t).resolve() for t in args.targets]
        for f in files:
            if not f.is_file() or f.suffix.lower() != ".txt":
                continue
            rel = str(f.relative_to(REPO_ROOT)).replace("\\", "/")
            before_text = git_head_text(rel)
            if before_text is None:
                continue
            after_text, _ = decode_text_bytes(f.read_bytes())
            rep = compare_lossless(rel, before_text, after_text)
            reports.append(rep)
            _maybe_emit(rep, rel, f.stem, before_text, after_text)

    failed = [r for r in reports if not r.passed]
    sem_flagged = [r for r in reports if r.passed and r.needs_semantic_review]

    if args.json:
        print(
            json.dumps(
                {
                    "checked_files": len(reports),
                    "clean_pass_count": len(reports) - len(failed) - len(sem_flagged),
                    "semantic_review_count": len(sem_flagged),
                    "failed_count": len(failed),
                    "emitted_reviews": emitted_reviews,
                    "reports": [asdict(r) for r in reports],
                },
                ensure_ascii=False,
                indent=2,
            )
        )
    else:
        print(
            f"对账完成：核验 {len(reports)} 份 | 绿牌纯格式通过: {len(reports) - len(failed) - len(sem_flagged)} "
            f"| 黄牌语义/去重复核: {len(sem_flagged)} | 红牌异常拦截: {len(failed)}"
        )
        for r in reports:
            if not r.passed:
                tag = "FAIL (红牌拦截)"
            elif r.needs_semantic_review:
                tag = "REVIEW (黄牌需复核)"
            else:
                tag = "PASS (绿牌纯格式)"
            print(f"[{tag}] {r.path}")
            for msg in r.missing_clauses:
                print(f"  ❌ {msg}")
            for msg in r.numeric_mismatches:
                print(f"  ❌ {msg}")
            for msg in r.credit_mismatches:
                print(f"  ❌ {msg}")
            for msg in r.semantic_changes[:6]:
                print(f"  ⚠️ [语义字段变更] {msg}")
            if len(r.semantic_changes) > 6:
                print(f"  ⚠️ ... 其余 {len(r.semantic_changes) - 6} 处语义字段变更")
            for msg in r.dedup_changes[:6]:
                print(f"  ⚠️ [重影清理去重] {msg}")
            if len(r.dedup_changes) > 6:
                print(f"  ⚠️ ... 其余 {len(r.dedup_changes) - 6} 处重影清理")
            for msg in r.added_clauses:
                print(f"  ⚠️ [新增说明词句] {msg}")
            for msg in r.added_sections:
                print(f"  ℹ️ {msg}")
        for rv in emitted_reviews:
            print(f"已生成人工签审单：{rv}")

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())
