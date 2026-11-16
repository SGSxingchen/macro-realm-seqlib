import importlib.util
from pathlib import Path
import sys
import tempfile
import unittest


REPO_ROOT = Path(__file__).resolve().parents[1]
WORKFLOW_SCRIPTS = REPO_ROOT / ".agents" / "skills" / "seqlib-resource-workflow" / "scripts"


def _load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


lint_mod = _load_module("lint_resources", WORKFLOW_SCRIPTS / "lint_resources.py")
lossless_mod = _load_module("verify_lossless", WORKFLOW_SCRIPTS / "verify_lossless.py")


class ResourceHarnessTest(unittest.TestCase):
    def _lint_temp_txt(self, rel_name: str, content: str, strict: bool = True):
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp) / rel_name
            p.parent.mkdir(parents=True, exist_ok=True)
            p.write_text(content, encoding="utf-8")
            issues, cards = lint_mod.check_txt_file(p, rel_name, strict=strict)
            return issues, cards

    def test_linter_catches_v3_blindspots(self):
        sample = """081】测试菌主
［制作人:闲狐狸］
【日期待复核：旧版微调】
2025年11月14日更新：S级技能新增联动折扣
[异化特质]
这里是一段紧跟在更新日志下方、没有小节标题包裹的背景介绍，会导致更新日志变成大标题。
E级：解锁基础性能
B级：解锁两项高级能力
————————————
[技艺条件]：冥想难度150
[能力名称]：真菌
[初级能力效果]：视野增加10%
[能力效果］：受到火属性100%易伤
[释放类型]：被动

【子实体】：buff，等级等同于特质等级

【神通】：炼就神通需要根器支撑

制作人：姜雷
【判定加值增加，冷却降低】
【该义体由基础义眼升级而来】
能力名称]：千里目（A级）
[能力效果]：侦察判定+20%
"""
        issues, cards = self._lint_temp_txt("序列库/特质改造/异化改造类/081】测试菌主.txt", sample, strict=True)
        codes = {i.code for i in issues}
        self.assertEqual(cards, 2)
        self.assertIn("BRACKETED_CREDIT_LINE", codes)
        self.assertIn("LEGACY_PENDING_DATE_LOG", codes)
        self.assertIn("UNBRACKETED_CHANGELOG", codes)
        self.assertIn("HEADER_LOG_SWALLOWS_INTRO", codes)
        self.assertIn("BARE_TRAIT_GROWTH_LINE", codes)
        self.assertIn("DECORATIVE_DIVIDER_LINE", codes)
        self.assertIn("PRE_CARD_FLOATING_FIELD", codes)
        self.assertIn("NON_STANDARD_FIELD_ALIAS", codes)
        self.assertIn("MIXED_BRACKET_FIELD", codes)
        self.assertIn("HEAVY_BRACKET_COLON_ORPHAN", codes)
        self.assertIn("SECTION_WITH_TRAILING_COLON", codes)
        self.assertIn("INLINE_CARD_CREATOR", codes)
        self.assertIn("UNPREFIXED_MOD_NOTE", codes)
        self.assertIn("CARD_TOP_UNLOCK_OR_SCOPE_RULE", codes)
        self.assertIn("BROKEN_CARD_LEFT_BRACKET", codes)
        self.assertNotIn("FULLWIDTH_BRACKET_FIELD", codes)

    def test_linter_catches_lenticular_and_zero_card_legacy(self):
        legacy_zero_card = """009】测试血莲

（制作人：红月）

〖能力名称〗血莲心经
〖能力效果〗每秒恢复10点生命值。
"""
        issues, cards = self._lint_temp_txt("序列库/特质改造/异化改造类/009】测试血莲.txt", legacy_zero_card, strict=True)
        codes = {i.code for i in issues}
        self.assertEqual(cards, 0)
        self.assertIn("LENTICULAR_BRACKET_FIELD", codes)
        self.assertIn("ZERO_CARDS_IN_RESOURCE", codes)

    def test_linter_catches_blank_line_and_leading_newline_hygiene(self):
        messy_spacing = """081】测试菌主


（制作人：闲狐狸）

【基础性能】
[能力名称]：真菌
[能力效果]：
效果1：受到火属性100%易伤
[补充说明]：
「子实体」：buff，等级等同于特质等级

「感染」：物理debuff，等级等同于特质等级
"""
        issues, _ = self._lint_temp_txt("序列库/特质改造/异化改造类/081】测试菌主.txt", messy_spacing, strict=True)
        codes = {i.code for i in issues}
        self.assertIn("CONSECUTIVE_BLANK_LINES", codes)
        self.assertIn("SECTION_MISSING_BLANK_AFTER", codes)
        self.assertIn("MISSING_BLANK_BEFORE_CARD", codes)
        self.assertIn("FIELD_LEADING_NEWLINE", codes)
        self.assertIn("BLANK_LINE_INSIDE_CARD", codes)

    def test_linter_passes_cleanly_after_v3_ai_normalization(self):
        normalized = """081】测试菌主

（制作人：闲狐狸）

【历史更新：早期版本整体微调】
【25.11.14：S级技能新增联动折扣】

【特质说明】

[异化特质]
【E】级：解锁基础性能
【B】级：解锁两项高级能力

【基础性能】

[能力名称]：真菌
[技艺规则]：冥想难度150
[初级效果]：视野增加10%
[能力效果]：受到火属性100%易伤
[释放类型]：被动
[补充说明]：「子实体」：buff，等级等同于特质等级
「感染」：物理debuff，等级等同于特质等级

【特定高级能力】

【说明：制作人：姜雷】
【能力修改：判定加值增加，冷却降低】
【说明：该义体由基础义眼升级而来】
[能力名称]：千里目（A级）
[能力效果]：侦察判定+20%
"""
        issues, cards = self._lint_temp_txt("序列库/特质改造/异化改造类/081】测试菌主.txt", normalized, strict=True)
        self.assertEqual(cards, 2)
        self.assertEqual(issues, [])

        # 存量老稿缺失制作人/审核人时，默认 --strict 不弹警告
        no_author_legacy = normalized.replace("（制作人：闲狐狸）\n\n", "")
        issues_legacy, _ = self._lint_temp_txt("序列库/特质改造/异化改造类/081】测试菌主.txt", no_author_legacy, strict=True)
        self.assertEqual(issues_legacy, [])

    def test_verify_lossless_green_pass_for_supplement_and_inline_creator_and_yellow_for_gate_or_dedup(self):
        before = """015】义体测试
［制作人:红月］
【26.5.8：千里目加强】
调整人：红月
审核人：沧羽、玛利亚
————————————
[技艺条件]：冥想难度150
制作人：姜雷
【判定加值增加，高亮效果冷却降低】
【该义体由基础歧路司义眼升级而来】
[义体名称]：歧路司义眼“千里目”（A级）
[初级能力效果]：视野范围+15%/30%/45%，回复1D5生命。
[能力效果］：侦察判定增加最高属性*0.8点。
[释放类型]：被动

【子实体】：buff，等级等同于你的特质等级。
"""
        after_green = """015】义体测试

（制作人：红月）
（调整人：红月）
（审核人：沧羽、玛利亚）

【26.5.8：千里目加强】

【说明：制作人：姜雷】
【能力修改：判定加值增加，高亮效果冷却降低】
【说明：该义体由基础歧路司义眼升级而来】
[义体名称]：歧路司义眼“千里目”（A级）
[技艺规则]：冥想难度150
[初级效果]：视野范围+15%/30%/45%，回复1D5生命。
[能力效果]：侦察判定增加最高属性*0.8点。
[释放类型]：被动
[补充说明]：
「子实体」：buff，等级等同于你的特质等级。
"""
        rep_green = lossless_mod.compare_lossless("序列库/技能表/科技侧/015】义体测试.txt", before, after_green)
        self.assertTrue(rep_green.passed, msg=f"Unexpected red card: {rep_green}")
        self.assertFalse(rep_green.needs_semantic_review, msg=f"Unexpected yellow card: {rep_green}")

        # 升格为 [使用要求]：触发黄牌语义复核
        after_yellow_gate = after_green.replace(
            "【说明：该义体由基础歧路司义眼升级而来】\n[义体名称]：歧路司义眼“千里目”（A级）",
            "[义体名称]：歧路司义眼“千里目”（A级）\n[使用要求]：该义体由基础歧路司义眼升级而来",
        )
        rep_gate = lossless_mod.compare_lossless("序列库/技能表/科技侧/015】义体测试.txt", before, after_yellow_gate)
        self.assertTrue(rep_gate.passed)
        self.assertTrue(rep_gate.needs_semantic_review)
        self.assertTrue(any("使用要求" in m for m in rep_gate.semantic_changes))

        # 清理卡顶 100% 重复重影（如《018】役亡师》）：passed=True 且挂黄牌 dedup_changes
        before_dup = """018】役亡师测试

（制作人：无语）

【死非终结】：恢复目标20%生命值，冷却3回合。
[技能名称]：死非终结
[能力效果]：恢复目标20%生命值，冷却3回合。
"""
        after_dup_cleaned = """018】役亡师测试

（制作人：无语）

[能力名称]：死非终结
[能力效果]：恢复目标20%生命值，冷却3回合。
"""
        rep_dedup = lossless_mod.compare_lossless("序列库/职业/神秘侧/018】役亡师测试.txt", before_dup, after_dup_cleaned)
        self.assertTrue(rep_dedup.passed, msg=f"Dedup should pass red gate: {rep_dedup}")
        self.assertTrue(rep_dedup.needs_semantic_review)
        self.assertTrue(len(rep_dedup.dedup_changes) >= 1)

    def test_verify_lossless_catches_ai_number_hallucination_and_dropped_clause(self):
        before = """001】测试表

（制作人：红月）

[能力名称]：精准射击（C级）
[能力效果]：暴击伤害提升15%/30%/45%，消耗10点魔力。
若目标处于眩晕状态，额外造成2D6真实伤害。
[冷却时间]：3回合
"""
        after_bad = """001】测试表

（制作人：红月）
（审核人：路人甲）

[能力名称]：精准射击（C级）
[能力效果]：暴击伤害提升15%/30%/40%，消耗10点魔力。
[冷却时间]：3回合
"""
        rep = lossless_mod.compare_lossless("序列库/技能表/001】测试表.txt", before, after_bad)
        self.assertFalse(rep.passed)
        self.assertTrue(any("45%" in m for m in rep.numeric_mismatches))
        self.assertTrue(any("2d6" in m for m in rep.numeric_mismatches))
        self.assertTrue(any("若目标处于眩晕状态" in m for m in rep.missing_clauses))
        self.assertTrue(any("路人甲" in m for m in rep.credit_mismatches))


if __name__ == "__main__":
    unittest.main()
