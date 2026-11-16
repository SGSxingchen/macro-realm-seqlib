# 宏观界域强化序列库 (Macro-Realm Sequence Library)

> 本文件与根目录 `AGENTS.md` 保持核心规范对齐。完整工作区规则见 **`AGENTS.md`**，资源文本与卡片唯一标准见 **`资源标准模板.txt`**，在线站前端规范见 **`docs/frontend-style-and-interaction.md`**。

## 1. 仓库专属 Skills（位于 `.agents/skills/`）
- **`seqlib-resource-workflow`**（`.agents/skills/seqlib-resource-workflow/SKILL.md`）：
  负责序列资源审稿、正式入库（新增 / 增量调整 / 全量重置 / 回归新增 / 下架移入荣誉室）、基于 AI 语义理解的无损规范化与 `[XX名称]` 泛匹配验卡。修改或规范化资源后请运行双重只读 Harness：
  ```bash
  python3 .agents/skills/seqlib-resource-workflow/scripts/lint_resources.py --git-changed --strict
  python3 .agents/skills/seqlib-resource-workflow/scripts/verify_lossless.py --git-changed
  ```
- **`seqlib-release`**（`.agents/skills/seqlib-release/SKILL.md`）：
  负责版本发布前体检、`tools/renumber_resources.py` 编号重排、`build_config.json` 与编者注/更新日志版本对齐、`build_chm.py` 打包、Wiki 同步与 Web 静态页构建验证。

## 2. 核心红线速览
1. **入库与荣誉室边界**：默认更新正式 `序列库/`（仅用户明确指定体验服/新过审目录时才写 `序列库/新过审序列/`）；`荣誉室/` 默认只读保留，重置回归在 `序列库/` 新增编号而不删改 `荣誉室/` 旧档。
2. **规则无损与高兼容**：整理格式严禁改动任何数值、冷却、消耗、持续、距离、判定或规则边界；存疑处保留原文并列入人工复核。
3. **模板与 `[XX名称]` 泛匹配成卡**：
   - 文件名禁 `/ \ : * ? " < > |`（正文内部正常允许 `/`、`*`、引号与半角 `:`），TXT 统一 UTF-8，首行标题严格等于文件名去 `.txt` 且标题后空一行；
   - 普通能力用 `[能力名称]：`，专用实体（称号、道具、能量池、模块、义体、符文、单位、奖励、职阶等）直接保留 `[XX名称]：` 泛匹配成卡；
   - 单张卡片正上方的最近修改备注统一加前缀写 `【能力修改：...】`；技艺规则固定写 `[技艺规则]：`；特质主成长树统一为 `【E】级：...`～`【S/EX】级：...`。
4. **Web 在线站（`web/frontend/`）**：坚持「无限流高维兑换终端」双主题（虚空棱镜暗色 / 白昼终端亮色、`2px` 微切角、`1px` 细线分割）与 `Reader/adaptive.ts` 100% 原文区间无损契约，详见 `docs/frontend-style-and-interaction.md`。
