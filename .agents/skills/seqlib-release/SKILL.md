---
name: seqlib-release
description: >-
  宏观界域强化序列库的版本发版、资源编号统一重排、根目录编者注与更新日志维护、`build_config.json` 配置对齐、CHM/ZIP 打包校验、Fandom Wiki 同步及 Web 在线站静态页更新流程。当用户要求准备新版本发布（如 V6.8）、执行全库重排编号、检查发版配置或同步多端产物时使用。
---

# 序列库版本发版与多端同步工作流

本技能用于统筹序列库大/小版本发布前的完整性门禁，防止漏配 `build_config.json`、编者注残留旧版本号、脏文件打入发布包或重排编号后标题失步。

## 发布前五步标准检查单

### 第 1 步：全库资源完整性与脏文件体检
运行 `seqlib-resource-workflow` 提供的只读校验器，确保正式库内无系统脏文件（如 `desktop.ini`）、无非 UTF-8 文本、无吞卡缺括号错误：
```bash
python3 .agents/skills/seqlib-resource-workflow/scripts/lint_resources.py --all
```
- 若发现 `DIRTY_SYSTEM_FILE`、`BROKEN_CARD_LEFT_BRACKET`、`NON_UTF8_ENCODING`，必须在重排与打包前修复或清理。

### 第 2 步：统一重排编号并同步首行标题
当本版本包含新增资源、下架移入 `荣誉室/` 或重置回归时，使用 `tools/renumber_resources.py`：
```bash
# 1. 先试运行审阅重排与首行同步计划（不落盘）
python3 tools/renumber_resources.py

# 2. 确认计划符合预期后正式落盘
python3 tools/renumber_resources.py --apply
```
- 落盘后再次运行 `lint_resources.py --all` 确认 `TITLE_MISMATCH` 归零。

### 第 3 步：根目录发布元数据与 `build_config.json` 对齐（高频漏点！）
准备新主/次版本（如从 `6.7` 升至 `6.8`）时，必须同步核对以下 4 处：
1. **`{版本号}序列库编者注.txt`**（如 `6.8序列库编者注.txt`）：
   - 核对首行标题、制作人与联系方式；
   - **重点核对文末「额外内容」中的本次版本号、对照基线版本号及《V{版本号}序列库更新日志》名称**，切勿复制上版文件后遗留旧版本号字样！
2. **`V{版本号}序列库更新日志.txt`** 与 **`第N批下架名单.txt`**：确认新增、调整、重置、下架条目与实际 Git diff 一致。
3. **`build_config.json`**：
   - 在 `"root_files"` 中新增对应主次版本号键（如 `"6.8"`），显式列出本版编者注、更新日志、本批下架名单与 `"资源标准模板.txt"`（否则 `build_chm.py` 会回退并把历代所有旧下架名单全部塞进 ZIP 包！）。
4. **`README.md` / `AGENTS.md` / `CLAUDE.md`**：若目录结构或制作人信息变更，同步更新说明。

### 第 4 步：本地构建验证（ZIP + Web 静态页）
1. **打包试跑**（macOS/Linux 可先验 `--skip-chm` 验证 ZIP 内容结构，Windows 或 GitHub Actions `windows-latest` 构建完整 CHM + ZIP）：
   ```bash
   python3 build_chm.py --version v6.8 --skip-chm
   ```
   检查生成的 ZIP 包内是否不含 `.obsidian/`、`desktop.ini` 或上个版本的旧编者注。
2. **Web 在线站静态条目页与无损契约回归**：
   ```bash
   cd web/frontend
   npm run build
   node --test tests/*.test.cjs
   ```
   确保 `dist/entry/` 静态页数量与正式库资源一致，且原文无损测试全绿。

### 第 5 步：打标发布与 Wiki / 在线站同步
1. 提交并推送 `main` 与 `v*` 标签触发 `.github/workflows/release.yml` 自动构建 CHM + ZIP Release：
   ```bash
   git tag -a v6.8 -m "发布 V6.8 序列库更新"
   git push origin main
   git push origin v6.8
   ```
2. **注意在线站部署触发条件**：`.github/workflows/deploy-web.yml` 默认监听 `web/**` 变更；若本次提交仅修改了 `序列库/` 文本而未触碰 `web/**`，推送后请手动触发一次 `deploy-web.yml`（`workflow_dispatch`）以刷新线上 `trpg.chordvers.org` 的搜索索引与静态 SEO 页面。
3. **Fandom Wiki 同步（按需）**：先跑 `--dry-run` 预览，确认分类无误后再正式同步：
   ```bash
   python3 wiki/sync_to_wiki.py --user BotName --password xxx --dry-run
   ```
