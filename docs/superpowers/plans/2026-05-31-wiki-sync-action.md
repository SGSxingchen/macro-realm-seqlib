# Wiki Sync Action Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增一个独立的 GitHub Actions 手动工作流，用于 dry-run 或真实同步 Fandom Wiki，并默认通过 diff 增量同步处理删除和重命名。

**Architecture:** 工作流独立放在 `.github/workflows/wiki-sync.yml`，只通过 `workflow_dispatch` 触发。参数组装集中在 PowerShell 步骤里完成，默认用最近 tag 作为 `--diff-from` 基准；真实同步通过 GitHub Secrets 读取凭据，dry-run 使用占位凭据。

**Tech Stack:** GitHub Actions、PowerShell、Python、pandoc、mwclient。

---

### Task 1: 工作流结构检查

**Files:**
- Create: `tests/test_wiki_sync_workflow.py`

- [ ] **Step 1: 写入失败检查**

创建 `tests/test_wiki_sync_workflow.py`，检查 `.github/workflows/wiki-sync.yml` 存在，并包含手动触发、输入项、Secrets 校验和同步脚本调用。

- [ ] **Step 2: 运行检查确认失败**

运行：`python -m unittest tests.test_wiki_sync_workflow -v`

预期：失败，原因是 `.github/workflows/wiki-sync.yml` 尚不存在。

### Task 2: 新增 Wiki 同步工作流

**Files:**
- Create: `.github/workflows/wiki-sync.yml`

- [ ] **Step 1: 新增工作流**

创建 `.github/workflows/wiki-sync.yml`，包含：

- 手动触发 `workflow_dispatch`
- `mode` 下拉选项：`dry-run`、`sync`
- `skip_honor` 下拉选项：`false`、`true`
- `sync_range` 下拉选项：`latest-tag`、`custom-ref`、`last-commit`、`full`
- `diff_from` 可选字符串，供 `custom-ref` 使用
- `filter` 可选字符串
- `delay` 可选字符串，默认 `5`
- checkout 使用 `fetch-depth: 0`
- 安装 Python、pandoc、`mwclient`
- dry-run 使用占位账号密码
- sync 读取 `WIKI_USER`、`WIKI_PASSWORD`
- sync 缺少 Secrets 时失败
- `latest-tag` 通过 `git describe --tags --abbrev=0` 生成 `--diff-from`
- `last-commit` 传入 `--incremental`
- 真实同步时阻止 `filter + full`，避免误删

- [ ] **Step 2: 运行检查确认通过**

运行：`python -m unittest tests.test_wiki_sync_workflow -v`

预期：通过。

### Task 3: 最终验证

**Files:**
- Verify: `.github/workflows/wiki-sync.yml`
- Verify: `tests/test_wiki_sync_workflow.py`

- [ ] **Step 1: 检查工作区差异**

运行：`git diff -- .github/workflows/wiki-sync.yml tests/test_wiki_sync_workflow.py docs/superpowers/specs/2026-05-31-wiki-sync-action-design.md docs/superpowers/plans/2026-05-31-wiki-sync-action.md`

- [ ] **Step 2: 检查状态**

运行：`git status --short`
