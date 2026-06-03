# Session Stats Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在现有 Web MVP 中新增结团统计模块，支持按月份批量上传结团 TXT，调用 OpenAI 抽取后自动写入 SQLite，并在前端独立页面展示玩家时长、轮回次数和主持维度。

**Architecture:** 后端新增聚焦的 `session_stats` 模块，负责 SQLite、OpenAI 抽取、导入和聚合查询；`main.py` 只挂载路由。前端新增统计页面组件和类型，复用现有 `api` 请求工具。

**Tech Stack:** Python、FastAPI、SQLite、OpenAI Responses API、React、TypeScript、Vite。

---

### Task 1: 后端数据层和导入规则

**Files:**
- Create: `web/backend/app/session_stats.py`
- Create: `tests/test_session_stats.py`
- Modify: `web/backend/requirements.txt`

- [ ] 写失败测试：用临时 SQLite 和假 AI 抽取器导入一个文件，断言 KP 与 PL 都计入玩家维度，KP 额外计入主持维度。
- [ ] 运行 `python -m unittest tests.test_session_stats -v`，确认因为模块缺失失败。
- [ ] 实现 SQLite schema、玩家 upsert、团入库、参与记录、统计聚合。
- [ ] 增加重复导入和缺字段失败测试。
- [ ] 运行 `python -m unittest tests.test_session_stats -v`，确认通过。

### Task 2: OpenAI 抽取封装

**Files:**
- Modify: `web/backend/app/session_stats.py`
- Modify: `tests/test_session_stats.py`

- [ ] 写失败测试：OpenAI 未配置时导入返回失败记录，不写入团统计。
- [ ] 实现 `OpenAIExtractor`，读取 `OPENAI_API_KEY`、`OPENAI_MODEL`、可选 `OPENAI_BASE_URL`，通过 Responses API 请求结构化 JSON。
- [ ] 保持测试使用假抽取器，不真实调用网络。
- [ ] 运行 `python -m unittest tests.test_session_stats -v`，确认通过。

### Task 3: FastAPI 路由

**Files:**
- Modify: `web/backend/app/main.py`
- Modify: `web/backend/app/session_stats.py`
- Create: `tests/test_session_stats_api.py`

- [ ] 写失败测试：`GET /api/session-stats/players` 可公开读取空统计。
- [ ] 写失败测试：未配置 `ADMIN_PASSWORD` 时 `POST /api/session-stats/import` 返回写操作不可用。
- [ ] 实现路由挂载、月份校验、批量上传、概览、玩家排行、团列表、异常列表和删除团。
- [ ] 运行 `python -m unittest tests.test_session_stats_api -v`，确认通过。

### Task 4: 前端统计页

**Files:**
- Modify: `web/frontend/src/types.ts`
- Modify: `web/frontend/src/App.tsx`
- Modify: `web/frontend/src/components/Header.tsx`
- Create: `web/frontend/src/components/SessionStats.tsx`
- Modify: `web/frontend/src/style.css`

- [ ] 新增前端类型和统计页组件。
- [ ] 在 Header 增加“结团统计”入口。
- [ ] 统计页实现月份输入、批量上传、概览、玩家排行、团列表和异常列表。
- [ ] 删除团按钮调用后端接口后刷新统计。
- [ ] 运行 `cd web/frontend && npm run build`，确认 TypeScript 和生产构建通过。

### Task 5: 集成验证

**Files:**
- Verify: `web/backend/app/session_stats.py`
- Verify: `web/backend/app/main.py`
- Verify: `web/frontend/src/components/SessionStats.tsx`
- Verify: `docs/superpowers/specs/2026-06-03-session-stats-design.md`
- Verify: `docs/superpowers/plans/2026-06-03-session-stats.md`

- [ ] 运行 `python -m unittest tests.test_session_stats tests.test_session_stats_api -v`。
- [ ] 运行 `python web/backend/smoke_check.py`。
- [ ] 运行 `cd web/frontend && npm run build`。
- [ ] 运行 `git diff --stat` 和 `git status --short`，确认改动范围符合计划。
