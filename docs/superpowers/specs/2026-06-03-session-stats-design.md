# 结团统计系统设计

## 目标

在现有 Web MVP 中新增一个独立的结团统计模块。维护者选择统计月份并批量上传 `.txt` 结团文件后，后端调用 OpenAI 大模型抽取团名、时长、KP 和 PL 列表，成功项自动写入 SQLite，失败项进入异常列表。前端提供独立统计页，展示每月团数、玩家游戏时长、轮回次数、主持次数和主持时长。

## 统计口径

- `1 个结团文件 = 1 团 = 1 次轮回`。
- 导入时手动选择月份，例如 `2026-05`；不识别精确日期。
- 每个 PL 计入：游戏次数 `+1`、游戏时长 `+团时长`、轮回次数 `+1`。
- KP 也是玩家，自动计入：游戏次数 `+1`、游戏时长 `+团时长`、轮回次数 `+1`。
- KP 额外计入主持维度：主持次数 `+1`、主持时长 `+团时长`。
- KP 不需要出现在 PL 字段中；统计聚合时 `kp` 和 `pl` 参与记录都算玩家维度，只有 `kp` 参与记录额外算主持维度。

## 架构

后端新增 `web/backend/app/session_stats.py`，集中负责 SQLite 初始化、OpenAI 抽取、导入入库和统计查询。数据库文件放在 `web/data/session_stats.sqlite`，不改动 `序列库/`、`荣誉室/`，也不影响 CHM/ZIP 构建和 Wiki 同步。

FastAPI 在 `web/backend/app/main.py` 中挂载结团统计路由。导入、删除等写操作走现有 `ADMIN_PASSWORD` 鉴权；概览、玩家排行和团列表公开可读。异常列表包含模型输出和内部错误信息，也走后台鉴权。

前端新增独立页面入口，和“阅读 / 更新 / 后台”并列。页面包含月份选择、批量上传、导入结果、月度概览、玩家排行、团列表和异常列表。

## 数据库

SQLite 使用以下表：

- `import_batches`：一次批量导入记录，包含月份、文件数、成功数、失败数、模型名和创建时间。
- `players`：玩家档案，包含显示名、QQ、酒馆绑定字段和创建时间。KP 有 QQ 时优先按 QQ 匹配；否则按昵称匹配。PL 默认按昵称匹配。
- `sessions`：成功入库的团记录，包含月份、团名、来源文件名、团时长、KP 玩家 ID、AI 原始 JSON、原文 hash 和创建时间。`month + source_filename + content_hash` 唯一，避免重复导入。
- `session_participants`：团参与记录，包含 `session_id`、`player_id`、`role`。`role` 为 `kp` 或 `pl`。
- `import_errors`：导入失败记录，包含批次、月份、文件名、失败原因、AI 原始输出和原文 hash。

统计结果不写死到总表，前端查询时通过 SQL 聚合得到。

## OpenAI 抽取

后端通过环境变量配置：

- `OPENAI_API_KEY`
- `OPENAI_MODEL`，默认 `gpt-5.4-mini`
- 可选 `OPENAI_BASE_URL`

抽取使用 OpenAI Responses API 的结构化 JSON 输出。输入包含文件名、选择的月份和全文内容。模型必须返回：

- `title`：团名
- `duration_hours`：团时长，单位小时，支持小数
- `kp.name`
- `kp.qq`
- `players[].name`
- `confidence`
- `warnings[]`

如果缺少 KP、时长、PL，或 JSON 不符合 schema，则该文件不计入统计，写入异常列表。

## API

- `POST /api/session-stats/import`
  上传 `month` 和多个 `.txt` 文件，返回批次 ID、成功数量、失败数量和逐文件结果。
- `GET /api/session-stats/overview?month=2026-05`
  返回月度团数、玩家人次、游戏小时、主持小时、导入成功/失败数。
- `GET /api/session-stats/players?month=2026-05&sort=hours`
  返回玩家排行。
- `GET /api/session-stats/sessions?month=2026-05`
  返回已入库团列表。
- `GET /api/session-stats/errors?month=2026-05`
  返回异常文件列表，需要后台鉴权。
- `DELETE /api/session-stats/sessions/{id}`
  删除误入库的团及参与记录，便于修正后重导。

## 前端

新增统计页：

- 顶部月份选择和批量 TXT 上传控件。
- 导入面板显示成功数、失败数和每个文件状态。
- 概览区显示本月团数、玩家人次、游戏小时、主持小时。
- 玩家表展示玩家名、QQ、游戏次数、游戏时长、轮回次数、主持次数、主持时长。
- 团列表展示团名、来源文件、时长、KP、PL 数量和删除按钮。
- 异常列表展示文件名和失败原因。

## 风险控制

- 不做人工复核主流程，但保留失败列表和删除重导能力。
- 数据库写入使用事务，单文件失败不影响其他文件。
- 同一文件内容重复导入时返回重复提示，不重复计入统计。
- OpenAI 未配置时导入接口返回明确错误，统计查询仍可用。
- 读取 TXT 时兼容 `utf-8-sig`、`utf-8`、`gbk`、`gb2312`、`big5`。

## 验证

- 后端单元测试覆盖数据库初始化、成功导入、重复导入、缺字段失败、统计聚合和删除团。
- 后端 smoke check 继续覆盖现有资源接口。
- 前端执行 TypeScript 构建，确认新增页面类型正确。
