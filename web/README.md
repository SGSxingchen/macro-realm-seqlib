# 序列库轻量 Web MVP

本目录提供一个不改变既有文件树事实源的轻量前后端：

- 后端：FastAPI，直接扫描/读写仓库根目录下的 `序列库/`、`荣誉室/` TXT 文件。
- 前端：Vite + React + TypeScript，提供分类浏览、智能搜索、详情阅读和维护后台。
- 不引入数据库、资源 ID、资源表或改名追踪表；路径和文件名仍是资源索引与导出依据。
- CHM/ZIP 继续使用根目录已有 `build_chm.py` 和 tag release workflow。

## 智能搜索

后端会在启动时构建进程内倒排索引，按 mtime 增量缓存，写操作触发失效。查询支持：

- **多 token AND**：用空格分隔 `百夫长 罗马`，全部命中才返回
- **混合容错**：允许 1 个 token 失配，前提是其他 token 强命中（标题包含/拼音命中），便于英中混搜（如 `Centurion 百夫长`）
- **拼音整串 / 首字母**：`baifuzhang` / `bfz` 都能命中"百夫长"
- **繁简归一**：输入繁体也能匹到简体资源（依赖 `opencc-python-reimplemented`）
- **子序列模糊**：`强驱散` 命中"强制驱散"（紧凑度门槛防散落假命中）
- **分面 facets**：返回侧、资源类型、作者三组聚合计数，前端做多选筛选

`pypinyin` 和 `opencc-python-reimplemented` 是**可选依赖**：缺失时自动降级为不带拼音/繁简归一的纯文本匹配，不会启动失败。

## 启动后端

```bash
cd web/backend
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# 后台写操作必需；不配置时只允许前台读取，管理写接口会拒绝
export ADMIN_PASSWORD='change-me'
# 可选：独立 cookie 签名密钥；默认从 ADMIN_PASSWORD 派生
export ADMIN_SECRET='long-random-secret'

# Wiki 真实同步才需要；dry-run 可不配置（后端会传入占位参数满足脚本 argparse）
export WIKI_USER='BotName'
export WIKI_PASSWORD='...'

# 结团统计导入必需：战报 TXT 由 LLM 解析，走 OpenAI 兼容接口
export OPENAI_API_KEY='sk-...'                       # 或 THIRD_PARTY_API_KEY
export OPENAI_BASE_URL='https://api.deepseek.com/v1' # 默认 https://api.openai.com/v1
export OPENAI_MODEL='deepseek-v4-pro'                # 默认 deepseek-v4-pro
# 可选：
# export OPENAI_API_MODE='chat'            # chat | responses | auto（默认 auto，deepseek 开头自动走 chat/completions）
# export OPENAI_TIMEOUT_SECONDS='180'      # 单次请求超时，默认 180 秒
# export OPENAI_REQUEST_RETRIES='1'        # 超时重试次数，默认 1

uvicorn app.main:app --reload --host 127.0.0.1 --port 8000
```

> 后端通过 `Path(__file__).parents[3]` 定位仓库根目录，因此请保持 `web/backend/app/main.py` 的相对位置。

## 启动前端

```bash
cd web/frontend
npm install
npm run dev
```

访问 Vite 输出的地址（默认 `http://localhost:5173`）。开发服务器已代理 `/api` 到 `http://127.0.0.1:8000`。

## 主要接口

### 前台

- `GET /api/health`：健康检查与后台是否配置。
- `GET /api/tree`：公开前台分类树，仅包含 `序列库/`。
- `GET /api/resources?q=&root=&category=&kinds=&sides=&authors=&limit=200&offset=0&include_content=false`：资源列表/智能搜索。
- `GET /api/resources/{path}`：资源详情，自动兼容 `utf-8/utf-8-sig/gbk/gb2312/big5`。
- `GET /api/raw/{path}`：纯文本内容。

`kinds`、`sides`、`authors` 是可重复 query 参数，例如：

```text
/api/resources?q=百夫长&kinds=职业&sides=战技侧&authors=沧羽
```

列表字段包含：相对路径、文件名、标题、根目录类型、分类路径、mtime、size、侧向、资源类型、作者、评分和片段。默认不返回全文；需要全文时传 `include_content=true`。

### 结团统计

前台「结团统计」页签：按月统计跑团数据。战报 TXT 由 LLM（OpenAI 兼容接口）解析入库，数据存 `web/data/session_stats.sqlite`。

- `GET /api/session-stats/overview?month=2026-06`：月度概览（团数/玩家人次/总游戏小时/主持小时）。
- `GET /api/session-stats/players?month=&sort=hours|games|hosts|name`：玩家统计（全量返回，前端翻页）。
- `GET /api/session-stats/sessions?month=`：团列表。
- `GET /api/session-stats/sessions/{id}`：团详情（参与者、raw_payload）。
- `PATCH /api/session-stats/sessions/{id}`：编辑标题/时长。
- `DELETE /api/session-stats/sessions/{id}`：删除团记录。
- `POST /api/session-stats/import-jobs`：上传 TXT 创建异步导入任务（multipart：month/concurrency/files）。
- `GET /api/session-stats/import-jobs/{job_id}`：轮询导入进度。
- `POST /api/session-stats/dedupe?month=`：清理内容完全相同的重复团记录。
- `GET /api/session-stats/errors?month=`：导入异常列表（需后台登录）。

导入需要配置 `OPENAI_API_KEY`（或 `THIRD_PARTY_API_KEY`），见上方「启动后端」。同一文件内容（content_hash）重复导入会自动跳过（SKIP）。

### 规范化审核

规范化审核是单独入口，不挂在前台导航，也不使用后台密码。访问 `/normalize-review` 可查看全部审核任务；如果链接带 `?id=...`，页面会默认聚焦该任务，但仍保留全部任务列表。

- `GET /api/normalization/reviews`：列出全部规范化审核任务。
- `POST /api/normalization/reviews`：创建审核任务。body: `{ resource_path, normalized_content, original_content?, title?, note? }`
- `GET /api/normalization/reviews/{id}`：读取某个审核任务的规范化前后文本。
- `POST /api/normalization/reviews/{id}/sign`：签名通过。body: `{ signer, note? }`

审核任务以 JSON 文件保存在 `web/normalization_reviews/`。每个任务记录规范化前文本、规范化后文本、签名列表和状态；当前规则为凑齐 1 个签名即通过。

### 后台鉴权

- `POST /api/admin/login`，body: `{ "password": "..." }`；密码只读取 `ADMIN_PASSWORD`。
- `POST /api/admin/logout`
- `GET /api/admin/me`

未配置 `ADMIN_PASSWORD` 时，所有后台写操作都会通过后端鉴权依赖直接拒绝（HTTP 503）；未登录或 Cookie 无效时返回 HTTP 401。

### 后台文件维护

所有路径必须位于 `序列库/` 或 `荣誉室/` 内且扩展名为 `.txt`，会拒绝绝对路径和 `..` 路径穿越。

- `POST /api/admin/resources`：新增 TXT。`{ path, content, overwrite }`
- `PUT /api/admin/resources/{path}`：粘贴替换保存现有 TXT。
- `POST /api/admin/upload/{path}`：上传 `.txt` 覆盖现有资源。
- `POST /api/admin/move`：重命名/移动。`{ old_path, new_path, overwrite }`
- `POST /api/admin/delete`：删除。`{ path }`
- `POST /api/admin/move-to-honor`：移入荣誉室，默认保留相对分类路径，也可传 `target_path`。

写操作返回操作前后的 `git status --short -- <path>`，便于维护者确认改动。

### Git / diff / 发布

- `GET /api/git/info`：最近 tag、当前 HEAD、分支、工作区状态。
- `GET /api/git/changes?from_ref=v6.4`：从指定 ref（默认最近 tag）到当前 `HEAD + 工作区` 的新增/修改/删除/重命名摘要，不做全文 diff。
- `POST /api/admin/publish`：执行 `git add`、`commit`、`tag`、`push origin main --tags`，body 示例：

```json
{ "version": "v6.5", "message": "发布 v6.5", "branch": "main", "push": true }
```

接口会返回每步完整 stdout/stderr；失败不吞错。请注意：本次实现不会自动调用该接口，也不会自动提交/推送。

发布前 `git add` 使用固定白名单路径，避免误加入 `node_modules/`、`dist/`、`tsconfig.tsbuildinfo` 等安装/构建产物。当前白名单为：`序列库/`、`荣誉室/`、仓库根目录 `*.txt`、`wiki/*.py`、以及 `web/` 下源码/配置/README（不包含前端构建输出和依赖目录）。

### Wiki 同步

- `POST /api/admin/wiki/sync?dry_run=true&skip_honor=false`

后端封装调用：

```bash
python wiki/sync_to_wiki.py [--user $WIKI_USER --password $WIKI_PASSWORD] [--dry-run] [--skip-honor]
```

密钥只从环境变量读取，前端不能传入也不会写入仓库。真实同步（`dry_run=false`）要求配置 `WIKI_USER/WIKI_PASSWORD` 或 `FANDOM_USER/FANDOM_PASSWORD`。


## 验证命令

建议在提交前至少运行：

```bash
# 后端静态/接口基础自检
python web/backend/smoke_check.py

# 不启动服务的鉴权/路径补充检查
python - <<'PY'
import os
from fastapi import HTTPException
from web.backend.app.main import require_admin, safe_resource_path

os.environ.pop('ADMIN_PASSWORD', None)
try:
    require_admin(None)
    raise AssertionError('write auth did not reject missing ADMIN_PASSWORD')
except HTTPException as e:
    assert e.status_code == 503
for bad in ['../README.md', '/tmp/x.txt', 'README.md', '序列库/bad.md', 'web/README.txt']:
    try:
        safe_resource_path(bad)
        raise AssertionError(f'unsafe path accepted: {bad}')
    except HTTPException:
        pass
print('auth/path checks OK')
PY

# 前端类型检查和生产构建
cd web/frontend
npm install
npm run build
```

`.gitignore` 已覆盖前端依赖/构建产物和 Python 缓存：`web/frontend/node_modules/`、`web/frontend/dist/`、`web/frontend/tsconfig.tsbuildinfo`、`__pycache__/`、`*.pyc`。

## Smoke check

不启动服务也可做基础检查：

```bash
python web/backend/smoke_check.py
```

它会验证：扫描资源、分类树、读取示例 TXT、智能搜索、分面 query 参数绑定、非法路径拦截、Git 摘要接口可执行。

## 注意事项 / 风险

- 文件写入统一保存为 UTF-8；读取会兼容多种历史编码。
- 发布接口会执行真实 git 命令和 push，生产环境建议只部署给可信维护者并通过反向代理限制访问。
- Wiki 同步脚本依赖 `mwclient`；若环境未安装，接口会在日志中返回脚本错误。
- `git add` 范围限制为内容目录、仓库根目录 `*.txt`、`wiki/*.py` 和 `web/` 源码/配置白名单；不会加入前端依赖/构建产物，也不会修改 CHM/ZIP 构建逻辑。
