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

# 维护后台写操作必需；结团统计也支持社区管理员/协管 QQ 登录管理
export ADMIN_PASSWORD='change-me'
# 可选：独立 cookie 签名密钥；默认从 ADMIN_PASSWORD 派生
export ADMIN_SECRET='long-random-secret'

# QQ 私聊验证码登录与社区留言板必需：后端调用 https://trpg.chordvers.org/onebot-sequences 的 Bearer 密钥
export ONEBOT_SEQUENCE_TOKEN='your-onebot-sequence-bearer-token'
# 可选：自定义验证码服务基址（默认 https://trpg.chordvers.org/onebot-sequences）
# export ONEBOT_SEQUENCE_BASE_URL='https://trpg.chordvers.org/onebot-sequences'
# 可选：社区管理员 QQ 白名单（逗号或空格分隔），对应 QQ 登录后可管理社区与结团统计；社区协管也有结团统计管理权限
export COMMUNITY_ADMIN_QQS='1348697221,853304398'
# 可选：独立社区会话 HMAC 密钥与 SQLite 存储路径（默认 web/data/community.sqlite）
# export COMMUNITY_SECRET='long-random-community-secret'

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
# 群结团文件后台自动增量同步：默认开启，同时配置 ONEBOT_SEQUENCE_TOKEN 与 OPENAI_API_KEY（或 THIRD_PARTY_API_KEY）时生效
# export SESSION_REPORT_AUTO_SYNC='1'      # 设为 0 关闭
# export SESSION_REPORT_AUTO_SYNC_SECONDS='300' # 检查间隔，默认 300 秒（最小 60 秒），跟进最新 2 个月份目录

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
- `GET /api/resources/{path}`：资源详情，统一解码入口兼容 `utf-8/utf-8-sig/utf-16/gbk/gb2312/big5`；正式资源写回统一保存为 UTF-8。
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
- `PATCH /api/session-stats/sessions/{id}`：编辑 `title` / `duration_hours`，支持通过 `kp: { name, qq }` 与 `players: [{ name, qq }]` 补录/修正参与者。
- `DELETE /api/session-stats/sessions/{id}`：删除团记录。
- `POST /api/session-stats/import-jobs`：上传 TXT 创建异步导入任务（multipart：month/concurrency/files）。
- `GET /api/session-stats/import-jobs/{job_id}`：轮询上传导入或群文件同步任务进度。
- `GET /api/session-stats/group-source`：读取 `onebot-sequence-cache` 自动识别的各月「XX月结团文件」目录列表及后台自动同步状态，随月份增长自动识别新月份目录。
- `POST /api/session-stats/sync-group`：按月从群文件缓存增量拉取未入库或内容已修订的结团报告，调用现有 LLM 自动填表。body: `{ month, concurrency, force_retry_errors }`（并发默认 2、范围 1–6，失败文件默认不重试）；返回异步 `job_id`，复用上述任务轮询接口。
- `POST /api/session-stats/dedupe?month=`：清理内容完全相同的重复团记录。
- `GET /api/session-stats/errors?month=`：导入异常列表（需结团统计管理权限）。

编辑、删除、导入、同步、去重、任务轮询与异常查询支持 `ADMIN_PASSWORD` 后台登录，或使用 `COMMUNITY_ADMIN_QQS` 白名单管理员/社区协管 QQ 登录。

导入需要配置 `OPENAI_API_KEY`（或 `THIRD_PARTY_API_KEY`），群文件拉取还需 `ONEBOT_SEQUENCE_TOKEN`，见上方「启动后端」。`SESSION_REPORT_AUTO_SYNC` 默认 `1` 开启，在上述密钥均配置时生效，每 `SESSION_REPORT_AUTO_SYNC_SECONDS=300` 秒增量跟进最新 2 个月份目录；设为 `0` 可关闭，失败文件可通过手动同步的 `force_retry_errors` 重试。

导入会对原结团文件的常见瑕疵自动修正：时长缺失或无效时从文件名/正文的「时长」「总计 XXH」「XX小时」「XXh结团」等多种写法兜底提取；剥离 KP 复合 `+卡名` 与姓名中混写的 QQ；过滤局部乱码产生的空玩家项，保留有效参与者。同月同内容（`content_hash`）重复导入会自动跳过（SKIP）；群文件以 `source_file_id` 识别同一文件，内容修订后覆盖原团记录，避免重复记团。

### 规范化审核

规范化审核是单独入口，不挂在前台导航，也不使用后台密码。访问 `/normalize-review` 可查看全部审核任务；如果链接带 `?id=...`，页面会默认聚焦该任务，但仍保留全部任务列表。

- `GET /api/normalization/reviews`：列出全部规范化审核任务。
- `POST /api/normalization/reviews`：创建审核任务。body: `{ resource_path, normalized_content, original_content?, title?, note? }`
- `GET /api/normalization/reviews/{id}`：读取某个审核任务的规范化前后文本。
- `POST /api/normalization/reviews/{id}/sign`：签名通过。body: `{ signer, note? }`

审核任务以 JSON 文件保存在 `web/normalization_reviews/`。每个任务记录规范化前文本、规范化后文本、签名列表和状态；当前规则为凑齐 1 个签名即通过。

### 主群官方 QQ Bot 短码登录与社区（轮回者论坛 / 抗改名序列评论区）

独立持久化于 `web/data/community.sqlite`（不改动任何 `序列库/`、`荣誉室/` TXT 源文件，在线自动兼容存量表升级）：

- **主群官方 Bot 网页登录**：网页生成绑定到当前浏览器的 6 位码。在宏观界域主群 @阿瓦娜，发送网页显示的 `/登录 验证码`，回到网页核对身份。首次注册填写展示昵称，之后沿用论坛昵称；会话有效期为 30 天。
  - `GET /api/auth/me` 读取当前浏览器的身份与认证配置。
  - `POST /api/auth/web-challenges` 生成或恢复有效期 5 分钟的验证请求。
  - `POST /api/auth/bot/verify-code` 由海豹插件 `tools/sealdice_forum_auth.js` 调用。后端校验桥接密钥、主群白名单和机器人身份。
  - `GET /api/auth/web-challenges/current` 根据 HttpOnly Cookie 查询核验状态，返回待确认的 QQ 身份。
  - `POST /api/auth/web-challenges/consume` 校验当前浏览器和身份确认凭据，在同一事务中消费请求并建立会话。
  - `PATCH /api/auth/profile`：已登录轮回者自助修改论坛展示昵称（同步更新历史发言作者昵称）。
  - `POST /api/auth/logout`：销毁当前浏览器会话。
- **轮回者论坛与抗重排序列评论区**（未登录公开只读，交易群轮回者登录后发帖、跟帖、点赞共鸣、关联能力卡锚点）：
  - `GET /api/community/home`：论坛紧凑主题列表（置顶优先 + 最后回复顶帖排序）与右侧全站最新序列评论动态流。
  - `GET /api/community/threads/{id}`：读取单帖 `1F / 2F / 3F` 楼层详情（直达链接 `?thread={id}`）。
  - `POST /api/community/board`：发布新论坛主题（`title + content`）或跟帖回复（`reply_to_id`）。
  - `GET /api/community/comments?path=序列库/...&anchor_id=&anchor_label=&sort=latest|hot|oldest`：按去序号稳定键 `resource_key` + `community_resource_aliases` 历史路径映射表读取序列评论（经历 `renumber_resources.py` 批量重排、跨侧别移动或后台重命名均不断链，能力卡升降阶按去品阶名称回退关联）。
  - `POST /api/community/comments`：发表序列评论（可选绑定 `anchor_id` / `anchor_label` 到具体能力卡，或 `reply_to_id` 回复）。
  - `POST /api/community/posts/{id}/like`：切换点赞/共鸣。
  - `DELETE /api/community/posts/{id}`：作者本人或管理员删除发言（主楼删除时同步归档下属回复）。
  - `POST /api/community/posts/{id}/pin`：管理员置顶/取消置顶主楼。
- **社区管理后台**（支持 `ADMIN_PASSWORD` 维护控制台「社区管理」页签，以及 `COMMUNITY_ADMIN_QQS` 白名单/协管 QQ）：
  - `GET /api/community/moderation/posts`：按版块、状态（含已删归档）、关键词/QQ 审计全部发言。
  - `GET /api/community/moderation/users`：查看已登录群成员、角色卡数、发言数与封禁状态。
  - `POST /api/community/moderation/ban`：按时长（24h / 72h / 7d / 永久）禁言或解封指定 QQ（支持未登录过的 QQ 预先拉黑）。
  - `POST /api/community/moderation/users/{qq}/role`：授予或撤销社区协管身份。
  - `POST /api/community/moderation/announcement`：直接以系统或维护者身份发布主页置顶公告。

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
