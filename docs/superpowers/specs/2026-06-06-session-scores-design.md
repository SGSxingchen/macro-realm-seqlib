# 结团统计评分功能 设计文档

日期:2026-06-06
状态:已确认
范围:后端(schema/迁移/提取/聚合/PATCH)+ 前端(详情编辑 + 玩家表列与排序)

## 背景与目标

结团统计目前只记录团的标题、时长、KP、PL、轮回、置信度。战报文件里其实还写了每个 PL 的「扮演 X 影响 X」、KP 的「评分 X」(X 为 0-10 整数)。本功能把这些分数纳入:LLM 从战报提取 → 人工可改 → 聚合进玩家统计。

要点:
- **来源**:LLM 提取 + 人工编辑(战报里有明确格式,提取有据可依)。
- **字段归属**:PL 有扮演分、影响分;KP 有评分。
- **取值**:0-10 整数,可空(战报没写或未打分时为 NULL)。
- **聚合**:进玩家统计表,展示平均分并可排序。

## 数据模型与迁移

`session_participants` 表新增 3 个可空整数列:

| 列 | 含义 | 适用角色 |
|---|---|---|
| `roleplay_score` | 扮演分 0-10 | PL |
| `impact_score` | 影响分 0-10 | PL |
| `rating` | 评分 0-10 | KP |

均允许 NULL。按角色填:PL 行填 `roleplay_score`/`impact_score`、`rating` 留空;KP 行填 `rating`、另两列留空。

**迁移机制(新增):** 当前 `_init_db` 只有 `create table if not exists`,已有库不会自动加列。新增轻量迁移:建表后用 `PRAGMA table_info(session_participants)` 读现有列名,缺哪列就 `ALTER TABLE session_participants ADD COLUMN <col> integer`。幂等,旧库平滑升级,老数据该 3 列为 NULL。迁移在 `_init_db` 内、建表语句之后执行。

**raw_payload 不变:** 与时长编辑一致,人工改的分数写进列,`raw_payload` 保留 LLM 原始输出。

## LLM 提取(schema + 提示词)

**`_schema()` 扩展:**
- `players[]` 每项 `properties` 加 `roleplay_score`、`impact_score`,类型 `{"type": ["number", "null"]}`,并加入该项的 `required` 数组。
- `kp` 的 `properties` 加 `rating`,类型 `{"type": ["number", "null"]}`,加入 `kp.required`。

(strict 模式要求所有属性必填,故加入 required;值允许 null 表达「战报没写」。)

**提示词扩展(`_prompt`):** 增加说明——战报中 PL 名字后通常跟「扮演 X 影响 X」、KP 后跟「评分 X」(X 为 0-10 整数);提取到对应字段;没有就给 null,**不要编造或估算**。

**校验(`_validation_error`):** 对 `players[].roleplay_score`、`players[].impact_score`、`kp.rating`,若非 null:必须能转 int 且 0 ≤ 值 ≤ 10,否则计入问题(进异常列表),与现有 confidence 校验同风格。允许浮点形态的整数(如 `8.0`),取整存储;非整数(如 `8.5`)判为问题。

**入库(`_insert_session` / `_insert_participant`):** `_insert_participant` 增加可选参数 `roleplay_score`/`impact_score`/`rating`(默认 None),写进新列。`_insert_session` 从 payload 取:KP 取 `kp.get("rating")`;每个 PL 取 `player.get("roleplay_score")`/`player.get("impact_score")`。

## 编辑 UX(团详情面板)

详情面板的参与者从只读改为**每人一行带评分输入框**:

- **PL 行**:玩家名 + QQ + 「扮演 [_]」「影响 [_]」两个数字输入(0-10 整数,可空);时长/轮回保持只读小字。
- **KP 行**:玩家名 + QQ + 「评分 [_]」一个输入 + 标「KP」。

**保存:** 沿用面板底部唯一的「保存」按钮,一次写入标题、时长、所有参与者分数改动。不新增按钮。

**前端状态:** 进入详情时,用 participants 初始化一个可编辑的分数草稿 state(按 player_id 索引)。保存时与原值比对,把有变化的参与者组成数组提交。

**接口扩展:** `PATCH /api/session-stats/sessions/{id}` body 增加可选字段:

```json
{
  "title": "...",
  "duration_hours": 4.5,
  "participants": [
    { "player_id": 12, "roleplay_score": 8, "impact_score": 7 },
    { "player_id": 3, "rating": 9 }
  ]
}
```

- `participants` 可选;每项必含 `player_id`,其余字段可选,只更新传入的列。
- 分数:整数 0-10,或 null(对应前端空字符串→清空)。非法值 → HTTP 400。
- `player_id` 不属于该 session → 跳过该项(不报错)或 400?**采用 400**,提示「参与者不属于本团」,避免静默吞错。
- 标题/时长逻辑不变;三者可任意组合提交;全空 → 维持现有「没有可更新字段」400。

**后端实现:** `update_session` 增加 `participants` 参数。逐条 `update session_participants set <col>=? where session_id=? and player_id=?`,rowcount=0 视为不属于本团 → 抛 ValueError。校验在 service 层(整数 0-10/None)。

**Pydantic 模型:** 新增 `SessionStatsParticipantScoreUpdate`(player_id:int、roleplay_score/impact_score/rating:int|None,带 ge=0/le=10);`SessionStatsSessionUpdateRequest` 加 `participants: list[...] | None`。

## 聚合进玩家表 + 显示

**玩家表新增 3 列**(接在「主持」后):平均扮演、平均影响、平均评分。

- 平均扮演 = 该玩家所有 PL 场 `roleplay_score` 的平均(NULL 不计)。
- 平均影响 = 同上,`impact_score`。
- 平均评分 = 该玩家所有 KP 场 `rating` 的平均。
- 显示保留 1 位小数;无评分场次显示「—」。

**聚合 SQL(`list_players`):** 增加三个聚合表达式:

```sql
avg(case when sp.role = 'pl' then sp.roleplay_score end) as avg_roleplay,
avg(case when sp.role = 'pl' then sp.impact_score end) as avg_impact,
avg(case when sp.role = 'kp' then sp.rating end) as avg_rating
```

SQLite `avg` 自动忽略 NULL。`_player_row` 把三者读出(可能为 None)。

**排序(`session_stats_players` 端点):** `sort` 新增 `roleplay`/`impact`/`rating` 三值,按对应平均值降序(None 视为最低,排在后面),次级键沿用 name。`SessionStatsPlayerSort` 类型同步。

**前端:**
- `types.ts`:`SessionStatsPlayer` 加 `avg_roleplay/avg_impact/avg_rating`(`number | null`);`SessionStatsParticipant` 加 `roleplay_score/impact_score/rating`(`number | null`);`SessionStatsPlayerSort` 加三个枚举值。
- 玩家表加 3 列(数值列,1 位小数 / 「—」),排序胶囊加三项「平均扮演 / 平均影响 / 平均评分」,共 7 项,窄屏 flex-wrap 换行。
- 表格从 7 列变 10 列,沿用 `stats-table-wrap` 横向滚动;`min-width` 适当调大。

## 错误处理

- 提取分数越界/非整数 → 进异常列表(不阻断整批导入)。
- PATCH 分数非法或 player_id 不属于本团 → HTTP 400,前端弹出错误,不静默。
- 聚合无数据 → 前端「—」,不报错。

## 验证

1. 后端:`pytest tests/`(现有 40 测试不回归)+ 新增评分相关单测(提取/校验/迁移/聚合/PATCH 至少各一条)。
2. `npm run build` 通过。
3. 开发服务器实跑:造带分数的测试数据 → 玩家表三列与排序、详情编辑保存、空分数显示「—」、旧库迁移(用迁移前的库验证自动加列)。
