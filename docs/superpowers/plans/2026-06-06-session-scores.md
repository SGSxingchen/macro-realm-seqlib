# 结团统计评分功能 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 给结团统计加 PL 扮演分/影响分、KP 评分:LLM 从战报提取 → 人工可改 → 聚合进玩家表。

**Architecture:** `session_participants` 表加 3 个可空整数列(roleplay_score/impact_score/rating),用 `PRAGMA table_info` + `ALTER TABLE` 做幂等迁移。LLM schema/提示词/校验扩展提取分数;`update_session` 扩展支持逐参与者改分;`list_players` 聚合平均分;前端详情面板改为可编辑、玩家表加 3 列与排序。

**Tech Stack:** Python 3 + FastAPI + sqlite3(后端,unittest 测试);React 19 + TS + Vite(前端,验证用 `npm run build`,无测试框架)。

**设计文档:** `docs/superpowers/specs/2026-06-06-session-scores-design.md`

**实现者必读背景:**

1. 后端核心文件 `web/backend/app/session_stats.py`(service + extractor),API 在 `web/backend/app/main.py`。
2. 数据库无版本号,迁移靠 `_init_db`(约 695 行)里 `create table if not exists` + 本计划新增的 `_migrate`。
3. 后端测试:`cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`。现有 40 个测试不能回归。测试在仓库根目录运行(`tests/test_session_stats.py`、`tests/test_session_stats_api.py`,内部把 `web/backend` 加进 sys.path)。
4. 前端构建:`cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib/web/frontend && npm run build`。若报 "Cannot find native binding":`npm install --no-save @rolldown/binding-linux-x64-gnu@1.0.1` 后重试。
5. 工作区有大量与本任务无关的未提交改动 —— commit 时**只 add 本计划涉及的文件**,绝不 `git add .`。
6. `FakeExtractor`(测试桩)在两个测试文件里各有定义:`extract(*, filename, content, month)` 返回预置 payload 列表。

---

### Task 1: 数据库迁移 + 评分列

**Files:**
- Modify: `web/backend/app/session_stats.py`(`_init_db` 约 695-760 行)
- Test: `tests/test_session_stats.py`

- [ ] **Step 1: 写失败测试 —— 旧库自动加列**

在 `tests/test_session_stats.py` 的 `SessionStatsServiceTest` 类内追加:

```python
    def test_migration_adds_score_columns_to_legacy_db(self):
        import sqlite3
        legacy_path = Path(self.tmp.name) / "legacy.sqlite"
        conn = sqlite3.connect(legacy_path)
        conn.executescript(
            """
            create table session_participants (
                session_id integer not null,
                player_id integer not null,
                role text not null,
                duration_hours real not null,
                reincarnation_count integer not null default 1,
                is_host integer not null default 0,
                primary key(session_id, player_id)
            );
            """
        )
        conn.commit()
        conn.close()

        SessionStatsService(legacy_path)  # 实例化应触发迁移

        conn = sqlite3.connect(legacy_path)
        cols = {row[1] for row in conn.execute("pragma table_info(session_participants)").fetchall()}
        conn.close()
        self.assertIn("roleplay_score", cols)
        self.assertIn("impact_score", cols)
        self.assertIn("rating", cols)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats.py::SessionStatsServiceTest::test_migration_adds_score_columns_to_legacy_db -q`
Expected: FAIL(`roleplay_score` 不在 cols 中)

- [ ] **Step 3: 在新建表里加列 + 加迁移方法**

(a) 在 `session_stats.py` 的 `create table if not exists session_participants (...)` 块里,把 `is_host integer not null default 0,` 之后、`primary key(...)` 之前插入三列:

```sql
                    is_host integer not null default 0,
                    roleplay_score integer,
                    impact_score integer,
                    rating integer,
                    primary key(session_id, player_id)
```

(b) 在 `_init_db` 方法体末尾(`executescript(...)` 调用之后、方法结束前)调用迁移:

```python
            self._migrate_participant_scores(conn)
```

注意:`_init_db` 里 `with self._connect() as conn:` 块内已有 conn,把上面这行放进同一个 `with` 块、`executescript` 之后。

(c) 在 `_init_db` 方法之后新增方法:

```python
    def _migrate_participant_scores(self, conn: sqlite3.Connection) -> None:
        existing = {row[1] for row in conn.execute("pragma table_info(session_participants)").fetchall()}
        for column in ("roleplay_score", "impact_score", "rating"):
            if column not in existing:
                conn.execute(f"alter table session_participants add column {column} integer")
```

- [ ] **Step 4: 跑测试确认通过**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats.py::SessionStatsServiceTest::test_migration_adds_score_columns_to_legacy_db -q`
Expected: PASS

- [ ] **Step 5: 全量回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`
Expected: 41 passed

- [ ] **Step 6: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/backend/app/session_stats.py tests/test_session_stats.py
git commit -m "结团评分:session_participants 加分数列与幂等迁移"
```

---

### Task 2: LLM schema + 提示词 + 校验

**Files:**
- Modify: `web/backend/app/session_stats.py`(`_schema` 约 220-253、`_prompt` 约 166-174、`_validation_error` 约 970-1006)
- Test: `tests/test_session_stats.py`

- [ ] **Step 1: 写失败测试 —— 越界分数进异常**

在 `SessionStatsServiceTest` 内追加:

```python
    def test_out_of_range_score_goes_to_import_errors(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="bad.txt", content="raw")],
            extractor=FakeExtractor([
                {
                    "title": "评分越界",
                    "duration_hours": 3,
                    "kp": {"name": "host", "qq": "1", "rating": 11},
                    "players": [{"name": "p", "qq": None, "roleplay_score": 5, "impact_score": 6}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.assertEqual(result["success_count"], 0)
        self.assertEqual(result["failure_count"], 1)
        errors = self.service.list_errors("2026-05")
        self.assertEqual(len(errors), 1)
        self.assertIn("评分", errors[0]["reason"])

    def test_non_integer_score_is_rejected(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="frac.txt", content="raw")],
            extractor=FakeExtractor([
                {
                    "title": "小数分",
                    "duration_hours": 3,
                    "kp": {"name": "host", "qq": "1", "rating": None},
                    "players": [{"name": "p", "qq": None, "roleplay_score": 8.5, "impact_score": 6}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        self.assertEqual(result["failure_count"], 1)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats.py -k "out_of_range_score or non_integer_score" -q`
Expected: FAIL(当前不校验分数,success_count=1)

- [ ] **Step 3: 扩展 schema**

`_schema()` 里 `kp` 的 `properties`/`required` 改为:

```python
                "kp": {
                    "type": "object",
                    "additionalProperties": False,
                    "required": ["name", "qq", "rating"],
                    "properties": {
                        "name": {"type": "string"},
                        "qq": {"type": ["string", "null"]},
                        "rating": {"type": ["number", "null"]},
                    },
                },
```

`players` 的 items 改为:

```python
                "players": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["name", "qq", "roleplay_score", "impact_score"],
                        "properties": {
                            "name": {"type": "string"},
                            "qq": {"type": ["string", "null"]},
                            "roleplay_score": {"type": ["number", "null"]},
                            "impact_score": {"type": ["number", "null"]},
                        },
                    },
                },
```

- [ ] **Step 4: 扩展提示词**

`_prompt` 的返回串里,在 `players 数组[{name, qq}]、` 之后、`confidence` 之前补充字段说明,并加一句抽取规则。把整段返回替换为:

```python
        return (
            "从下面的 TRPG 结团文本中抽取结构化统计数据，只返回 JSON。"
            "字段：title 字符串、duration_hours 数字、kp 对象{name, qq, rating}、"
            "players 数组[{name, qq, roleplay_score, impact_score}]、confidence 数字、warnings 字符串数组。"
            "无法识别 QQ 时填 null；duration_hours 必须是数字小时。"
            "战报中 PL 名字后通常跟『扮演X 影响X』、KP 后跟『评分X』（X 为 0-10 整数）："
            "把 PL 的扮演分填 roleplay_score、影响分填 impact_score，KP 的评分填 rating；"
            "原文没有写明时填 null，不要编造或估算。"
            f"\n月份：{month}\n文件名：{filename}\n原文：\n{content}"
        )
```

- [ ] **Step 5: 扩展校验**

在 `_validation_error` 方法内,`players` 校验的 for 循环里(现有 `if not isinstance(player, dict) ...` 之后),对每个 player 追加分数校验;并在 kp 校验块后加 rating 校验。先在该方法体顶部确认有 `problems: list[str]`(已有)。

(a) 在 kp 校验(`if not isinstance(kp, dict) ...`)之后插入:

```python
        if isinstance(kp, dict):
            self._collect_score_problem(problems, kp.get("rating"), "KP 评分")
```

(b) 把 players 的 for 循环改为(在原有姓名校验后加分数校验):

```python
            for index, player in enumerate(players, start=1):
                if not isinstance(player, dict) or not self._clean_text(player.get("name")):
                    problems.append(f"玩家列表第 {index} 项缺少姓名")
                    continue
                self._collect_score_problem(problems, player.get("roleplay_score"), f"第 {index} 项扮演分")
                self._collect_score_problem(problems, player.get("impact_score"), f"第 {index} 项影响分")
```

(c) 在 `_validation_error` 方法之后新增静态辅助方法:

```python
    @staticmethod
    def _collect_score_problem(problems: list[str], value: Any, label: str) -> None:
        if value is None or value == "":
            return
        try:
            number = float(value)
        except (TypeError, ValueError):
            problems.append(f"{label}不是数字")
            return
        if not math.isfinite(number) or number != int(number):
            problems.append(f"{label}必须是整数")
        elif number < 0 or number > 10:
            problems.append(f"{label}必须在 0 到 10 之间")
```

（`math` 在文件顶部已 import；`Any` 已 import。）

- [ ] **Step 6: 跑测试确认通过 + 回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`
Expected: 43 passed

- [ ] **Step 7: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/backend/app/session_stats.py tests/test_session_stats.py
git commit -m "结团评分:LLM schema/提示词/校验提取分数"
```

---

### Task 3: 导入时写入分数

**Files:**
- Modify: `web/backend/app/session_stats.py`(`_insert_participant` 约 950-968、`_insert_session` 约 861-907、`get_session_detail` 约 567-578)
- Test: `tests/test_session_stats.py`

- [ ] **Step 1: 写失败测试 —— 导入后详情含分数**

```python
    def test_import_persists_scores_in_detail(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="s.txt", content="raw")],
            extractor=FakeExtractor([
                {
                    "title": "带分团",
                    "duration_hours": 4,
                    "kp": {"name": "kphost", "qq": "1", "rating": 9},
                    "players": [{"name": "pa", "qq": "2", "roleplay_score": 8, "impact_score": 7}],
                    "confidence": 0.9,
                    "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        session_id = result["items"][0]["session_id"]
        detail = self.service.get_session_detail(session_id)
        by_name = {p["name"]: p for p in detail["participants"]}
        self.assertEqual(by_name["pa"]["roleplay_score"], 8)
        self.assertEqual(by_name["pa"]["impact_score"], 7)
        self.assertIsNone(by_name["pa"]["rating"])
        self.assertEqual(by_name["kphost"]["rating"], 9)
        self.assertIsNone(by_name["kphost"]["roleplay_score"])
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats.py -k import_persists_scores -q`
Expected: FAIL(detail 无 roleplay_score 键 / 值不对)

- [ ] **Step 3: `_insert_participant` 加分数参数**

把 `_insert_participant` 整个替换为:

```python
    def _insert_participant(
        self,
        conn: sqlite3.Connection,
        session_id: int,
        player_id: int,
        role: str,
        duration_hours: float,
        *,
        is_host: bool,
        roleplay_score: int | None = None,
        impact_score: int | None = None,
        rating: int | None = None,
    ) -> None:
        conn.execute(
            """
            insert into session_participants(
                session_id, player_id, role, duration_hours, reincarnation_count, is_host,
                roleplay_score, impact_score, rating
            )
            values (?, ?, ?, ?, 1, ?, ?, ?, ?)
            """,
            (session_id, player_id, role, duration_hours, 1 if is_host else 0, roleplay_score, impact_score, rating),
        )
```

- [ ] **Step 4: `_insert_session` 透传分数**

在 `_insert_session` 中:

(a) KP 插入那行(`self._insert_participant(conn, session_id, kp_id, "kp", duration_hours, is_host=True)`)改为:

```python
        self._insert_participant(
            conn, session_id, kp_id, "kp", duration_hours, is_host=True,
            rating=self._coerce_score(kp.get("rating")),
        )
```

(b) PL 循环里的插入(`self._insert_participant(conn, session_id, player_id, "pl", duration_hours, is_host=False)`)改为:

```python
            self._insert_participant(
                conn, session_id, player_id, "pl", duration_hours, is_host=False,
                roleplay_score=self._coerce_score(player.get("roleplay_score")),
                impact_score=self._coerce_score(player.get("impact_score")),
            )
```

(c) 新增静态辅助(放在 `_insert_participant` 之后):

```python
    @staticmethod
    def _coerce_score(value: Any) -> int | None:
        if value is None or value == "":
            return None
        return int(float(value))
```

（校验已在 Task 2 保证 0-10 整数;这里只做存储期类型归一。）

- [ ] **Step 5: `get_session_detail` 返回分数**

`get_session_detail` 里:

(a) 参与者查询的 select 列表加三列。把 `sp.is_host` 那行改为:

```python
                    sp.is_host,
                    sp.roleplay_score,
                    sp.impact_score,
                    sp.rating
```

(b) 组装 participants 的字典推导,每项加三键(`is_host` 那行之后):

```python
                "is_host": bool(p["is_host"]),
                "roleplay_score": p["roleplay_score"],
                "impact_score": p["impact_score"],
                "rating": p["rating"],
```

- [ ] **Step 6: 跑测试 + 回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`
Expected: 44 passed

- [ ] **Step 7: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/backend/app/session_stats.py tests/test_session_stats.py
git commit -m "结团评分:导入写入分数并在详情返回"
```

---

### Task 4: 玩家表聚合平均分

**Files:**
- Modify: `web/backend/app/session_stats.py`(`list_players` 约 466-489、`_player_row` 约 1017-1028)
- Test: `tests/test_session_stats.py`

- [ ] **Step 1: 写失败测试 —— 平均分聚合**

```python
    def test_list_players_aggregates_average_scores(self):
        self.service.import_text_files(
            month="2026-05",
            files=[
                SessionTextFile(filename="a.txt", content="A"),
                SessionTextFile(filename="b.txt", content="B"),
            ],
            extractor=FakeExtractor([
                {
                    "title": "团一", "duration_hours": 3,
                    "kp": {"name": "kp1", "qq": "10", "rating": 8},
                    "players": [{"name": "pa", "qq": "20", "roleplay_score": 6, "impact_score": 4}],
                    "confidence": 0.9, "warnings": [],
                },
                {
                    "title": "团二", "duration_hours": 3,
                    "kp": {"name": "kp1", "qq": "10", "rating": 10},
                    "players": [{"name": "pa", "qq": "20", "roleplay_score": 8, "impact_score": None}],
                    "confidence": 0.9, "warnings": [],
                },
            ]),
            model_name="fake-model",
        )
        players = {row["name"]: row for row in self.service.list_players("2026-05")}
        self.assertAlmostEqual(players["pa"]["avg_roleplay"], 7.0)   # (6+8)/2
        self.assertAlmostEqual(players["pa"]["avg_impact"], 4.0)     # 仅一场有值
        self.assertIsNone(players["pa"]["avg_rating"])               # pa 从未当 KP
        self.assertAlmostEqual(players["kp1"]["avg_rating"], 9.0)    # (8+10)/2
        self.assertIsNone(players["kp1"]["avg_roleplay"])            # kp1 从未当 PL
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats.py -k aggregates_average_scores -q`
Expected: FAIL(KeyError: avg_roleplay)

- [ ] **Step 3: 聚合 SQL 加三列**

`list_players` 的 select 里,`host_hours` 那行之后加三个聚合(注意它后面紧跟 `from`,在 `... as host_hours` 后补逗号):

```python
                    coalesce(sum(case when sp.is_host = 1 then sp.duration_hours else 0 end), 0) as host_hours,
                    avg(case when sp.role = 'pl' then sp.roleplay_score end) as avg_roleplay,
                    avg(case when sp.role = 'pl' then sp.impact_score end) as avg_impact,
                    avg(case when sp.role = 'kp' then sp.rating end) as avg_rating
```

- [ ] **Step 4: `_player_row` 输出三列**

`_player_row` 的返回字典,`host_hours` 那行之后加三键:

```python
            "host_hours": float(row["host_hours"]),
            "avg_roleplay": None if row["avg_roleplay"] is None else float(row["avg_roleplay"]),
            "avg_impact": None if row["avg_impact"] is None else float(row["avg_impact"]),
            "avg_rating": None if row["avg_rating"] is None else float(row["avg_rating"]),
```

- [ ] **Step 5: 跑测试 + 回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`
Expected: 45 passed

- [ ] **Step 6: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/backend/app/session_stats.py tests/test_session_stats.py
git commit -m "结团评分:玩家表聚合平均扮演/影响/评分"
```

---

### Task 5: update_session 支持改分

**Files:**
- Modify: `web/backend/app/session_stats.py`(`update_session` 约 581-605)
- Test: `tests/test_session_stats.py`

- [ ] **Step 1: 写失败测试 —— 改分与越界/错配**

```python
    def _seed_one_session(self):
        result = self.service.import_text_files(
            month="2026-05",
            files=[SessionTextFile(filename="s.txt", content="raw")],
            extractor=FakeExtractor([
                {
                    "title": "团", "duration_hours": 4,
                    "kp": {"name": "kp", "qq": "1", "rating": None},
                    "players": [{"name": "pa", "qq": "2", "roleplay_score": None, "impact_score": None}],
                    "confidence": 0.9, "warnings": [],
                }
            ]),
            model_name="fake-model",
        )
        return result["items"][0]["session_id"]

    def test_update_session_writes_participant_scores(self):
        session_id = self._seed_one_session()
        detail = self.service.get_session_detail(session_id)
        pa = next(p for p in detail["participants"] if p["name"] == "pa")
        kp = next(p for p in detail["participants"] if p["name"] == "kp")
        updated = self.service.update_session(
            session_id,
            participants=[
                {"player_id": pa["id"], "roleplay_score": 9, "impact_score": 7},
                {"player_id": kp["id"], "rating": 8},
            ],
        )
        result = {p["name"]: p for p in updated["participants"]}
        self.assertEqual(result["pa"]["roleplay_score"], 9)
        self.assertEqual(result["pa"]["impact_score"], 7)
        self.assertEqual(result["kp"]["rating"], 8)

    def test_update_session_rejects_out_of_range_score(self):
        session_id = self._seed_one_session()
        pa_id = next(p["id"] for p in self.service.get_session_detail(session_id)["participants"] if p["name"] == "pa")
        with self.assertRaises(ValueError):
            self.service.update_session(session_id, participants=[{"player_id": pa_id, "roleplay_score": 11}])

    def test_update_session_rejects_foreign_player(self):
        session_id = self._seed_one_session()
        with self.assertRaises(ValueError):
            self.service.update_session(session_id, participants=[{"player_id": 999999, "rating": 5}])
```

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats.py -k update_session -q`
Expected: FAIL(`update_session()` 不接受 participants 参数 → TypeError)

- [ ] **Step 3: 扩展 update_session**

把 `update_session` 的签名与方法体替换为(保留原 title/duration 逻辑,新增 participants 处理):

```python
    def update_session(
        self,
        session_id: int,
        *,
        title: str | None = None,
        duration_hours: float | None = None,
        participants: list[dict[str, Any]] | None = None,
    ) -> dict[str, Any] | None:
        updates: list[str] = []
        params: list[Any] = []
        if title is not None:
            clean_title = self._clean_text(title)
            if not clean_title:
                raise ValueError("标题不能为空")
            updates.append("title = ?")
            params.append(clean_title)
        if duration_hours is not None:
            duration = float(duration_hours)
            if not math.isfinite(duration) or duration <= 0:
                raise ValueError("时长必须为正数")
            updates.append("duration_hours = ?")
            params.append(duration)

        with self._connect() as conn:
            if updates:
                params.append(session_id)
                cursor = conn.execute(f"update sessions set {', '.join(updates)} where id = ?", params)
                if cursor.rowcount == 0:
                    return None
                if duration_hours is not None:
                    conn.execute("update session_participants set duration_hours = ? where session_id = ?", (float(duration_hours), session_id))
            elif participants:
                exists = conn.execute("select 1 from sessions where id = ?", (session_id,)).fetchone()
                if exists is None:
                    return None
            for item in participants or []:
                self._apply_participant_scores(conn, session_id, item)
        return self.get_session_detail(session_id)

    def _apply_participant_scores(self, conn: sqlite3.Connection, session_id: int, item: dict[str, Any]) -> None:
        player_id = item.get("player_id")
        if player_id is None:
            raise ValueError("参与者缺少 player_id")
        columns: list[str] = []
        values: list[Any] = []
        for key, label in (("roleplay_score", "扮演分"), ("impact_score", "影响分"), ("rating", "评分")):
            if key not in item:
                continue
            value = self._validate_score(item.get(key), label)
            columns.append(f"{key} = ?")
            values.append(value)
        if not columns:
            return
        values.extend([session_id, player_id])
        cursor = conn.execute(
            f"update session_participants set {', '.join(columns)} where session_id = ? and player_id = ?",
            values,
        )
        if cursor.rowcount == 0:
            raise ValueError("参与者不属于本团")

    @staticmethod
    def _validate_score(value: Any, label: str) -> int | None:
        if value is None or value == "":
            return None
        try:
            number = float(value)
        except (TypeError, ValueError) as exc:
            raise ValueError(f"{label}必须是 0 到 10 的整数") from exc
        if not math.isfinite(number) or number != int(number) or number < 0 or number > 10:
            raise ValueError(f"{label}必须是 0 到 10 的整数")
        return int(number)
```

- [ ] **Step 4: 跑测试 + 回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`
Expected: 48 passed

- [ ] **Step 5: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/backend/app/session_stats.py tests/test_session_stats.py
git commit -m "结团评分:update_session 支持逐参与者改分"
```

---

### Task 6: API 层 —— PATCH 模型 + 端点 + 排序

**Files:**
- Modify: `web/backend/app/main.py`(`SessionStatsSessionUpdateRequest` 约 396-398、PATCH 端点 约 594-606、players 端点 约 471-484)
- Test: `tests/test_session_stats_api.py`

- [ ] **Step 1: 写失败测试 —— PATCH 改分 + 排序**

在 `tests/test_session_stats_api.py` 的 `SessionStatsApiTest` 内追加:

```python
    def test_patch_updates_participant_scores(self):
        old_password = os.environ.get("ADMIN_PASSWORD")
        os.environ["ADMIN_PASSWORD"] = "test-pass"
        try:
            imported = main_app.SESSION_STATS_SERVICE.import_text_files(
                month="2026-05",
                files=[SessionTextFile(filename="r.txt", content="raw")],
                extractor=FakeExtractor([
                    {
                        "title": "团", "duration_hours": 2,
                        "kp": {"name": "host", "qq": "1", "rating": None},
                        "players": [{"name": "p", "qq": "2", "roleplay_score": None, "impact_score": None}],
                        "confidence": 0.9, "warnings": [],
                    }
                ]),
                model_name="fake-model",
            )
            session_id = imported["items"][0]["session_id"]
            detail = self.client.get(f"/api/session-stats/sessions/{session_id}").json()
            p_id = next(x["id"] for x in detail["participants"] if x["name"] == "p")

            self.client.post("/api/admin/login", json={"password": "test-pass"})
            patched = self.client.patch(
                f"/api/session-stats/sessions/{session_id}",
                json={"participants": [{"player_id": p_id, "roleplay_score": 9, "impact_score": 6}]},
            )
            self.assertEqual(patched.status_code, 200)
            updated = {x["name"]: x for x in patched.json()["session"]["participants"]}
            self.assertEqual(updated["p"]["roleplay_score"], 9)
            self.assertEqual(updated["p"]["impact_score"], 6)

            bad = self.client.patch(
                f"/api/session-stats/sessions/{session_id}",
                json={"participants": [{"player_id": p_id, "roleplay_score": 99}]},
            )
            self.assertEqual(bad.status_code, 422)
        finally:
            if old_password is None:
                os.environ.pop("ADMIN_PASSWORD", None)
            else:
                os.environ["ADMIN_PASSWORD"] = old_password

    def test_players_endpoint_accepts_score_sort(self):
        response = self.client.get("/api/session-stats/players?month=2026-05&sort=roleplay")
        self.assertEqual(response.status_code, 200)
```

注:`roleplay_score=99` 超过 le=10,Pydantic 校验返回 422(在 service 之前拦截)。

- [ ] **Step 2: 跑测试确认失败**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/test_session_stats_api.py -k "participant_scores or score_sort" -q`
Expected: FAIL(participants 字段被忽略 / sort=roleplay 报 422)

- [ ] **Step 3: Pydantic 模型**

`main.py` 里把 `SessionStatsSessionUpdateRequest` 替换为(并在其上方新增子模型):

```python
class SessionStatsParticipantScoreUpdate(BaseModel):
    player_id: int
    roleplay_score: int | None = Field(default=None, ge=0, le=10)
    impact_score: int | None = Field(default=None, ge=0, le=10)
    rating: int | None = Field(default=None, ge=0, le=10)


class SessionStatsSessionUpdateRequest(BaseModel):
    title: str | None = Field(default=None, max_length=200)
    duration_hours: float | None = Field(default=None, gt=0)
    participants: list[SessionStatsParticipantScoreUpdate] | None = None
```

- [ ] **Step 4: PATCH 端点透传 participants**

把 PATCH 端点替换为:

```python
@app.patch("/api/session-stats/sessions/{session_id}")
def session_stats_update_session(session_id: int, body: SessionStatsSessionUpdateRequest, _admin: None = Depends(require_admin)):
    if body.title is None and body.duration_hours is None and not body.participants:
        raise HTTPException(400, "没有可更新字段")
    participants = [p.model_dump(exclude_unset=True) for p in body.participants] if body.participants else None
    try:
        detail = session_stats_service().update_session(
            session_id,
            title=body.title,
            duration_hours=body.duration_hours,
            participants=participants,
        )
    except ValueError as exc:
        raise HTTPException(400, str(exc)) from exc
    if not detail:
        raise HTTPException(404, "团记录不存在")
    return {"ok": True, "session": detail}
```

注:`exclude_unset=True` 保证「没传的分数列」不出现在 dict 里(service 只更新出现的键),「显式传 null」会保留为 None(清空)。

- [ ] **Step 5: players 端点加排序**

把 `session_stats_players` 替换为:

```python
@app.get("/api/session-stats/players")
def session_stats_players(month: str, sort: Literal["games", "hours", "hosts", "name", "roleplay", "impact", "rating"] = "hours"):
    month = validate_stats_month(month)
    items = session_stats_service().list_players(month)

    def score_key(field: str):
        return lambda x: (x[field] is not None, x[field] if x[field] is not None else 0)

    if sort == "games":
        items.sort(key=lambda x: (-x["game_count"], -x["game_hours"], x["name"]))
    elif sort == "hosts":
        items.sort(key=lambda x: (-x["host_count"], -x["host_hours"], x["name"]))
    elif sort == "name":
        items.sort(key=lambda x: x["name"])
    elif sort in ("roleplay", "impact", "rating"):
        field = {"roleplay": "avg_roleplay", "impact": "avg_impact", "rating": "avg_rating"}[sort]
        items.sort(key=lambda x: x["name"])
        items.sort(key=score_key(field), reverse=True)
    else:
        items.sort(key=lambda x: (-x["game_hours"], -x["game_count"], x["name"]))
    return {"items": items, "count": len(items), "month": month}
```

（分数排序:有值的排前、降序;无值(None)沉底。先按 name 稳定排再按分数降序,保证同分按名稳定。）

- [ ] **Step 6: 跑测试 + 回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q`
Expected: 50 passed

- [ ] **Step 7: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/backend/app/main.py tests/test_session_stats_api.py
git commit -m "结团评分:PATCH 改分接口与玩家表分数排序"
```

---

### Task 7: 前端类型

**Files:**
- Modify: `web/frontend/src/types.ts`(约 168-208)

- [ ] **Step 1: 扩展类型**

(a) `SessionStatsPlayerSort` 改为:

```typescript
export type SessionStatsPlayerSort = 'hours' | 'games' | 'hosts' | 'name' | 'roleplay' | 'impact' | 'rating';
```

(b) `SessionStatsPlayer` 的字段末尾(`host_hours: number;` 之后)加:

```typescript
  avg_roleplay: number | null;
  avg_impact: number | null;
  avg_rating: number | null;
```

(c) `SessionStatsParticipant` 的字段末尾(`is_host?: boolean | null;` 之后)加:

```typescript
  roleplay_score?: number | null;
  impact_score?: number | null;
  rating?: number | null;
```

- [ ] **Step 2: 构建验证**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib/web/frontend && npm run build`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/frontend/src/types.ts
git commit -m "结团评分:前端类型加分数字段"
```

---

### Task 8: 前端玩家表 3 列 + 排序胶囊

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`(玩家表 section 与排序胶囊)
- Modify: `web/frontend/src/style.css`(`.stats-table` min-width)

- [ ] **Step 1: 排序胶囊加三项**

在玩家表 `sort-chips` 的常量数组里追加三项。把该数组改为:

```tsx
              {([['hours', '游戏时长'], ['games', '游戏次数'], ['hosts', '主持次数'], ['roleplay', '平均扮演'], ['impact', '平均影响'], ['rating', '平均评分'], ['name', '玩家名']] as const).map(([value, label]) => (
```

- [ ] **Step 2: 表头加三列**

在玩家表 `<thead>` 的 `<th className="col-num">主持</th>` 之后加三列:

```tsx
              <th className="col-num">平均扮演</th>
              <th className="col-num">平均影响</th>
              <th className="col-num">平均评分</th>
```

- [ ] **Step 3: 表体加三列 + 格式化函数**

(a) 在 `SessionStats.tsx` 组件外、`formatHours` 旁新增:

```tsx
function formatScore(value: number | null | undefined) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '—';
  return value.toFixed(1);
}
```

(b) 在玩家行 `<td className="col-num">{player.host_count}</td>` 之后加三格:

```tsx
                  <td className="col-num">{formatScore(player.avg_roleplay)}</td>
                  <td className="col-num">{formatScore(player.avg_impact)}</td>
                  <td className="col-num">{formatScore(player.avg_rating)}</td>
```

- [ ] **Step 4: 表格 min-width 调大**

`web/frontend/src/style.css` 里 `.stats-table { ... min-width: 640px; }` 改为 `min-width: 820px;`(10 列需要更多横向空间)。

- [ ] **Step 5: 构建验证**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib/web/frontend && npm run build`
Expected: PASS

- [ ] **Step 6: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/frontend/src/components/SessionStats.tsx web/frontend/src/style.css
git commit -m "结团评分:玩家表加平均分列与排序"
```

---

### Task 9: 前端详情面板编辑分数

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`(state、saveSessionDetail、openSessionDetail、participants JSX)
- Modify: `web/frontend/src/style.css`(参与者编辑行样式)

- [ ] **Step 1: 加分数草稿 state**

在组件 state 区(`editDuration` 附近)加:

```tsx
  const [editScores, setEditScores] = useState<Record<string, { roleplay: string; impact: string; rating: string }>>({});
```

- [ ] **Step 2: 打开详情时初始化草稿**

`openSessionDetail` 里成功拿到 `detail` 后(`setEditDuration(...)` 之后)加:

```tsx
      const draft: Record<string, { roleplay: string; impact: string; rating: string }> = {};
      for (const p of detail.participants) {
        draft[String(p.id)] = {
          roleplay: typeof p.roleplay_score === 'number' ? String(p.roleplay_score) : '',
          impact: typeof p.impact_score === 'number' ? String(p.impact_score) : '',
          rating: typeof p.rating === 'number' ? String(p.rating) : '',
        };
      }
      setEditScores(draft);
```

- [ ] **Step 3: 保存时携带 participants**

把 `saveSessionDetail` 的 body 构造与请求段替换。在 `const duration = Number(editDuration); ...` 校验之后,组装分数并校验:

```tsx
    const scorePayload: Array<{ player_id: number | string; roleplay_score?: number | null; impact_score?: number | null; rating?: number | null }> = [];
    for (const p of sessionDetail.participants) {
      const draft = editScores[String(p.id)];
      if (!draft) continue;
      const entry: { player_id: number | string; roleplay_score?: number | null; impact_score?: number | null; rating?: number | null } = { player_id: p.id };
      let changed = false;
      const fields: Array<['roleplay' | 'impact' | 'rating', 'roleplay_score' | 'impact_score' | 'rating', number | null | undefined]> = [
        ['roleplay', 'roleplay_score', p.roleplay_score],
        ['impact', 'impact_score', p.impact_score],
        ['rating', 'rating', p.rating],
      ];
      for (const [draftKey, apiKey, original] of fields) {
        const raw = draft[draftKey].trim();
        let next: number | null;
        if (raw === '') {
          next = null;
        } else {
          const n = Number(raw);
          if (!Number.isInteger(n) || n < 0 || n > 10) {
            setDetailError('评分必须是 0 到 10 的整数。');
            return;
          }
          next = n;
        }
        const orig = typeof original === 'number' ? original : null;
        if (next !== orig) { entry[apiKey] = next; changed = true; }
      }
      if (changed) scorePayload.push(entry);
    }
```

然后把 PATCH 请求的 body 改为带 participants(仅在有改动时传):

```tsx
      const result = await api<SessionStatsSessionPatchResponse>(`/api/session-stats/sessions/${routePath(String(sessionDetail.id))}`, {
        method: 'PATCH',
        body: JSON.stringify({
          title: editTitle.trim(),
          duration_hours: duration,
          ...(scorePayload.length ? { participants: scorePayload } : {}),
        }),
      });
```

保存成功后重置草稿。在 `setEditDuration(...)`(成功分支)之后加:

```tsx
      const nextDraft: Record<string, { roleplay: string; impact: string; rating: string }> = {};
      for (const p of result.session.participants) {
        nextDraft[String(p.id)] = {
          roleplay: typeof p.roleplay_score === 'number' ? String(p.roleplay_score) : '',
          impact: typeof p.impact_score === 'number' ? String(p.impact_score) : '',
          rating: typeof p.rating === 'number' ? String(p.rating) : '',
        };
      }
      setEditScores(nextDraft);
```

- [ ] **Step 4: 参与者 JSX 改为可编辑**

把 `stats-detail-participants` 里 `stats-participant-list` 的 map 整体替换为:

```tsx
                  <div className="stats-participant-list">
                    {sessionDetail.participants.map(participant => {
                      const draft = editScores[String(participant.id)] || { roleplay: '', impact: '', rating: '' };
                      const setField = (key: 'roleplay' | 'impact' | 'rating', value: string) =>
                        setEditScores(prev => ({ ...prev, [String(participant.id)]: { ...draft, [key]: value } }));
                      return (
                        <div className="participant-edit" key={participant.id}>
                          <b>{participant.name || '未命名'}{participant.is_host ? ' · KP' : ''}</b>
                          <small>{participant.qq || '未记录 QQ'} · {formatHours(participant.duration_hours)} 小时 · 轮回 {participant.reincarnation_count ?? 0}</small>
                          <div className="participant-scores">
                            {participant.is_host ? (
                              <label><span>评分</span><input type="number" min="0" max="10" step="1" value={draft.rating} onChange={e => setField('rating', e.target.value)} /></label>
                            ) : (
                              <>
                                <label><span>扮演</span><input type="number" min="0" max="10" step="1" value={draft.roleplay} onChange={e => setField('roleplay', e.target.value)} /></label>
                                <label><span>影响</span><input type="number" min="0" max="10" step="1" value={draft.impact} onChange={e => setField('impact', e.target.value)} /></label>
                              </>
                            )}
                          </div>
                        </div>
                      );
                    })}
                  </div>
```

- [ ] **Step 5: 样式**

`web/frontend/src/style.css` 的 `/* 团详情 ... */` 区块里,`.stats-participant-list span { ... }` 规则之后追加:

```css
.participant-edit {
  display: grid;
  gap: 4px;
  min-width: 0;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-s);
  background: var(--panel-2);
  padding: 8px 10px;
}
.participant-edit b { color: var(--text); font-weight: 600; font-size: 13px; }
.participant-edit small { color: var(--muted); font-size: 11px; }
.participant-scores { display: flex; gap: 8px; margin-top: 4px; }
.participant-scores label { display: inline-flex; align-items: center; gap: 4px; }
.participant-scores span { color: var(--muted); font-size: 11px; }
.participant-scores input { width: 52px; padding: 3px 6px; font-size: 12px; }
```

- [ ] **Step 6: 构建验证**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib/web/frontend && npm run build`
Expected: PASS

- [ ] **Step 7: Commit**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib
git add web/frontend/src/components/SessionStats.tsx web/frontend/src/style.css
git commit -m "结团评分:详情面板逐人编辑分数"
```

---

### Task 10: 手动验证(开发服务器)

**前置:** 后端(`ADMIN_PASSWORD=dev-admin python3 -m uvicorn app.main:app --port 8000`,在 `web/backend` 目录)与前端(`npm run dev`)都在跑。本机用系统 Python 跑后端(WSL 下 `.venv` 是 Windows 的不可用)。

- [ ] **Step 1: 造带分数的测试数据**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 - <<'PY'
import sqlite3, hashlib, json
from datetime import datetime
conn = sqlite3.connect('web/data/session_stats.sqlite')
now = datetime.now().isoformat()
b = conn.execute("insert into import_batches(month,file_count,success_count,failure_count,model_name,created_at) values('2026-06',2,2,0,'seed',?)", (now,)).lastrowid
def player(name, qq):
    return conn.execute("insert into players(name,qq,created_at,updated_at) values(?,?,?,?)", (name,qq,now,now)).lastrowid
kp = player('评分KP','9001'); pa = player('扮演甲','9002'); pb = player('影响乙','9003')
for i,(rp_a,im_a,rt) in enumerate([(8,6,9),(6,4,7)]):
    t=f'评分测试团{i+1}'; h=hashlib.sha1(t.encode()).hexdigest()
    sid=conn.execute("insert into sessions(batch_id,month,title,duration_hours,kp_player_id,source_filename,content_hash,model_name,confidence,raw_payload,created_at) values(?,?,?,?,?,?,?,?,?,?,?)",
        (b,'2026-06',t,4,kp,f'{t}.txt',h,'seed',0.9,json.dumps({'seed':True}),now)).lastrowid
    conn.execute("insert into session_participants values(?,?,?,?,1,1,?,?,?)",(sid,kp,'kp',4,None,None,rt))
    conn.execute("insert into session_participants values(?,?,?,?,1,0,?,?,?)",(sid,pa,'pl',4,rp_a,im_a,None))
    conn.execute("insert into session_participants values(?,?,?,?,1,0,?,?,?)",(sid,pb,'pl',4,5,8,None))
conn.commit(); print('seeded', conn.execute("select count(*) from sessions where month='2026-06'").fetchone())
PY
```

- [ ] **Step 2: 玩家表三列与排序**

打开 `http://localhost:5173/?tab=stats`,切到 2026-06:
- 玩家表出现「平均扮演 / 平均影响 / 平均评分」三列,扮演甲 = 7.0/5.0/—,评分KP = —/—/8.0
- 点排序胶囊「平均扮演」「平均评分」,排序变化、无分者(—)沉底

- [ ] **Step 3: 详情编辑分数**

点任一团「查看」→ 详情面板:
- PL 行显示「扮演 [_] 影响 [_]」输入框并带原值;KP 行显示「评分 [_]」
- 改一个 PL 的扮演分 → 点「保存」→ 成功后玩家表平均分随之变化
- 输入 11 或 8.5 → 保存时前端报「评分必须是 0 到 10 的整数」

- [ ] **Step 4: 旧库迁移验证**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 - <<'PY'
import sqlite3, tempfile, os, sys
sys.path.insert(0, 'web/backend')
from app.session_stats import SessionStatsService
p = tempfile.mktemp(suffix='.sqlite')
c = sqlite3.connect(p)
c.executescript("create table session_participants(session_id integer,player_id integer,role text,duration_hours real,reincarnation_count integer default 1,is_host integer default 0,primary key(session_id,player_id));")
c.commit(); c.close()
SessionStatsService(p)
c = sqlite3.connect(p)
print('cols:', sorted(r[1] for r in c.execute('pragma table_info(session_participants)')))
os.unlink(p)
PY
```
Expected: cols 含 `roleplay_score`、`impact_score`、`rating`

- [ ] **Step 5: 暗/亮主题核对**

详情编辑行、玩家表新列在暗/亮两套主题下对比度正常,输入框可读。

- [ ] **Step 6: 清理测试数据**

```bash
cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 - <<'PY'
import sqlite3
c = sqlite3.connect('web/data/session_stats.sqlite')
for t in ('session_participants','sessions','players','import_batches','import_errors'):
    c.execute(f'delete from {t}')
c.commit(); print('cleared')
PY
```

- [ ] **Step 7: 最终回归**

Run: `cd /mnt/d/WorkSpace/ClaudeCode/macro-realm-seqlib && python3 -m pytest tests/ -q && cd web/frontend && npm run build`
Expected: 50 passed + 构建通过

---

## Self-Review 记录

- **Spec coverage:** 数据模型+迁移 → Task 1;LLM schema/提示词/校验 → Task 2;导入写入 → Task 3;玩家表聚合 → Task 4;update_session 改分 → Task 5;PATCH 接口/Pydantic/排序 → Task 6;前端类型 → Task 7;玩家表列+排序 → Task 8;详情编辑 → Task 9;验证(含迁移/暗亮/回归)→ Task 10。全覆盖。
- **类型一致性:** `roleplay_score`/`impact_score`/`rating` 三个列名在 schema/插入/详情/聚合/PATCH/前端一致;`avg_roleplay`/`avg_impact`/`avg_rating` 在 SQL/_player_row/types/玩家表一致;`SessionStatsPlayerSort` 七值与后端 `Literal` 七值对齐。`_collect_score_problem`(Task 2,导入校验)与 `_validate_score`(Task 5,编辑校验)是两个不同用途的方法,命名区分,均存在。
- **占位符扫描:** 无 TBD/TODO;所有代码块完整。
- **校验层次:** 导入走 `_collect_score_problem`(进异常列表,不抛);编辑走 Pydantic ge/le(422)+ service `_validate_score`(400 兜底),两条路径分明。
