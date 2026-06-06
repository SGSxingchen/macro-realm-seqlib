# 结团统计页 UI 重设计 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 把结团统计页从表单式堆叠改为仪表盘式布局(限宽主列、折叠导入面板、KPI 卡、排序胶囊、前端翻页),只改前端两个文件。

**Architecture:** `SessionStats.tsx` 保留全部数据请求/导入轮询/编辑删除逻辑函数,只重排 JSX 和新增 4 个 state;`style.css` 删除旧 stats 块、重写新 stats 块,全部基于现有 CSS 变量适配暗/亮主题。新增一个可复用 `Pagination` 组件。

**Tech Stack:** React 19 + TypeScript + Vite(rolldown)。无前端测试框架,验证手段是 `npm run build`(含 tsc 类型检查)+ 开发服务器手动核对。

**设计文档:** `docs/superpowers/specs/2026-06-06-session-stats-ui-redesign-design.md`

**重要背景(实现者必读):**

1. `web/frontend/src/style.css` 有**两个** `/* ---------- 结团统计 ---------- */` 块:
   - 旧块约 214–488 行:其中 `.section-head`(312)、`.status-badge`(320–343)、`.notice-line`(345)、`.error-box`(353)被 `RecentUpdates.tsx`、`NormalizationReview.tsx`、`Admin/index.tsx`、`ui/StatusBadge.tsx` 共用,**必须保留**;其余 `.stats-*` 规则删除。
   - 新块约 1738–2070 行:整块替换为本计划 Task 7 的新样式。
   - 1550 行附近的 `.stats-grid` 属于 Admin 的 ChangeSummaryPanel,**不要动**。
   - 490 行附近 `/* ---------- 滚动条 ---------- */` 的 `.terminal-scroll` 是共享的,**不要动**。
2. 后端 `/api/session-stats/players` 一次性返回全月列表(无分页参数),翻页是纯前端切片。
3. 构建命令:`cd web/frontend && npm run build`。本机 WSL 环境 rolldown 的 Linux binding 是 `npm install --no-save` 补的,如 build 报 "Cannot find native binding",运行 `npm install --no-save @rolldown/binding-linux-x64-gnu@1.0.1`。
4. 当前仓库工作区有大量与本任务无关的未提交改动,commit 时**只 add 本计划涉及的文件**,不要 `git add .`。

---

### Task 1: Pagination 可复用组件

**Files:**
- Create: `web/frontend/src/components/ui/Pagination.tsx`

- [ ] **Step 1: 创建组件**

```tsx
// web/frontend/src/components/ui/Pagination.tsx
/** 传统页码条:‹ 1 2 3 … N ›。pageCount <= 1 时不渲染。 */
export function Pagination({ page, pageCount, onChange }: {
  page: number;
  pageCount: number;
  onChange: (page: number) => void;
}) {
  if (pageCount <= 1) return null;

  // 始终含首末页,当前页 ±1,间隙折叠为省略号
  const wanted = new Set<number>([1, pageCount, page - 1, page, page + 1]);
  const pages: (number | '…')[] = [];
  let prev = 0;
  for (let i = 1; i <= pageCount; i++) {
    if (!wanted.has(i)) continue;
    if (i - prev > 1) pages.push('…');
    pages.push(i);
    prev = i;
  }

  return (
    <nav className="pager" aria-label="分页">
      <button type="button" disabled={page <= 1} onClick={() => onChange(page - 1)}>‹</button>
      {pages.map((p, i) => p === '…'
        ? <span key={`gap-${i}`} className="pager-gap">…</span>
        : (
          <button
            key={p}
            type="button"
            className={p === page ? 'active' : ''}
            onClick={() => onChange(p)}
          >{p}</button>
        ))}
      <button type="button" disabled={page >= pageCount} onClick={() => onChange(page + 1)}>›</button>
    </nav>
  );
}
```

- [ ] **Step 2: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: 通过(组件尚未被引用,只验证语法/类型)

- [ ] **Step 3: Commit**

```bash
git add web/frontend/src/components/ui/Pagination.tsx
git commit -m "新增页码条组件"
```

---

### Task 2: SessionStats 状态与数据层(月份切换、上月对比、翻页 state、去掉重复函数)

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`

- [ ] **Step 1: 顶部新增月份工具函数**

在 `currentMonth()` 函数(约 25–28 行)之后插入:

```tsx
/** month 形如 "2026-06"。delta 为 ±1 切上/下月。 */
function shiftMonth(month: string, delta: number) {
  const m = /^(\d{4})-(\d{2})$/.exec(month);
  if (!m) return currentMonth();
  const date = new Date(Number(m[1]), Number(m[2]) - 1 + delta, 1);
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
}

function monthLabel(month: string) {
  const m = /^(\d{4})-(\d{2})$/.exec(month);
  return m ? `${m[1]} 年 ${Number(m[2])} 月` : month;
}
```

- [ ] **Step 2: 新增 state 与 import 调整**

组件内 state 区(约 80–103 行),在 `const [error, setError] = useState('');` 之后新增:

```tsx
const [prevOverview, setPrevOverview] = useState<SessionStatsOverview | null>(null);
const [importOpen, setImportOpen] = useState(false);
const [playersPage, setPlayersPage] = useState(1);
const [sessionsPage, setSessionsPage] = useState(1);
```

文件顶部 import 区新增:

```tsx
import { Pagination } from './ui/Pagination';
```

- [ ] **Step 3: loadStats 并行加载上月 overview**

`loadStats`(约 105–139 行)中,在 `const errorsReq = readErrors(month);` 后加一行:

```tsx
const prevOverviewReq = api<SessionStatsOverview>(`/api/session-stats/overview${buildQuery({ month: shiftMonth(month, -1) })}`);
```

`Promise.allSettled` 那行追加该请求:

```tsx
const [overviewRes, playersRes, sessionsRes, errorsRes, prevOverviewRes] = await Promise.allSettled([overviewReq, playersReq, sessionsReq, errorsReq, prevOverviewReq]);
```

在 `if (overviewRes.status === 'fulfilled') ...` 块附近新增(上月请求失败静默,不进 error banner):

```tsx
setPrevOverview(prevOverviewRes.status === 'fulfilled' ? prevOverviewRes.value : null);
```

- [ ] **Step 4: 月份/排序变化时重置页码**

在 `useEffect(() => { loadStats()... }, [loadStats]);`(约 141 行)之后新增:

```tsx
useEffect(() => { setPlayersPage(1); setSessionsPage(1); }, [month]);
useEffect(() => { setPlayersPage(1); }, [sort]);
```

- [ ] **Step 5: 删除重复的 startImport 函数**

`startImport`(约 256–299 行)与 `startImportFiles` 逐行重复 —— 整个删除。后续 JSX 用 `startImportFiles(files)` 代替。

- [ ] **Step 6: 新增翻页与时长条派生值**

在 `importProgress` 派生值(约 301 行)附近新增:

```tsx
const PAGE_SIZE = 20;
const playersPageCount = Math.max(1, Math.ceil(players.items.length / PAGE_SIZE));
const pagedPlayers = players.items.slice((playersPage - 1) * PAGE_SIZE, playersPage * PAGE_SIZE);
const sessionsPageCount = Math.max(1, Math.ceil(sessions.items.length / PAGE_SIZE));
const pagedSessions = sessions.items.slice((sessionsPage - 1) * PAGE_SIZE, sessionsPage * PAGE_SIZE);
// 时长条比例尺:全月最大游戏时长(跨页一致)
const maxGameHours = players.items.reduce((max, p) => Math.max(max, p.game_hours || 0), 0);
```

- [ ] **Step 7: 新增置信度分档与场均时长工具函数**

放在组件外、`formatConfidence` 旁:

```tsx
function confidenceTier(value: number | null | undefined): 'good' | 'mid' | 'low' | 'none' {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 'none';
  const percent = value <= 1 ? value * 100 : value;
  if (percent >= 90) return 'good';
  if (percent >= 70) return 'mid';
  return 'low';
}

function avgHours(total: number, count: number) {
  if (!count) return '—';
  return `${(total / count).toFixed(1)}h`;
}
```

- [ ] **Step 8: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: 可能因 JSX 仍引用已删除的 `startImport` 报错 —— 若报错,先把 JSX 中 `onClick={startImport}` 改为 `onClick={() => startImportFiles(files)}`(约 385 行),再 build 通过。

- [ ] **Step 9: Commit**

```bash
git add web/frontend/src/components/SessionStats.tsx
git commit -m "结团统计:月份切换/上月对比/翻页状态与数据层"
```

---

### Task 3: JSX 重排 — 顶栏 + 折叠导入面板

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`(return JSX,约 364 行起)

- [ ] **Step 1: 替换顶栏与导入区 JSX**

把现有 `<div className="stats-toolbar">…</div>`、`<div className="stats-action-row">…</div>` 以及紧随的「导入结果」`stats-card`(约 366–442 行)整体替换为:

```tsx
<div className="stats-head">
  <h2>结团统计</h2>
  <div className="month-switch">
    <button type="button" onClick={() => setMonth(shiftMonth(month, -1))} aria-label="上一月">‹</button>
    <label>
      <b>{monthLabel(month)}</b>
      <input type="month" value={month} onChange={e => setMonth(e.target.value || currentMonth())} />
    </label>
    <button type="button" onClick={() => setMonth(shiftMonth(month, 1))} aria-label="下一月">›</button>
  </div>
  <span className="stats-head-spacer" />
  <button
    type="button"
    className={importing || importOpen ? 'import-toggle active' : 'import-toggle'}
    onClick={() => setImportOpen(open => !open || importing ? true : false)}
  >
    {importing ? `导入中 ${importProgress}%` : '⬆ 导入战报'}
  </button>
  <button type="button" className="icon-btn" onClick={loadStats} disabled={loading} aria-label="刷新">↻</button>
</div>

{(importOpen || importing) && (
  <section className="stats-card import-panel">
    <div className="section-head">
      <h3>导入战报</h3>
      <label className="stats-sort">
        <span>并发</span>
        <select value={concurrency} onChange={e => setConcurrency(Number(e.target.value) || 1)} disabled={importing}>
          {[1, 2, 3, 4, 5, 6].map(value => <option key={value} value={value}>{value}</option>)}
        </select>
      </label>
      <button type="button" className="danger" onClick={dedupeMonth} disabled={deduping || importing}>{deduping ? '去重中' : '去重'}</button>
      <button type="button" className="icon-btn" onClick={() => !importing && setImportOpen(false)} disabled={importing} aria-label="收起">×</button>
    </div>
    <label
      className="import-drop"
      onDragOver={e => e.preventDefault()}
      onDrop={e => {
        e.preventDefault();
        const dropped = Array.from(e.dataTransfer.files || []).filter(f => f.name.toLowerCase().endsWith('.txt'));
        if (dropped.length) setFiles(dropped);
      }}
    >
      <span className="import-drop-icon">⬆</span>
      <span>{files.length ? selectedFileText : '拖入 TXT 战报,或点击选择文件'}</span>
      <input type="file" accept=".txt,text/plain" multiple onChange={onPickFiles} />
    </label>
    <div className="import-actions">
      <button type="button" className="primary" onClick={() => startImportFiles(files)} disabled={importing || !files.length}>
        {importing ? '导入中' : `开始导入${files.length ? `（${files.length} 个文件）` : ''}`}
      </button>
      {failedImportCount > 0 && (
        <button type="button" onClick={retryFailedFiles} disabled={importing}>重试失败（{failedImportCount}）</button>
      )}
    </div>
    {importJob && (
      <div className="stats-progress">
        <div className="stats-progress-bar"><span style={{ width: `${importProgress}%` }} /></div>
        <div className="stats-progress-meta">
          <span>{importJob.status === 'running' ? '正在解析' : importJob.status === 'completed' ? '导入完成' : '导入失败'}</span>
          <span>{importJob.processed_count}/{importJob.total_count} · 成功 {importJob.success_count} / 失败 {importJob.failure_count} / 跳过 {importJob.skip_count}</span>
          {importJob.current_filename && <span className="mono-ellipsis">{importJob.current_filename}</span>}
        </div>
      </div>
    )}
    {importNotice && <p className="stats-muted">{importNotice}</p>}
    {visibleImportItems.length ? (
      <div className="stats-result-list">
        {visibleImportItems.map((item, i) => {
          const ok = importItemStatus(item);
          return (
            <div className="stats-result" key={`${importItemName(item)}-${i}`}>
              <span className={`status-badge ${ok ? 'ok' : 'bad'}`}>{ok ? '成功' : '失败'}</span>
              <b>{importItemName(item)}</b>
              <small>{importItemMessage(item)}</small>
            </div>
          );
        })}
      </div>
    ) : null}
  </section>
)}
```

说明:
- 原 `stats-action-row`(去重按钮行)删除 —— 去重已收进导入面板头部。
- 原独立的「导入结果」卡片删除 —— 进度/结果/重试合并进导入面板。
- 顶层 `<section className="stats-view terminal-scroll">` 保留不动;其内层包一个新的限宽容器:把 `<section className="stats-view terminal-scroll">` 的直接子内容包进 `<div className="stats-main">…</div>`(限宽 880px 由 CSS 控制)。

- [ ] **Step 2: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add web/frontend/src/components/SessionStats.tsx
git commit -m "结团统计:顶栏重构与折叠导入面板"
```

---

### Task 4: JSX 重排 — KPI 四卡

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`

- [ ] **Step 1: 替换 KPI 区块**

把现有 `<div className="stats-overview">…</div>`(原 395–400 行一带)替换为:

```tsx
<div className="stats-kpis">
  <div className="kpi">
    <span className="kpi-label">本月团数</span>
    <b className="kpi-num">{overview.session_count}</b>
    <span className="kpi-sub">{prevOverview ? `上月 ${prevOverview.session_count}` : '—'}</span>
  </div>
  <div className="kpi">
    <span className="kpi-label">玩家人次</span>
    <b className="kpi-num">{overview.participant_count}</b>
    <span className="kpi-sub">去重后 {players.count} 人</span>
  </div>
  <div className="kpi">
    <span className="kpi-label">总游戏小时</span>
    <b className="kpi-num">{formatHours(overview.total_game_hours)}<small>h</small></b>
    <span className="kpi-sub">场均 {avgHours(overview.total_game_hours, overview.session_count)}</span>
  </div>
  <div className="kpi kpi-gold">
    <span className="kpi-label">主持小时</span>
    <b className="kpi-num">{formatHours(overview.total_host_hours)}<small>h</small></b>
    <span className="kpi-sub">{prevOverview ? `上月 ${formatHours(prevOverview.total_host_hours)}h` : '—'}</span>
  </div>
</div>
```

- [ ] **Step 2: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add web/frontend/src/components/SessionStats.tsx
git commit -m "结团统计:KPI 卡与上月对比"
```

---

### Task 5: JSX 重排 — 玩家统计(排序胶囊 + 时长条 + 翻页 + 空状态)

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`

- [ ] **Step 1: 替换玩家统计 section**

把现有「玩家统计」`stats-card`(原 444–486 行一带)替换为:

```tsx
<section className="stats-card">
  <div className="section-head">
    <h3>玩家统计</h3>
    <span className="stats-count">{players.count} 人</span>
    <div className="sort-chips" role="tablist" aria-label="排序">
      {([['hours', '游戏时长'], ['games', '游戏次数'], ['hosts', '主持次数'], ['name', '玩家名']] as const).map(([value, label]) => (
        <button
          key={value}
          type="button"
          className={sort === value ? 'on' : ''}
          onClick={() => setSort(value)}
        >{label}</button>
      ))}
    </div>
  </div>
  {players.items.length ? (
    <>
      <div className="stats-table-wrap terminal-scroll">
        <table className="stats-table">
          <thead>
            <tr>
              <th className="col-rank">#</th>
              <th>玩家</th>
              <th>QQ</th>
              <th className="col-num">次数</th>
              <th className="col-hours">游戏时长</th>
              <th className="col-num">轮回</th>
              <th className="col-num">主持</th>
            </tr>
          </thead>
          <tbody>
            {pagedPlayers.map((player, i) => {
              const rank = (playersPage - 1) * PAGE_SIZE + i + 1;
              const topRank = sort !== 'name' && rank <= 3;
              return (
                <tr key={player.id}>
                  <td className={topRank ? 'col-rank top' : 'col-rank'}>{rank}</td>
                  <td className="col-name">{player.name || '未命名'}</td>
                  <td className="col-qq">{player.qq || '—'}</td>
                  <td className="col-num">{player.game_count}</td>
                  <td className="col-hours">
                    <span className="hbar">
                      <span className="hbar-track"><i style={{ width: `${maxGameHours ? Math.round((player.game_hours / maxGameHours) * 100) : 0}%` }} /></span>
                      <b>{formatHours(player.game_hours)}h</b>
                    </span>
                  </td>
                  <td className="col-num">{player.reincarnation_count}</td>
                  <td className="col-num">{player.host_count}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <Pagination page={playersPage} pageCount={playersPageCount} onChange={setPlayersPage} />
    </>
  ) : (
    <div className="stats-empty">
      <span className="stats-empty-icon">📊</span>
      <p>本月还没有数据,导入战报后这里会出现统计</p>
    </div>
  )}
</section>
```

注意:原表头有「游戏时长」和「主持时长」两列,新表合并主持时长进「主持」列?**不**——设计稿定的列是:# / 玩家 / QQ / 次数 / 游戏时长(条)/ 轮回 / 主持(次数)。主持时长不再单列(它在 KPI 卡有汇总;按主持次数排序时仍可见次数)。这是设计定稿的简化,不要加回。

- [ ] **Step 2: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add web/frontend/src/components/SessionStats.tsx
git commit -m "结团统计:玩家表排序胶囊/时长条/翻页"
```

---

### Task 6: JSX 重排 — 团列表(行式 + 置信徽章 + 翻页)

**Files:**
- Modify: `web/frontend/src/components/SessionStats.tsx`

- [ ] **Step 1: 替换团列表 section**

把现有「团列表」`stats-card`(原 488–514 行一带)替换为:

```tsx
<section className="stats-card">
  <div className="section-head">
    <h3>团列表</h3>
    <span className="stats-count">{sessions.count} 条</span>
  </div>
  {sessions.items.length ? (
    <>
      <div className="session-rows">
        {pagedSessions.map(session => (
          <article className="session-row" key={session.id}>
            <b className="session-title" title={session.title || undefined}>{session.title || '未命名团'}</b>
            <span className="session-kv">KP <b>{session.kp_name || '未知'}</b></span>
            <span className="session-kv">PL <b>×{session.pl_count}</b></span>
            <span className={`conf-badge ${confidenceTier(session.confidence)}`}>{formatConfidence(session.confidence)}</span>
            <span className="session-dur">{formatHours(session.duration_hours)}h</span>
            <span className="session-ops">
              <button type="button" onClick={() => openSessionDetail(session.id)} disabled={detailLoading}>查看</button>
              <button type="button" className="danger" onClick={() => deleteSession(session.id)}>删除</button>
            </span>
          </article>
        ))}
      </div>
      <Pagination page={sessionsPage} pageCount={sessionsPageCount} onChange={setSessionsPage} />
    </>
  ) : (
    <div className="stats-empty">
      <span className="stats-empty-icon">🎲</span>
      <p>暂无结团记录,导入战报后这里会列出每一团</p>
    </div>
  )}
</section>
```

说明:原行内显示的 `source_filename` 移除(团详情里仍有);置信度从纯文本变 `conf-badge` 徽章。团详情卡片和异常列表的 JSX **保持原样不动**。

- [ ] **Step 2: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: PASS

- [ ] **Step 3: Commit**

```bash
git add web/frontend/src/components/SessionStats.tsx
git commit -m "结团统计:团列表行式布局与置信徽章"
```

---

### Task 7: CSS 重写

**Files:**
- Modify: `web/frontend/src/style.css`

- [ ] **Step 1: 清理旧块(约 214–488 行)**

该块内**保留**这些共享规则(其他组件在用):
- `.section-head`(312–318)
- `.status-badge` / `.status-badge.ok` / `.status-badge.bad`(320–343)
- `.notice-line`(345–351)
- `.error-box`(353–357)
- `.stats-progress` / `.stats-progress-bar` / `.stats-progress-bar span` / `.stats-progress-meta`(366–396,导入面板还在用)
- `.stats-result-list` / `.stats-result` / `.stats-result small`(导入结果列表还在用,合并到下面新块也可)

**删除**该块内其余所有 `.stats-*` 规则(`.stats-view`、`.stats-toolbar`、`.stats-field`、`.stats-sort`、`.stats-file`、`.stats-overview*`、`.stats-card`、`.stats-table*`、`.stats-session*`、`.stats-empty`、`.stats-error-item*`)。把保留的共享规则的注释头改成 `/* ---------- 共享 UI(badge/notice/section-head) ---------- */`。

- [ ] **Step 2: 整块替换新块(约 1738–2070 行)**

把第二个 `/* ---------- 结团统计 ---------- */` 块(从 `.stats-view {` 到 `.stats-error-item b { ... }` 结束,紧邻 `/* ---------- 对比模态 ---------- */` 之前)整体替换为:

```css
/* ---------- 结团统计(仪表盘版) ---------- */
.stats-view {
  height: calc(100vh - 88px);
  overflow: auto;
  padding: 4px 2px 24px;
}
.stats-main {
  max-width: 880px;
  margin: 0 auto;
  display: grid;
  gap: 18px;
  align-content: start;
}

/* 顶栏 */
.stats-head {
  display: flex;
  align-items: center;
  gap: 10px;
}
.stats-head h2 {
  margin: 0;
  font-size: 18px;
  font-weight: 700;
  color: var(--text);
}
.stats-head-spacer { flex: 1; }
.month-switch {
  display: inline-flex;
  align-items: stretch;
  border: 1px solid var(--btn-bd);
  border-radius: var(--radius-m);
  background: var(--panel);
  overflow: hidden;
  margin-left: 6px;
}
.month-switch button {
  border: 0;
  border-radius: 0;
  background: transparent;
  padding: 5px 11px;
  color: var(--muted);
  box-shadow: none;
}
.month-switch button:hover:not(:disabled) { color: var(--accent); background: var(--btn-bg-hover); }
.month-switch label {
  position: relative;
  display: flex;
  align-items: center;
  padding: 5px 13px;
  border-left: 1px solid var(--line-soft);
  border-right: 1px solid var(--line-soft);
  cursor: pointer;
}
.month-switch label b {
  font-size: 13px;
  font-weight: 600;
  color: var(--text);
  font-variant-numeric: tabular-nums;
  white-space: nowrap;
}
.month-switch label input {
  position: absolute;
  inset: 0;
  opacity: 0;
  cursor: pointer;
}
.import-toggle {
  background: var(--accent-strong);
  border-color: var(--accent-strong);
  color: #fff;
  font-weight: 600;
}
.import-toggle:hover:not(:disabled) { background: var(--accent); border-color: var(--accent); }
.import-toggle.active { box-shadow: 0 0 0 2px var(--accent-bg); }
.icon-btn { padding: 7px 11px; }

/* 导入面板 */
.import-panel .section-head { margin-bottom: 0; }
.import-panel .section-head h3 { margin-right: auto; }
.import-drop {
  position: relative;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 10px;
  min-height: 72px;
  padding: 16px;
  border: 1.5px dashed var(--line-strong);
  border-radius: var(--radius-m);
  background: var(--panel-2);
  color: var(--muted);
  cursor: pointer;
  text-align: center;
}
.import-drop:hover { border-color: var(--accent); color: var(--text-soft); background: var(--btn-bg-hover); }
.import-drop-icon { font-size: 18px; opacity: .7; }
.import-drop input {
  position: absolute;
  inset: 0;
  opacity: 0;
  cursor: pointer;
}
.import-actions {
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}
.import-actions .primary {
  background: var(--accent-strong);
  border-color: var(--accent-strong);
  color: #fff;
  font-weight: 600;
}
.import-actions .primary:hover:not(:disabled) { background: var(--accent); border-color: var(--accent); }
.mono-ellipsis {
  font-family: var(--mono);
  font-size: 11px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  max-width: 280px;
}

/* KPI */
.stats-kpis {
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  gap: 10px;
}
.kpi {
  display: grid;
  gap: 2px;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-m);
  background: var(--panel);
  padding: 12px 14px;
  box-shadow: var(--shadow-soft);
}
.kpi-label {
  font-size: 12px;
  color: var(--muted);
}
.kpi-num {
  font-size: 26px;
  font-weight: 700;
  color: var(--text);
  line-height: 1.15;
  letter-spacing: -.02em;
  font-variant-numeric: tabular-nums;
}
.kpi-num small {
  font-size: 13px;
  font-weight: 500;
  color: var(--muted);
  margin-left: 2px;
}
.kpi-gold .kpi-num { color: var(--warm); }
.kpi-sub {
  font-size: 11px;
  color: var(--subtle);
}

/* 卡片与区块头 */
.stats-card {
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-m);
  background: var(--panel);
  padding: 14px 16px;
  display: grid;
  gap: 12px;
  box-shadow: var(--shadow-soft);
}
.stats-card h3 {
  margin: 0;
  font-size: 15px;
  font-weight: 700;
  color: var(--text);
}
.stats-count {
  color: var(--subtle);
  font-size: 12px;
  margin-right: auto;
}
.stats-card .section-head { margin-bottom: 0; justify-content: flex-start; }
.stats-muted {
  margin: 0;
  color: var(--text-soft);
  white-space: pre-wrap;
  font-size: 13px;
}
.stats-sort { display: inline-flex; align-items: center; gap: 6px; }
.stats-sort span { color: var(--muted); font-size: 12px; }

/* 排序胶囊 */
.sort-chips {
  display: inline-flex;
  gap: 3px;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-m);
  background: var(--panel-2);
  padding: 3px;
}
.sort-chips button {
  border: 0;
  border-radius: var(--radius-s);
  background: transparent;
  box-shadow: none;
  padding: 3px 11px;
  font-size: 12px;
  color: var(--muted);
}
.sort-chips button:hover:not(:disabled) { color: var(--text-soft); background: var(--btn-bg-hover); }
.sort-chips button.on {
  background: var(--btn-bg-active);
  color: var(--accent);
  font-weight: 600;
}

/* 玩家表 */
.stats-table-wrap {
  overflow: auto;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-s);
}
.stats-table {
  width: 100%;
  border-collapse: collapse;
  min-width: 640px;
}
.stats-table th, .stats-table td {
  padding: 9px 10px;
  border-bottom: 1px solid var(--line-soft);
  text-align: left;
  vertical-align: middle;
}
.stats-table th {
  color: var(--subtle);
  background: var(--panel-2);
  font-size: 11.5px;
  font-weight: 600;
  white-space: nowrap;
}
.stats-table td { color: var(--text-soft); font-size: 13px; }
.stats-table tbody tr:hover td { background: var(--btn-bg-hover); }
.stats-table tr:last-child td { border-bottom: 0; }
.stats-table .col-rank { width: 36px; color: var(--subtle); font-size: 12px; font-weight: 600; font-variant-numeric: tabular-nums; }
.stats-table .col-rank.top { color: var(--warm); }
.stats-table .col-name { color: var(--text); font-weight: 600; overflow-wrap: anywhere; }
.stats-table .col-qq { font-family: var(--mono); font-size: 11.5px; color: var(--subtle); overflow-wrap: anywhere; }
.stats-table th.col-num, .stats-table td.col-num { text-align: right; font-variant-numeric: tabular-nums; width: 56px; }
.stats-table .col-hours { min-width: 170px; }
.hbar {
  display: flex;
  align-items: center;
  gap: 8px;
}
.hbar-track {
  flex: 1;
  height: 5px;
  border-radius: 3px;
  background: var(--line-soft);
  overflow: hidden;
}
.hbar-track i {
  display: block;
  height: 100%;
  border-radius: inherit;
  background: var(--accent-strong);
}
.hbar b {
  min-width: 44px;
  text-align: right;
  font-size: 12.5px;
  font-weight: 600;
  color: var(--accent);
  font-variant-numeric: tabular-nums;
}

/* 页码条 */
.pager {
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 4px;
  flex-wrap: wrap;
}
.pager button {
  min-width: 30px;
  padding: 4px 9px;
  font-size: 12.5px;
  font-variant-numeric: tabular-nums;
}
.pager button.active {
  background: var(--btn-bg-active);
  border-color: var(--accent-strong);
  color: var(--accent);
  font-weight: 600;
}
.pager-gap { color: var(--subtle); padding: 0 2px; }

/* 团列表 */
.session-rows {
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-s);
  overflow: hidden;
}
.session-row {
  display: flex;
  align-items: center;
  gap: 12px;
  padding: 10px 14px;
  border-bottom: 1px solid var(--line-soft);
  background: var(--panel);
}
.session-row:last-child { border-bottom: 0; }
.session-row:hover { background: var(--btn-bg-hover); }
.session-title {
  flex: 1;
  min-width: 0;
  color: var(--text);
  font-weight: 600;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.session-kv {
  color: var(--muted);
  font-size: 12px;
  white-space: nowrap;
}
.session-kv b { color: var(--text-soft); font-weight: 600; }
.conf-badge {
  font-size: 11px;
  padding: 1px 8px;
  border-radius: 999px;
  white-space: nowrap;
}
.conf-badge.good { background: var(--green-bg); color: var(--green); }
.conf-badge.mid { background: var(--warm-bg); color: var(--warm); }
.conf-badge.low { background: var(--red-bg); color: var(--red); }
.conf-badge.none { background: var(--panel-2); color: var(--subtle); border: 1px solid var(--line-soft); }
.session-dur {
  font-size: 13px;
  font-weight: 600;
  color: var(--accent);
  white-space: nowrap;
  font-variant-numeric: tabular-nums;
  min-width: 44px;
  text-align: right;
}
.session-ops { display: flex; gap: 6px; }
.session-ops button { padding: 3px 10px; font-size: 12px; }

/* 空状态 */
.stats-empty {
  display: grid;
  justify-items: center;
  gap: 6px;
  padding: 28px 16px;
  border: 1px dashed var(--line);
  border-radius: var(--radius-s);
  background: var(--panel-2);
  color: var(--muted);
  text-align: center;
}
.stats-empty p { margin: 0; font-size: 13px; }
.stats-empty-icon { font-size: 22px; opacity: .6; }

/* 团详情(沿用原结构,微调) */
.stats-detail-panel { scroll-margin-top: 12px; }
.stats-field { display: grid; gap: 4px; }
.stats-field span { color: var(--muted); font-size: 11px; }
.stats-detail-grid {
  display: grid;
  grid-template-columns: minmax(260px, .8fr) minmax(320px, 1fr);
  gap: 14px;
  align-items: start;
}
.stats-detail-edit {
  display: grid;
  grid-template-columns: minmax(180px, 1fr) 140px;
  gap: 10px;
  align-items: end;
}
.stats-detail-actions {
  grid-column: 1 / -1;
  display: flex;
  gap: 8px;
  flex-wrap: wrap;
}
.stats-detail-meta {
  display: grid;
  grid-template-columns: repeat(2, minmax(0, 1fr));
  gap: 8px;
}
.stats-detail-meta span,
.stats-participant-list span {
  display: grid;
  gap: 2px;
  min-width: 0;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-s);
  background: var(--panel-2);
  padding: 8px 10px;
  color: var(--text-soft);
  overflow-wrap: anywhere;
}
.stats-detail-meta b,
.stats-participant-list b {
  color: var(--muted);
  font-size: 11px;
  font-weight: 600;
}
.stats-detail-participants,
.stats-detail-raw { display: grid; gap: 8px; }
.stats-detail-participants h4,
.stats-detail-raw h4 {
  margin: 2px 0 0;
  color: var(--text);
  font-size: 13px;
}
.stats-participant-list {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
  gap: 8px;
}
.stats-participant-list small { color: var(--muted); font-size: 12px; }
.stats-detail-raw pre {
  max-height: 260px;
  margin: 0;
  overflow: auto;
  white-space: pre-wrap;
  word-break: break-word;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-s);
  background: var(--panel-2);
  color: var(--text-soft);
  padding: 10px 12px;
  font-family: var(--mono);
  font-size: 12px;
  line-height: 1.6;
}

/* 异常列表 */
.stats-error-list { display: grid; gap: 8px; }
.stats-error-item {
  display: flex;
  flex-wrap: wrap;
  gap: 8px;
  border: 1px solid var(--line-soft);
  border-radius: var(--radius-s);
  background: var(--panel-2);
  padding: 10px;
}
.stats-error-item span {
  display: inline-flex;
  gap: 6px;
  max-width: 100%;
  border: 1px solid var(--line);
  border-radius: var(--radius-xs);
  background: var(--panel-3);
  padding: 4px 8px;
  color: var(--text-soft);
  font-size: 12px;
}
.stats-error-item b {
  color: var(--muted);
  font-family: var(--mono);
  font-weight: 500;
}

/* 窄屏适配 */
@media (max-width: 720px) {
  .stats-kpis { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .stats-head { flex-wrap: wrap; }
  .session-row { flex-wrap: wrap; }
  .session-title { flex-basis: 100%; }
  .stats-detail-grid { grid-template-columns: 1fr; }
}
```

注意:`.stats-file`、`.stats-toolbar`、`.stats-overview`、`.stats-action-row`、`.stats-session`(旧网格版)、`.stats-session-main/meta/actions`、`.stats-field-small`、`.stats-session-list` 这些类名在新 JSX 中已不存在,**不要保留**。

- [ ] **Step 3: 确认无残留引用**

Run: `cd web/frontend && grep -n "stats-toolbar\|stats-overview\|stats-file\|stats-action-row\|stats-session-main\|stats-session-meta\|stats-session-actions\|stats-session-list\|stats-field-small" src/ -r`
Expected: 无输出(JSX 与 CSS 中都已清除)

- [ ] **Step 4: 构建验证**

Run: `cd web/frontend && npm run build`
Expected: PASS

- [ ] **Step 5: Commit**

```bash
git add web/frontend/src/style.css web/frontend/src/components/SessionStats.tsx
git commit -m "结团统计:仪表盘版样式重写"
```

---

### Task 8: 手动验证(开发服务器)

**前置:** 后端(`uvicorn app.main:app --port 8000`,环境变量 `ADMIN_PASSWORD` 任意值)与前端(`npm run dev`)都在跑。

- [ ] **Step 1: 空状态验证**

打开 `http://localhost:5173/?tab=stats`(当前月无数据):
- KPI 四卡显示 0,标签/次级行排版正确
- 玩家统计和团列表显示居中空状态(图标 + 引导文案)
- 不出现页码条
- 导入面板默认折叠,只有顶部「⬆ 导入战报」按钮

- [ ] **Step 2: 交互验证**

- 点月份切换 `‹` `›`:月份变化、数据重新加载、KPI 上月对比变化
- 点月份文字:弹出原生 month picker
- 点「⬆ 导入战报」:面板展开,含拖拽区/并发/开始导入/去重;点 × 收起
- 排序胶囊四个选项可切换(无数据时不报错)

- [ ] **Step 3: 翻页验证(若有数据月份)**

切到有历史数据的月份(如有):玩家表/团列表 >20 条时出现页码条,切页排名连续;≤20 条不出现。无历史数据则跳过,标注「翻页未实测」。

- [ ] **Step 4: 暗/亮主题截图核对**

右上角主题切换,两套主题下分别检查 KPI 卡、表格、徽章、空状态的对比度与边线。

- [ ] **Step 5: 回归检查其他页面**

- `/?tab=recent`(最近更新)与后台页:`.section-head`、`.status-badge`、`.notice-line` 样式未损坏
- 后台的 ChangeSummaryPanel(`.stats-grid`)正常

- [ ] **Step 6: 最终提交(如有调整)**

```bash
git add web/frontend/src
git commit -m "结团统计:手动验证后的微调"
```

---

## Self-Review 记录

- **Spec coverage:** 顶栏/折叠导入/KPI(含上月对比)/排序胶囊/时长条/翻页/置信徽章/空状态/团详情留存/异常列表留存/暗亮主题 → Task 2–7 全覆盖;验证 → Task 8。
- **类型一致性:** `SessionStatsOverview`、`SessionStatsPlayerSort` 等均为 `types.ts` 现有类型,无新增;`Pagination` props 在 Task 1 定义、Task 5/6 使用一致。
- **占位符扫描:** 无 TBD/TODO;所有代码块完整。
