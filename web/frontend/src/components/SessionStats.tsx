import { ChangeEvent, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { api, buildQuery, routePath } from '../api';
import { Pagination } from './ui/Pagination';
import {
  SessionStatsDedupeResponse,
  SessionStatsErrorItem,
  SessionStatsErrorsResponse,
  SessionStatsImportFileResult,
  SessionStatsImportJob,
  SessionStatsImportResponse,
  SessionStatsOverview,
  SessionStatsPlayerSort,
  SessionStatsSessionDetail,
  SessionStatsSessionPatchResponse,
  SessionStatsPlayersResponse,
  SessionStatsSessionsResponse,
} from '../types';

const emptyOverview: SessionStatsOverview = {
  session_count: 0,
  participant_count: 0,
  total_game_hours: 0,
  total_host_hours: 0,
};

function currentMonth() {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}`;
}

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

function formatHours(value: number | null | undefined) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '0';
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function formatConfidence(value: number | null | undefined) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '未标注';
  const percent = value <= 1 ? value * 100 : value;
  return `${percent.toFixed(0)}%`;
}

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

function importItemName(item: SessionStatsImportFileResult) {
  return item.filename || item.source_filename || item.title || '未命名文件';
}

function importItemStatus(item: SessionStatsImportFileResult) {
  if (item.skipped) return true;
  if (typeof item.success === 'boolean') return item.success;
  if (typeof item.ok === 'boolean') return item.ok;
  return !item.error;
}

function importItemMessage(item: SessionStatsImportFileResult) {
  return item.reason || item.error || item.detail || item.message || (item.session_id ? `记录 ID：${item.session_id}` : '已处理');
}

function stringifyValue(value: unknown) {
  if (value === null || value === undefined || value === '') return '未记录';
  if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean') return String(value);
  return JSON.stringify(value);
}

function rawPayloadSummary(value: unknown) {
  if (value === null || value === undefined || value === '') return '未记录';
  const text = typeof value === 'string' ? value : JSON.stringify(value, null, 2);
  if (!text) return '未记录';
  return text.length > 1200 ? `${text.slice(0, 1200)}\n...` : text;
}

async function readErrors(month: string): Promise<{ data: SessionStatsErrorsResponse | null; gated: boolean; error: string }> {
  const res = await fetch(`/api/session-stats/errors${buildQuery({ month })}`, { credentials: 'include' });
  if (res.status === 401 || res.status === 503) return { data: null, gated: true, error: '' };
  if (!res.ok) {
    const text = await res.text();
    return { data: null, gated: false, error: text || `HTTP ${res.status}` };
  }
  return { data: await res.json() as SessionStatsErrorsResponse, gated: false, error: '' };
}

export function SessionStats() {
  const [month, setMonth] = useState(currentMonth);
  const [concurrency, setConcurrency] = useState(2);
  const [sort, setSort] = useState<SessionStatsPlayerSort>('hours');
  const [overview, setOverview] = useState<SessionStatsOverview>(emptyOverview);
  const [players, setPlayers] = useState<SessionStatsPlayersResponse>({ items: [], count: 0, month: currentMonth() });
  const [sessions, setSessions] = useState<SessionStatsSessionsResponse>({ items: [], count: 0, month: currentMonth() });
  const [errors, setErrors] = useState<SessionStatsErrorItem[]>([]);
  const [errorsGated, setErrorsGated] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [prevOverview, setPrevOverview] = useState<SessionStatsOverview | null>(null);
  const [importOpen, setImportOpen] = useState(false);
  const [playersPage, setPlayersPage] = useState(1);
  const [sessionsPage, setSessionsPage] = useState(1);
  const [files, setFiles] = useState<File[]>([]);
  const [importing, setImporting] = useState(false);
  const [deduping, setDeduping] = useState(false);
  const [importResult, setImportResult] = useState<SessionStatsImportResponse | null>(null);
  const [importJob, setImportJob] = useState<SessionStatsImportJob | null>(null);
  const [importNotice, setImportNotice] = useState('');
  const [sessionDetail, setSessionDetail] = useState<SessionStatsSessionDetail | null>(null);
  const [detailLoading, setDetailLoading] = useState(false);
  const [detailSaving, setDetailSaving] = useState(false);
  const [detailError, setDetailError] = useState('');
  const [editTitle, setEditTitle] = useState('');
  const [editDuration, setEditDuration] = useState('');
  const reqIdRef = useRef(0);
  const pollRef = useRef<number | undefined>(undefined);

  const loadStats = useCallback(async () => {
    const reqId = ++reqIdRef.current;
    setLoading(true);
    setError('');
    const overviewReq = api<SessionStatsOverview>(`/api/session-stats/overview${buildQuery({ month })}`);
    const playersReq = api<SessionStatsPlayersResponse>(`/api/session-stats/players${buildQuery({ month, sort })}`);
    const sessionsReq = api<SessionStatsSessionsResponse>(`/api/session-stats/sessions${buildQuery({ month })}`);
    const errorsReq = readErrors(month);
    const prevOverviewReq = api<SessionStatsOverview>(`/api/session-stats/overview${buildQuery({ month: shiftMonth(month, -1) })}`);

    const [overviewRes, playersRes, sessionsRes, errorsRes, prevOverviewRes] = await Promise.allSettled([overviewReq, playersReq, sessionsReq, errorsReq, prevOverviewReq]);
    if (reqId !== reqIdRef.current) return;

    const messages: string[] = [];
    if (overviewRes.status === 'fulfilled') setOverview(overviewRes.value);
    else messages.push(`概览读取失败：${overviewRes.reason instanceof Error ? overviewRes.reason.message : String(overviewRes.reason)}`);

    setPrevOverview(prevOverviewRes.status === 'fulfilled' ? prevOverviewRes.value : null);

    if (playersRes.status === 'fulfilled') setPlayers(playersRes.value);
    else messages.push(`玩家表读取失败：${playersRes.reason instanceof Error ? playersRes.reason.message : String(playersRes.reason)}`);

    if (sessionsRes.status === 'fulfilled') setSessions(sessionsRes.value);
    else messages.push(`团列表读取失败：${sessionsRes.reason instanceof Error ? sessionsRes.reason.message : String(sessionsRes.reason)}`);

    if (errorsRes.status === 'fulfilled') {
      setErrors(errorsRes.value.data?.items || []);
      setErrorsGated(errorsRes.value.gated);
      if (errorsRes.value.error) messages.push(`异常列表读取失败：${errorsRes.value.error}`);
    } else {
      setErrors([]);
      setErrorsGated(false);
      messages.push(`异常列表读取失败：${errorsRes.reason instanceof Error ? errorsRes.reason.message : String(errorsRes.reason)}`);
    }

    setError(messages.join('\n'));
    setLoading(false);
  }, [month, sort]);

  useEffect(() => { loadStats().catch(e => setError(e instanceof Error ? e.message : String(e))); }, [loadStats]);
  useEffect(() => { setPlayersPage(1); setSessionsPage(1); }, [month]);
  useEffect(() => { setPlayersPage(1); }, [sort]);
  useEffect(() => () => { if (pollRef.current) clearTimeout(pollRef.current); }, []);

  const selectedFileText = useMemo(() => {
    if (!files.length) return '未选择文件';
    if (files.length === 1) return files[0].name;
    return `${files.length} 个文件`;
  }, [files]);

  const onPickFiles = (event: ChangeEvent<HTMLInputElement>) => {
    setFiles(Array.from(event.target.files || []));
  };

  const pollImportJob = async (jobId: string) => {
    try {
      const job = await api<SessionStatsImportJob>(`/api/session-stats/import-jobs/${routePath(jobId)}`);
      setImportJob(job);
      setImportResult({
        success_count: job.success_count,
        failure_count: job.failure_count,
        skip_count: job.skip_count,
        items: job.items || [],
      });
      if (job.status === 'running') {
        pollRef.current = window.setTimeout(() => pollImportJob(jobId), 1000);
        return;
      }
      setImporting(false);
      if (job.status === 'failed') setImportNotice(job.error || '导入任务失败');
      await loadStats();
    } catch (e: unknown) {
      setImportNotice(e instanceof Error ? e.message : String(e));
      setImporting(false);
    }
  };

  const startImportFiles = async (targetFiles: File[], mode: 'retry' | 'all' = 'all') => {
    if (!targetFiles.length) {
      alert('请先选择 txt 文件。');
      return;
    }
    const fd = new FormData();
    fd.append('month', month);
    fd.append('concurrency', String(concurrency));
    targetFiles.forEach(file => fd.append('files', file));
    setImporting(true);
    setImportNotice(mode === 'retry' ? `正在重试 ${targetFiles.length} 个失败文件。` : '');
    setImportResult(null);
    setImportJob(null);
    if (pollRef.current) clearTimeout(pollRef.current);
    try {
      const res = await fetch('/api/session-stats/import-jobs', { method: 'POST', credentials: 'include', body: fd });
      const text = await res.text();
      let payload: unknown = null;
      if (text) {
        try {
          payload = JSON.parse(text);
        } catch {
          setImportNotice(text);
        }
      }
      if (!res.ok) {
        const detail = payload && typeof payload === 'object' && 'detail' in payload ? String((payload as { detail?: unknown }).detail) : '';
        setImportNotice(prev => prev || detail || `导入请求失败：HTTP ${res.status}`);
        setImporting(false);
        return;
      }
      const job = payload as SessionStatsImportJob | null;
      if (!job?.job_id) {
        setImportNotice('导入任务创建失败：后端没有返回 job_id');
        setImporting(false);
        return;
      }
      setImportJob(job);
      pollImportJob(job.job_id);
    } catch (e: unknown) {
      setImportNotice(e instanceof Error ? e.message : String(e));
      setImporting(false);
    }
  };

  const retryFailedFiles = async () => {
    const failedNames = new Set((importResult?.items || []).filter(item => !importItemStatus(item)).map(importItemName));
    if (!failedNames.size) {
      alert('当前没有可重试的失败文件。');
      return;
    }
    const selectedByName = new Map(files.map(file => [file.name, file]));
    const retryFiles = Array.from(failedNames).map(name => selectedByName.get(name)).filter((file): file is File => !!file);
    const missingCount = failedNames.size - retryFiles.length;
    if (!retryFiles.length) {
      setImportNotice('找不到失败文件的本地选择记录，请重新选择原 TXT 后再重试。');
      return;
    }
    if (missingCount) {
      setImportNotice(`有 ${missingCount} 个失败文件不在当前选择中，只重试找到的 ${retryFiles.length} 个文件。`);
    }
    await startImportFiles(retryFiles, 'retry');
  };

  const dedupeMonth = async () => {
    if (!confirm(`确认清理 ${month} 中内容完全相同的重复团记录？每组只保留最早导入的一条。`)) return;
    setDeduping(true);
    setImportNotice('');
    try {
      const result = await api<SessionStatsDedupeResponse>(`/api/session-stats/dedupe${buildQuery({ month })}`, { method: 'POST' });
      setImportNotice(result.deleted_count ? `已删除 ${result.deleted_count} 条重复团记录。` : '没有发现内容完全相同的重复团记录。');
      await loadStats();
    } catch (e: unknown) {
      setImportNotice(e instanceof Error ? e.message : String(e));
    } finally {
      setDeduping(false);
    }
  };

  const importProgress = importJob && importJob.total_count > 0
    ? Math.min(100, Math.round((importJob.processed_count / importJob.total_count) * 100))
    : 0;
  const visibleImportItems = (importResult?.items || []).filter(item => !item.skipped);
  const failedImportCount = (importResult?.items || []).filter(item => !importItemStatus(item)).length;

  const PAGE_SIZE = 20;
  const playersPageCount = Math.max(1, Math.ceil(players.items.length / PAGE_SIZE));
  const pagedPlayers = players.items.slice((playersPage - 1) * PAGE_SIZE, playersPage * PAGE_SIZE);
  const sessionsPageCount = Math.max(1, Math.ceil(sessions.items.length / PAGE_SIZE));
  const pagedSessions = sessions.items.slice((sessionsPage - 1) * PAGE_SIZE, sessionsPage * PAGE_SIZE);
  // 时长条比例尺:全月最大游戏时长(跨页一致)
  const maxGameHours = players.items.reduce((max, p) => Math.max(max, p.game_hours || 0), 0);

  const deleteSession = async (id: string | number) => {
    if (!confirm('确认删除这条结团记录？')) return;
    try {
      await api<{ ok?: boolean }>(`/api/session-stats/sessions/${routePath(String(id))}`, { method: 'DELETE' });
      if (sessionDetail && String(sessionDetail.id) === String(id)) setSessionDetail(null);
      await loadStats();
    } catch (e: unknown) {
      alert(e instanceof Error ? e.message : String(e));
    }
  };

  const openSessionDetail = async (id: string | number) => {
    setDetailLoading(true);
    setDetailError('');
    setSessionDetail(null);
    try {
      const detail = await api<SessionStatsSessionDetail>(`/api/session-stats/sessions/${routePath(String(id))}`);
      setSessionDetail(detail);
      setEditTitle(detail.title || '');
      setEditDuration(typeof detail.duration_hours === 'number' ? String(detail.duration_hours) : '');
    } catch (e: unknown) {
      setDetailError(e instanceof Error ? e.message : String(e));
    } finally {
      setDetailLoading(false);
    }
  };

  const closeSessionDetail = () => {
    setSessionDetail(null);
    setDetailError('');
  };

  const saveSessionDetail = async () => {
    if (!sessionDetail) return;
    const duration = Number(editDuration);
    if (!Number.isFinite(duration) || duration < 0) {
      setDetailError('时长必须是大于等于 0 的数字。');
      return;
    }
    setDetailSaving(true);
    setDetailError('');
    try {
      const result = await api<SessionStatsSessionPatchResponse>(`/api/session-stats/sessions/${routePath(String(sessionDetail.id))}`, {
        method: 'PATCH',
        body: JSON.stringify({ title: editTitle.trim(), duration_hours: duration }),
      });
      setSessionDetail(result.session);
      setEditTitle(result.session.title || '');
      setEditDuration(typeof result.session.duration_hours === 'number' ? String(result.session.duration_hours) : '');
      await loadStats();
    } catch (e: unknown) {
      setDetailError(e instanceof Error ? e.message : String(e));
    } finally {
      setDetailSaving(false);
    }
  };

  return (
    <section className="stats-view terminal-scroll">
      <div className="stats-main">
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
          onClick={() => setImportOpen(open => importing ? true : !open)}
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

      {error && <div className="notice-line error-box">{error}</div>}

      <div className="stats-overview">
        <span><b>{overview.session_count}</b>本月团数</span>
        <span><b>{overview.participant_count}</b>玩家人次</span>
        <span><b>{formatHours(overview.total_game_hours)}</b>总游戏小时</span>
        <span><b>{formatHours(overview.total_host_hours)}</b>主持小时</span>
      </div>

      {(importJob || importResult || importNotice) && (
        <section className="stats-card">
          <div className="section-head">
            <h3>导入结果</h3>
            {importJob ? (
              <span className="status-badge">
                {importJob.processed_count}/{importJob.total_count} 已处理 · {importJob.success_count} 成功 / {importJob.failure_count} 失败
              </span>
            ) : importResult ? (
              <span className="status-badge">{importResult.success_count} 成功 / {importResult.failure_count} 失败</span>
            ) : null}
            {(importJob || importResult) && <span className="status-badge">{(importJob?.skip_count ?? importResult?.skip_count) || 0} SKIP</span>}
            <button type="button" onClick={retryFailedFiles} disabled={importing || !failedImportCount}>重试失败</button>
          </div>
          {importJob && (
            <div className="stats-progress">
              <div className="stats-progress-bar"><span style={{ width: `${importProgress}%` }} /></div>
              <div className="stats-progress-meta">
                <span>{importJob.status === 'running' ? '正在解析' : importJob.status === 'completed' ? '导入完成' : '导入失败'}</span>
                <span>{importProgress}%</span>
                {importJob.current_filename && <span>{importJob.current_filename}</span>}
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

      <section className="stats-card">
        <div className="section-head">
          <h3>玩家统计</h3>
          <label className="stats-sort">
            <span>排序</span>
            <select value={sort} onChange={e => setSort(e.target.value as SessionStatsPlayerSort)}>
              <option value="hours">游戏时长</option>
              <option value="games">游戏次数</option>
              <option value="hosts">主持次数</option>
              <option value="name">玩家名</option>
            </select>
          </label>
        </div>
        <div className="stats-table-wrap terminal-scroll">
          <table className="stats-table">
            <thead>
              <tr>
                <th>玩家名</th>
                <th>QQ</th>
                <th>游戏次数</th>
                <th>游戏时长</th>
                <th>轮回次数</th>
                <th>主持次数</th>
                <th>主持时长</th>
              </tr>
            </thead>
            <tbody>
              {players.items.map(player => (
                <tr key={player.id}>
                  <td data-label="玩家名">{player.name || '未命名'}</td>
                  <td data-label="QQ">{player.qq || '未记录'}</td>
                  <td data-label="游戏次数">{player.game_count}</td>
                  <td data-label="游戏时长">{formatHours(player.game_hours)}</td>
                  <td data-label="轮回次数">{player.reincarnation_count}</td>
                  <td data-label="主持次数">{player.host_count}</td>
                  <td data-label="主持时长">{formatHours(player.host_hours)}</td>
                </tr>
              ))}
              {!players.items.length && <tr><td colSpan={7}>暂无玩家统计。</td></tr>}
            </tbody>
          </table>
        </div>
      </section>

      <section className="stats-card">
        <div className="section-head">
          <h3>团列表</h3>
          <span className="status-badge">{sessions.count} 条</span>
        </div>
        <div className="stats-session-list">
          {sessions.items.map(session => (
            <article className="stats-session" key={session.id}>
              <div className="stats-session-main">
                <b>{session.title || '未命名团'}</b>
                <small>{session.source_filename || '未知来源'}</small>
              </div>
              <div className="stats-session-meta">
                <span><b>时长</b>{formatHours(session.duration_hours)} 小时</span>
                <span><b>KP</b>{session.kp_name || '未知 KP'}{session.kp_qq ? ` (${session.kp_qq})` : ''}</span>
                <span><b>PL</b>{session.pl_count}</span>
                <span><b>置信度</b>{formatConfidence(session.confidence)}</span>
              </div>
              <div className="stats-session-actions">
                <button type="button" onClick={() => openSessionDetail(session.id)} disabled={detailLoading}>查看</button>
                <button type="button" className="danger" onClick={() => deleteSession(session.id)}>删除</button>
              </div>
            </article>
          ))}
          {!sessions.items.length && <div className="stats-empty">暂无结团记录。</div>}
        </div>
      </section>

      {(sessionDetail || detailLoading || detailError) && (
        <section className="stats-card stats-detail-panel">
          <div className="section-head">
            <h3>团详情 / 编辑基础信息</h3>
            {sessionDetail && <span className="status-badge">ID {sessionDetail.id}</span>}
          </div>
          {detailLoading ? (
            <div className="stats-empty">正在读取团详情。</div>
          ) : sessionDetail ? (
            <>
              {detailError && <div className="notice-line error-box">{detailError}</div>}
              <div className="stats-detail-grid">
                <div className="stats-detail-edit">
                  <label className="stats-field">
                    <span>标题</span>
                    <input value={editTitle} onChange={e => setEditTitle(e.target.value)} />
                  </label>
                  <label className="stats-field">
                    <span>时长（小时）</span>
                    <input type="number" min="0" step="0.5" value={editDuration} onChange={e => setEditDuration(e.target.value)} />
                  </label>
                  <div className="stats-detail-actions">
                    <button type="button" onClick={saveSessionDetail} disabled={detailSaving}>{detailSaving ? '保存中' : '保存'}</button>
                    <button type="button" onClick={closeSessionDetail} disabled={detailSaving}>关闭</button>
                  </div>
                </div>
                <div className="stats-detail-meta">
                  <span><b>文件名</b>{sessionDetail.source_filename || '未知来源'}</span>
                  <span><b>月份</b>{sessionDetail.month || '未记录'}</span>
                  <span><b>KP</b>{sessionDetail.kp?.name || '未知 KP'}{sessionDetail.kp?.qq ? ` (${sessionDetail.kp.qq})` : ''}</span>
                  <span><b>模型</b>{sessionDetail.model_name || '未记录'}</span>
                  <span><b>置信度</b>{formatConfidence(sessionDetail.confidence)}</span>
                  <span><b>创建时间</b>{sessionDetail.created_at || '未记录'}</span>
                </div>
              </div>
              <div className="stats-detail-participants">
                <h4>PL / 参与者</h4>
                {sessionDetail.participants.length ? (
                  <div className="stats-participant-list">
                    {sessionDetail.participants.map(participant => (
                      <span key={participant.id}>
                        <b>{participant.name || '未命名'}</b>
                        <small>
                          {participant.qq || '未记录 QQ'}
                          {participant.role ? ` · ${participant.role}` : ''}
                          {participant.is_host ? ' · KP' : ''}
                          {` · ${formatHours(participant.duration_hours)} 小时`}
                          {` · 轮回 ${participant.reincarnation_count ?? 0}`}
                        </small>
                      </span>
                    ))}
                  </div>
                ) : (
                  <div className="stats-empty">暂无参与者记录。</div>
                )}
              </div>
              <div className="stats-detail-raw">
                <h4>raw_payload 简要</h4>
                <pre>{rawPayloadSummary(sessionDetail.raw_payload)}</pre>
              </div>
            </>
          ) : (
            <div className="notice-line error-box">{detailError}</div>
          )}
        </section>
      )}

      <section className="stats-card">
        <div className="section-head">
          <h3>异常列表</h3>
          <span className="status-badge">{errors.length} 条</span>
        </div>
        {errorsGated ? (
          <div className="stats-empty">需要后台登录后查看异常列表</div>
        ) : errors.length ? (
          <div className="stats-error-list">
            {errors.map((item, index) => (
              <div className="stats-error-item" key={index}>
                {Object.entries(item).map(([key, value]) => (
                  <span key={key}><b>{key}</b>{stringifyValue(value)}</span>
                ))}
              </div>
            ))}
          </div>
        ) : (
          <div className="stats-empty">暂无异常。</div>
        )}
      </section>
      </div>
    </section>
  );
}
