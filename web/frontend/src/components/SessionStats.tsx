import { ChangeEvent, useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { api, buildQuery, routePath } from '../api';
import {
  SessionStatsErrorItem,
  SessionStatsErrorsResponse,
  SessionStatsImportFileResult,
  SessionStatsImportJob,
  SessionStatsImportResponse,
  SessionStatsOverview,
  SessionStatsPlayerSort,
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

function formatHours(value: number | null | undefined) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '0';
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function formatConfidence(value: number | null | undefined) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return '未标注';
  const percent = value <= 1 ? value * 100 : value;
  return `${percent.toFixed(0)}%`;
}

function importItemName(item: SessionStatsImportFileResult) {
  return item.filename || item.source_filename || item.title || '未命名文件';
}

function importItemStatus(item: SessionStatsImportFileResult) {
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
  const [sort, setSort] = useState<SessionStatsPlayerSort>('hours');
  const [overview, setOverview] = useState<SessionStatsOverview>(emptyOverview);
  const [players, setPlayers] = useState<SessionStatsPlayersResponse>({ items: [], count: 0, month: currentMonth() });
  const [sessions, setSessions] = useState<SessionStatsSessionsResponse>({ items: [], count: 0, month: currentMonth() });
  const [errors, setErrors] = useState<SessionStatsErrorItem[]>([]);
  const [errorsGated, setErrorsGated] = useState(false);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [files, setFiles] = useState<File[]>([]);
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState<SessionStatsImportResponse | null>(null);
  const [importJob, setImportJob] = useState<SessionStatsImportJob | null>(null);
  const [importNotice, setImportNotice] = useState('');
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

    const [overviewRes, playersRes, sessionsRes, errorsRes] = await Promise.allSettled([overviewReq, playersReq, sessionsReq, errorsReq]);
    if (reqId !== reqIdRef.current) return;

    const messages: string[] = [];
    if (overviewRes.status === 'fulfilled') setOverview(overviewRes.value);
    else messages.push(`概览读取失败：${overviewRes.reason instanceof Error ? overviewRes.reason.message : String(overviewRes.reason)}`);

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

  const startImport = async () => {
    if (!files.length) {
      alert('请先选择 txt 文件。');
      return;
    }
    const fd = new FormData();
    fd.append('month', month);
    files.forEach(file => fd.append('files', file));
    setImporting(true);
    setImportNotice('');
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

  const importProgress = importJob && importJob.total_count > 0
    ? Math.min(100, Math.round((importJob.processed_count / importJob.total_count) * 100))
    : 0;

  const deleteSession = async (id: string | number) => {
    if (!confirm('确认删除这条结团记录？')) return;
    try {
      await api<{ ok?: boolean }>(`/api/session-stats/sessions/${routePath(String(id))}`, { method: 'DELETE' });
      await loadStats();
    } catch (e: unknown) {
      alert(e instanceof Error ? e.message : String(e));
    }
  };

  return (
    <section className="stats-view terminal-scroll">
      <div className="stats-toolbar">
        <div>
          <p className="eyebrow">结团统计</p>
          <h2>{month} 结团数据</h2>
        </div>
        <label className="stats-field">
          <span>月份</span>
          <input type="month" value={month} onChange={e => setMonth(e.target.value || currentMonth())} />
        </label>
        <label className="stats-file">
          <span>{selectedFileText}</span>
          <input type="file" accept=".txt,text/plain" multiple onChange={onPickFiles} />
        </label>
        <button type="button" onClick={startImport} disabled={importing || !files.length}>{importing ? '导入中' : '开始导入'}</button>
        <button type="button" onClick={loadStats} disabled={loading}>{loading ? '刷新中' : '刷新'}</button>
      </div>

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
          {importResult?.items.length ? (
            <div className="stats-result-list">
              {importResult.items.map((item, i) => {
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
                  <td>{player.name || '未命名'}</td>
                  <td>{player.qq || '未记录'}</td>
                  <td>{player.game_count}</td>
                  <td>{formatHours(player.game_hours)}</td>
                  <td>{player.reincarnation_count}</td>
                  <td>{player.host_count}</td>
                  <td>{formatHours(player.host_hours)}</td>
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
              <div>
                <b>{session.title || '未命名团'}</b>
                <small>{session.source_filename || '未知来源'}</small>
              </div>
              <span>{formatHours(session.duration_hours)} 小时</span>
              <span>{session.kp_name || '未知 KP'}{session.kp_qq ? ` (${session.kp_qq})` : ''}</span>
              <span>{session.pl_count} PL</span>
              <span>{formatConfidence(session.confidence)}</span>
              <button type="button" className="danger" onClick={() => deleteSession(session.id)}>删除</button>
            </article>
          ))}
          {!sessions.items.length && <div className="stats-empty">暂无结团记录。</div>}
        </div>
      </section>

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
    </section>
  );
}
