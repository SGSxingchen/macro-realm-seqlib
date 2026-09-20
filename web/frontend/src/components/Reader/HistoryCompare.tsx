import { useEffect, useRef, useState } from 'react';
import { api, buildQuery, routePath } from '../../api';
import type { ChangeDetail, GitChanges } from '../../types';

type Candidate = { path: string; reason: string };

/** Reuses existing read-only Git endpoints; never asserts file identity from a name. */
export function HistoryCompare({ path, title }: { path: string; title: string }) {
  const [baseline, setBaseline] = useState('');
  const [oldPath, setOldPath] = useState(path);
  const [candidates, setCandidates] = useState<Candidate[]>([]);
  const [searched, setSearched] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState('');
  const [diff, setDiff] = useState<ChangeDetail | null>(null);
  const request = useRef<AbortController | null>(null);
  const requestVersion = useRef(0);
  useEffect(() => () => { request.current?.abort(); }, []);

  const invalidate = () => {
    request.current?.abort();
    requestVersion.current++;
    setBusy(false); setDiff(null); setError('');
  };
  const start = () => {
    request.current?.abort();
    const controller = new AbortController();
    const version = ++requestVersion.current;
    request.current = controller;
    setBusy(true); setError(''); setDiff(null);
    return { controller, current: () => !controller.signal.aborted && requestVersion.current === version };
  };
  const discover = async () => {
    const { controller, current } = start();
    setCandidates([]); setSearched(false);
    try {
      const result = await api<GitChanges>('/api/git/changes' + buildQuery({ from_ref: baseline.trim(), public_only: true }), { signal: controller.signal });
      if (!current()) return;
      const raw = result.raw as { committed?: { returncode?: number } } | null;
      if (raw?.committed?.returncode) throw new Error('Git 无法读取此基准版本。请检查标签是否存在、服务器是否包含历史记录。');
      setBaseline(result.from_ref);
      const renamed = result.readable.renamed.filter(item => item.path === path && item.old_path).map(item => ({ path: item.old_path!, reason: 'Git 变更记录中的改名候选（仍需核对内容）' }));
      setCandidates(renamed.filter((item, index) => renamed.findIndex(other => other.path === item.path) === index));
      setSearched(true);
    } catch (e) { if (current()) setError(e instanceof Error ? e.message : '查找失败'); }
    finally { if (current()) setBusy(false); }
  };
  const compare = async (candidatePath = oldPath) => {
    if (!candidatePath.startsWith('序列库/') || !candidatePath.toLowerCase().endsWith('.txt') || candidatePath.split('/').includes('..')) {
      setError('旧路径必须是序列库内的完整 TXT 路径。'); return;
    }
    const { controller, current } = start();
    setOldPath(candidatePath);
    try {
      const result = await api<ChangeDetail>('/api/git/change-detail/' + routePath(path) + buildQuery({ from_ref: baseline.trim(), old_path: candidatePath, kind: candidatePath === path ? 'modified' : 'renamed', public_only: true }), { signal: controller.signal });
      if (current()) setDiff(result);
    } catch (e) { if (current()) setError(e instanceof Error ? e.message : '对照失败'); }
    finally { if (current()) setBusy(false); }
  };

  return <section className="realm-history" aria-label="历史版本对照">
    <header><span className="realm-overline">VERSION COMPARISON / READ ONLY</span><h2>历史版本对照</h2><p>{title}</p></header>
    <div className="history-caution">比较「指定版本的旧文件」与「服务器当前档案」。同路径、同名称、Git 相似度均不是身份保证；请先核对旧路径和正文，尤其是重排编号或全面重置的资源。</div>
    <form onSubmit={e => { e.preventDefault(); void compare(); }}>
      <label>基准版本 / Git 标签<input aria-label="历史基准版本" placeholder="例如 v6.6.1；留空使用后端默认基线" value={baseline} onChange={e => { invalidate(); setBaseline(e.target.value); setCandidates([]); setSearched(false); }} /></label>
      <button type="button" disabled={busy} onClick={() => void discover()}>查找改名候选</button>
      <label className="history-path-field">旧版文件路径<input aria-label="旧版文件路径" value={oldPath} onChange={e => { invalidate(); setOldPath(e.target.value); }} /></label>
      <button type="submit" disabled={busy}>按此路径比较</button>
    </form>
    {candidates.length > 0 && <div className="history-candidates"><h3>候选路径：不会自动确认关联</h3>{candidates.map(item => <div key={item.path}><p>{item.path}</p><small>{item.reason}</small><button disabled={busy} onClick={() => void compare(item.path)}>使用此候选比较</button></div>)}</div>}
    {searched && !candidates.length && <p className="history-status">没有找到改名候选。这不代表资源从未存在；可填写已知旧路径，或换一个基准版本。此入口不声称覆盖跨多次改名的完整历史链。</p>}
    {busy && <p role="status">正在读取版本记录…</p>}
    {error && <p role="alert" className="realm-error">{error}</p>}
    {diff && <div className="history-result">
      <div className="history-sources"><p><b>旧版</b> {diff.from_ref}<br />{diff.old_path || path}</p><p><b>当前</b> 服务器工作区 / latest<br />{diff.path}</p></div>
      {!diff.old_exists || !diff.new_exists ? <p role="alert" className="history-caution">{!diff.old_exists ? '旧版本中找不到指定文件。请核对基准版本和旧路径；不能把缺失记录当作整份资源全部新增。' : '当前文件不可用，无法完成比较。'}</p> : <>
        <p className="history-status">文本差异：新增 {diff.additions} 行 / 删除 {diff.deletions} 行。纯文本差异不等于强度或规则变更判断。</p>
        {diff.truncated && <p role="alert" className="history-caution">差异内容已被接口截断；下面不是完整差异。</p>}
        <div className="history-diff" aria-label="历史文本差异">{diff.rows.map((row, index) => row.type === 'gap' ? <div className="history-gap" key={index}>… 未变更的上下文已省略 …</div> : <div className={`history-line ${row.type}`} key={index}><span className="history-line-number" aria-hidden="true">{row.old_no ?? '–'} / {row.new_no ?? '–'}</span><span className="history-change-sign" aria-label={row.type === 'added' ? '新增' : row.type === 'removed' ? '删除' : '未变'}>{row.type === 'added' ? '+' : row.type === 'removed' ? '−' : ' '}</span><pre>{row.text ?? ''}</pre></div>)}</div>
        {!diff.additions && !diff.deletions && <p>接口未报告文本差异。</p>}
      </>}
    </div>}
    <footer>只读取版本记录，不还原、不覆盖文件。暂不支持任意两个旧版本互比或自动确认资源谱系。</footer>
  </section>;
}
