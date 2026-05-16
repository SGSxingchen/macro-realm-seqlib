import React, { useEffect, useState } from 'react';
import { api, routePath } from '../../api';
import {
  AdminState, AdminTab, CmdResult, Detail, GitChanges, GitInfo,
  HonorCategoriesResponse, HonorCategory, PackageFile, PublishResult,
} from '../../types';
import { kb, suggestHonorCategory } from '../../utils';
import { StatusBadge } from '../ui/StatusBadge';
import { CommandSteps } from './CommandSteps';
import { ChangeSummaryPanel } from './ChangeSummaryPanel';

function HumanLog({ title, data }: { title: string; data: unknown }) {
  if (!data) return null;
  return <details className="raw-details admin-raw"><summary>{title}</summary><pre className="log terminal-scroll">{typeof data === 'string' ? data : JSON.stringify(data, null, 2)}</pre></details>;
}

type Props = { detail: Detail | null; reload: () => void; onResourceMoved: () => void; onBackToRead: () => void };

export function AdminPanel({ detail, reload, onResourceMoved, onBackToRead }: Props) {
  const [me, setMe] = useState<AdminState>({ admin_configured: false, authenticated: false });
  const [password, setPassword] = useState('');
  const [tab, setTab] = useState<AdminTab>('overview');
  const [edit, setEdit] = useState('');
  const [notice, setNotice] = useState('');
  const [rawLog, setRawLog] = useState<unknown>(null);
  const [gitInfo, setGitInfo] = useState<GitInfo | null>(null);
  const [changes, setChanges] = useState<GitChanges | null>(null);
  const [newPath, setNewPath] = useState('序列库/新分类/001】新资源.txt');
  const [movePath, setMovePath] = useState(detail?.path || '');
  const [honorCategories, setHonorCategories] = useState<HonorCategory[]>([]);
  const [honorCategory, setHonorCategory] = useState('其他');
  const [honorTitle, setHonorTitle] = useState('');
  const [honorSuggested, setHonorSuggested] = useState('其他');
  const [version, setVersion] = useState('v6.5');
  const [message, setMessage] = useState('更新序列库 latest');
  const [uploadFile, setUploadFile] = useState<File | null>(null);
  const [publishResult, setPublishResult] = useState<PublishResult | null>(null);
  const [wikiResult, setWikiResult] = useState<CmdResult | null>(null);
  const [packageFiles, setPackageFiles] = useState<PackageFile[]>([]);
  const [packageFile, setPackageFile] = useState<PackageFile | null>(null);
  const [packageContent, setPackageContent] = useState('');
  const [newPackageName, setNewPackageName] = useState('新版序列库说明.txt');

  useEffect(() => { api<AdminState>('/api/admin/me').then(setMe).catch(() => {}); }, []);
  useEffect(() => {
    const suggested = detail ? suggestHonorCategory(detail.path, detail.category) : '其他';
    setEdit(detail?.content || '');
    setMovePath(detail?.path || '');
    setHonorTitle(detail?.filename ? detail.filename.replace(/^\d+】/, '').replace(/\.txt$/i, '') : '');
    setHonorSuggested(suggested);
    setHonorCategory(suggested);
  }, [detail?.path, detail?.content, detail?.filename, detail?.category]);
  const refreshGit = async () => { const data = await api<GitInfo>('/api/git/info'); setGitInfo(data); return data; };
  useEffect(() => {
    if (me.authenticated) {
      refreshGit().catch(() => {});
      api<HonorCategoriesResponse>('/api/admin/honor-categories' + (detail?.path ? `?path=${encodeURIComponent(detail.path)}` : ''))
        .then(r => { setHonorCategories(r.items); if (!detail && r.suggested_category) { setHonorSuggested(r.suggested_category); setHonorCategory(r.suggested_category); } })
        .catch(() => {});
    }
  }, [me.authenticated]);

  const runAction = async (label: string, fn: () => Promise<unknown>) => {
    try { const data = await fn(); setNotice(`${label}完成`); setRawLog(data); reload(); refreshGit().catch(() => {}); return data; }
    catch (e: unknown) { const msg = e instanceof Error ? e.message : String(e); setNotice(`${label}失败：${msg}`); setRawLog(msg); }
  };

  const loadChanges = async () => { const data = await api<GitChanges>('/api/git/changes'); setChanges(data); setRawLog(null); return data; };
  const copyChanges = async () => { if (changes) await navigator.clipboard?.writeText(changes.markdown || changes.text).catch(() => {}); };
  const upload = async () => {
    if (!detail || !uploadFile) return null;
    const form = new FormData(); form.append('file', uploadFile);
    return api(`/api/admin/upload/${routePath(detail.path)}`, { method: 'POST', body: form });
  };
  const doWiki = async (dryRun: boolean) => { const data = await api<CmdResult>(`/api/admin/wiki/sync?dry_run=${dryRun ? 'true' : 'false'}`, { method: 'POST' }); setWikiResult(data); setRawLog(data); return data; };
  const doPublish = async () => { const data = await api<PublishResult>('/api/admin/publish', { method: 'POST', body: JSON.stringify({ version, message, push: true }) }); setPublishResult(data); setRawLog(data); refreshGit().catch(() => {}); return data; };
  const loadPackageFiles = async () => { const data = await api<{ items: PackageFile[] }>('/api/admin/package-files'); setPackageFiles(data.items); return data; };
  const openPackageFile = async (path: string) => { const data = await api<PackageFile>('/api/admin/package-files/' + routePath(path)); setPackageFile(data); setPackageContent(data.content || ''); return data; };
  const savePackageFile = async () => { if (!packageFile) return null; return api('/api/admin/package-files/' + routePath(packageFile.path), { method: 'PUT', body: JSON.stringify({ path: packageFile.path, content: packageContent }) }); };
  const createPackageFile = async () => api('/api/admin/package-files', { method: 'POST', body: JSON.stringify({ path: newPackageName, content: packageContent, overwrite: false }) });
  useEffect(() => { if (me.authenticated && tab === 'package' && packageFiles.length === 0) loadPackageFiles().catch(() => {}); }, [me.authenticated, tab]);
  const selectedHonor = honorCategories.find(c => c.category === honorCategory);
  const honorPreview = detail && selectedHonor ? `荣誉室/${honorCategory}/${selectedHonor.next_prefix}${honorTitle || detail.title}.txt` : '';
  const moveCurrentToHonor = async () => {
    if (!detail) return null;
    const data = await api('/api/admin/move-to-honor', { method: 'POST', body: JSON.stringify({ path: detail.path, category: honorCategory, title: honorTitle }) });
    onResourceMoved();
    return data;
  };

  if (!me.authenticated) return (
    <section className="admin-card auth-panel console-auth">
      <div className="section-head">
        <div><p className="eyebrow">MAINTENANCE CONSOLE</p><h2>维护控制台接入</h2></div>
        <button type="button" className="ghost-toggle" onClick={onBackToRead}>返回查阅</button>
      </div>
      <p className="hint">{me.admin_configured ? '输入 ADMIN_PASSWORD 后进入维护控制台。' : '未配置 ADMIN_PASSWORD，后台写操作禁用。'}</p>
      <div className="row">
        <input type="password" placeholder="ADMIN_PASSWORD" value={password} onChange={e => setPassword(e.target.value)} />
        <button type="button" onClick={() => runAction('登录', async () => { const r = await api('/api/admin/login', { method: 'POST', body: JSON.stringify({ password }) }); setMe({ ...me, authenticated: true }); return r; })}>解锁控制台</button>
      </div>
      <HumanLog title="查看登录响应" data={rawLog} />
    </section>
  );

  const dirty = !!gitInfo?.is_dirty;
  const ab = gitInfo?.ahead_behind;
  const normalizedVersion = version.startsWith('v') ? version : `v${version}`;
  const tagExists = !!gitInfo?.latest_tag && gitInfo.latest_tag === normalizedVersion;
  const headHasVersionTag = !!gitInfo?.head_tags?.includes(normalizedVersion);
  const tabs: Array<[AdminTab, string]> = [
    ['overview', '总览'], ['edit', '资源编辑'], ['package', '发行文件'],
    ['changes', '变更摘要'], ['publish', '发布版本'], ['wiki', 'Wiki 同步'],
  ];

  return (
    <section className="admin-console">
      <aside className="admin-nav">
        <p className="eyebrow">MAINTENANCE</p>
        <h2>维护控制台</h2>
        <button type="button" className="back-read" onClick={onBackToRead}>← 返回查阅</button>
        {tabs.map(([id, label]) => <button type="button" key={id} className={tab === id ? 'active' : ''} onClick={() => setTab(id)}>{label}</button>)}
        <div className="console-mini">
          <StatusBadge ok={!dirty} warn={dirty}>{dirty ? '工作区有改动' : '工作区干净'}</StatusBadge>
          <StatusBadge ok={!!gitInfo?.wiki_configured} warn={!gitInfo?.wiki_configured}>{gitInfo?.wiki_configured ? 'Wiki 已配置' : 'Wiki 未配置'}</StatusBadge>
        </div>
      </aside>
      <div className="admin-main terminal-scroll">
        <div className="console-top">
          <div><p className="eyebrow">CONTROL CENTER</p><h2>{tabs.find(t => t[0] === tab)?.[1]}</h2></div>
          <button onClick={() => runAction('刷新 Git 状态', refreshGit)}>刷新状态</button>
        </div>
        {notice && <div className="notice-line">{notice}</div>}
        {tab === 'overview' && (
          <div className="console-grid">
            <section className="console-card wide">
              <h3>Git 状态</h3>
              <div className="kv-grid">
                <span>分支</span><b>{gitInfo?.branch_name || '未知'}</b>
                <span>HEAD</span><b>{gitInfo?.head_short || '未知'}</b>
                <span>最近 tag</span><b>{gitInfo?.latest_tag || '无'}</b>
                <span>跟踪分支</span><b>{gitInfo?.tracking || '未设置'}</b>
                <span>origin/main</span><b>{gitInfo?.remote_main || '不可用'}</b>
              </div>
              <div className="badge-row">
                <StatusBadge ok={!dirty} warn={dirty}>{dirty ? '工作区有未提交改动' : '工作区干净'}</StatusBadge>
                {ab && <StatusBadge ok={ab.ahead === 0 && ab.behind === 0} warn={ab.ahead > 0 || ab.behind > 0}>ahead {ab.ahead} / behind {ab.behind}</StatusBadge>}
              </div>
              <pre className="status-text">{gitInfo?.status_branch || gitInfo?.status_short || '暂无状态'}</pre>
            </section>
            <section className="console-card">
              <h3>后台环境</h3>
              <div className="badge-col">
                <StatusBadge ok={!!me.admin_configured}>ADMIN_PASSWORD {me.admin_configured ? '已配置' : '未配置'}</StatusBadge>
                <StatusBadge ok={!!gitInfo?.wiki_configured} warn={!gitInfo?.wiki_configured}>Wiki 凭据 {gitInfo?.wiki_configured ? '已配置' : '未配置'}</StatusBadge>
                <StatusBadge ok={true}>公开前台仅序列库</StatusBadge>
              </div>
            </section>
            <section className="console-card">
              <h3>当前资源</h3>
              {detail ? (
                <div className="resource-summary">
                  <b>{detail.title}</b><small>{detail.path}</small>
                  <span>{detail.category || '根目录'} · {kb(detail.size)}</span>
                </div>
              ) : <p className="hint">尚未选择资源。请回到查阅页选择资源，或进入"资源编辑"新增路径。</p>}
            </section>
          </div>
        )}
        {tab === 'edit' && (
          <div className="edit-layout">
            <section className="console-card">
              <h3>当前选中资源</h3>
              {detail ? (
                <div className="kv-grid">
                  <span>标题</span><b>{detail.title}</b>
                  <span>路径</span><b>{detail.path}</b>
                  <span>分类</span><b>{detail.category || '根目录'}</b>
                  <span>大小</span><b>{kb(detail.size)}</b>
                </div>
              ) : <p className="hint">没有选中资源：可先去查阅页点击资源，或在下方新增资源路径后创建。</p>}
              <label>新增资源路径</label>
              <input value={newPath} onChange={e => setNewPath(e.target.value)} />
              <button onClick={() => runAction('新增资源', () => api('/api/admin/resources', { method: 'POST', body: JSON.stringify({ path: newPath, content: edit, overwrite: false }) }))}>以编辑区内容新增 TXT</button>
            </section>
            <section className="console-card wide">
              <h3>内容编辑区</h3>
              <textarea className="editor terminal-scroll" value={edit} onChange={e => setEdit(e.target.value)} placeholder="选择资源后在此编辑；保存会覆盖真实 txt。" />
              <div className="action-row">
                <button disabled={!detail} onClick={() => confirm('确认用编辑区内容覆盖真实 TXT 文件？') && runAction('保存资源', () => api(`/api/admin/resources/${routePath(detail!.path)}`, { method: 'PUT', body: JSON.stringify({ path: detail!.path, content: edit }) }))}>保存：覆盖真实 TXT</button>
                <input type="file" accept=".txt,text/plain" onChange={e => setUploadFile(e.target.files?.[0] || null)} />
                <button disabled={!detail || !uploadFile} onClick={() => confirm('确认上传 TXT 覆盖当前资源？') && runAction('上传覆盖', upload)}>上传覆盖</button>
              </div>
              <label>移动/重命名目标路径</label>
              <div className="row">
                <input value={movePath} onChange={e => setMovePath(e.target.value)} />
                <button disabled={!detail || !movePath} onClick={() => confirm(`确认移动/重命名到：${movePath}？`) && runAction('移动资源', () => api('/api/admin/move', { method: 'POST', body: JSON.stringify({ old_path: detail!.path, new_path: movePath }) }))}>移动/重命名</button>
              </div>
              <div className="honor-move-box">
                <h4>移入荣誉室</h4>
                <p className="hint">荣誉室只使用一级大类；系统会按目标大类自动取下一个编号。</p>
                <div className="suggestion-line">根据当前路径建议归入：<b>{honorSuggested}</b></div>
                <label>荣誉室大类</label>
                <select value={honorCategory} onChange={e => setHonorCategory(e.target.value)}>
                  {honorCategories.map(c => <option key={c.category} value={c.category}>{c.category}（现有 {c.count}，下一个 {c.next_prefix}）</option>)}
                </select>
                <label>标题主体</label>
                <input value={honorTitle} onChange={e => setHonorTitle(e.target.value)} placeholder="默认使用原文件名去编号" />
                <div className="path-preview">目标预览：{honorPreview || '请先选择资源'}</div>
                <button disabled={!detail || !honorCategory || !honorTitle} className="danger" onClick={() => confirm(`确认移动到 ${honorPreview}？移入后前台不再显示。`) && runAction('移入荣誉室', moveCurrentToHonor)}>按大类重新编号并移入荣誉室</button>
              </div>
              <div className="danger-zone">
                <button disabled={!detail} className="danger" onClick={() => confirm('危险：确认删除真实文件？此操作不可由前端撤销。') && runAction('删除资源', () => api('/api/admin/delete', { method: 'POST', body: JSON.stringify({ path: detail!.path }) }))}>删除真实文件</button>
              </div>
            </section>
          </div>
        )}
        {tab === 'package' && (
          <div className="package-layout">
            <section className="console-card">
              <div className="section-head">
                <div><h3>ZIP 根目录发行文件</h3><p className="hint">这些是发行包根目录说明文件，不会出现在前台资源列表。</p></div>
                <button onClick={() => runAction('刷新发行文件', loadPackageFiles)}>刷新列表</button>
              </div>
              <div className="package-list">
                {packageFiles.map(f => (
                  <button key={f.path} className={packageFile?.path === f.path ? 'package-item active' : 'package-item'} onClick={() => openPackageFile(f.path)}>
                    <b>{f.title}</b>
                    <small>{f.path} · {kb(f.size)} · {f.editable ? '可编辑' : '只读'}</small>
                  </button>
                ))}
              </div>
            </section>
            <section className="console-card wide">
              <h3>编辑发行文件</h3>
              {packageFile ? (
                <>
                  <div className="kv-grid">
                    <span>文件</span><b>{packageFile.path}</b>
                    <span>状态</span><b>{packageFile.editable ? '根目录 TXT，可保存' : '只读展示'}</b>
                  </div>
                  <p className="hint">这是 ZIP/发行包根目录文件。保存会覆盖仓库根目录同名 .txt。</p>
                  <textarea className="editor terminal-scroll" disabled={!packageFile.editable} value={packageContent} onChange={e => setPackageContent(e.target.value)} />
                  <button disabled={!packageFile.editable} onClick={() => confirm(`确认覆盖根目录发行文件 ${packageFile.path}？`) && runAction('保存发行文件', savePackageFile)}>保存发行文件</button>
                </>
              ) : <p className="hint">点击左侧文件开始编辑；非 .txt 文件只读展示。</p>}
              <div className="new-package-box">
                <h4>新增根目录 TXT</h4>
                <label>文件名</label>
                <input value={newPackageName} onChange={e => setNewPackageName(e.target.value)} placeholder="例如 V6.5序列库更新日志.txt" />
                <button onClick={() => confirm(`确认新增根目录发行文件 ${newPackageName}？`) && runAction('新增发行文件', createPackageFile)}>以当前编辑区内容新增</button>
              </div>
            </section>
          </div>
        )}
        {tab === 'changes' && (
          <section className="console-card wide">
            <div className="section-head">
              <div><h3>上个 tag → latest 更新草稿</h3><p className="hint">按新增、修改、删除、移动分组，不包含全文 diff。</p></div>
              <button onClick={() => runAction('生成变更摘要', loadChanges)}>生成/刷新摘要</button>
            </div>
            {changes ? <ChangeSummaryPanel data={changes} onCopy={copyChanges} /> : <p className="hint">点击生成摘要后可复制为更新日志草稿。</p>}
          </section>
        )}
        {tab === 'publish' && (
          <div className="console-grid">
            <section className="console-card wide">
              <h3>发布前检查</h3>
              <div className="badge-col">
                <StatusBadge ok={!dirty} warn={dirty}>{dirty ? '工作区有未提交改动：发布会尝试 git add/commit' : '工作区干净'}</StatusBadge>
                <StatusBadge ok={!tagExists} warn={tagExists}>{tagExists ? `版本 tag ${normalizedVersion} 已存在，创建 tag 可能失败` : '输入版本 tag 当前未匹配最近 tag'}</StatusBadge>
                <StatusBadge ok={headHasVersionTag} warn={!headHasVersionTag}>{headHasVersionTag ? 'HEAD 已在输入版本 tag 上' : 'HEAD 不在输入版本 tag 上：如 tag 已存在需人工确认'}</StatusBadge>
                {ab && <StatusBadge ok={ab.behind === 0} warn={ab.behind > 0}>远端同步：ahead {ab.ahead} / behind {ab.behind}</StatusBadge>}
              </div>
            </section>
            <section className="console-card">
              <h3>发布参数</h3>
              <label>版本号</label>
              <input value={version} onChange={e => setVersion(e.target.value)} />
              <label>提交/tag 信息</label>
              <input value={message} onChange={e => setMessage(e.target.value)} />
              <p className="hint">将执行：git status → git add 白名单 → git commit → git tag → git push origin main --tags。</p>
              <button className="publish" onClick={() => confirm('确认提交当前改动、打 tag 并推送到 GitHub？') && runAction('发布版本', doPublish)}>提交当前改动并打 tag / 推送到 GitHub</button>
            </section>
            <section className="console-card wide">
              <h3>发布执行结果</h3>
              {publishResult ? <><StatusBadge ok={publishResult.ok}>{publishResult.ok ? '发布流程成功' : '发布流程失败/部分失败'}</StatusBadge><CommandSteps steps={publishResult.steps} /></> : <p className="hint">尚未执行发布。</p>}
            </section>
          </div>
        )}
        {tab === 'wiki' && (
          <div className="console-grid">
            <section className="console-card">
              <h3>Wiki 凭据</h3>
              <StatusBadge ok={!!gitInfo?.wiki_configured} warn={!gitInfo?.wiki_configured}>{gitInfo?.wiki_configured ? '已配置 WIKI/FANDOM 凭据' : '未配置 Wiki 凭据'}</StatusBadge>
              <p className="hint">前端不会显示密钥明文；后端只读取环境变量。</p>
            </section>
            <section className="console-card">
              <h3>同步操作</h3>
              <button onClick={() => runAction('Wiki dry-run', () => doWiki(true))}>执行 dry-run（不写入 Wiki）</button>
              <button className="danger" disabled={!gitInfo?.wiki_configured} onClick={() => confirm('确认执行真实 Wiki 同步？会写入线上 Wiki。') && runAction('真实 Wiki 同步', () => doWiki(false))}>真实同步到 Wiki</button>
            </section>
            <section className="console-card wide">
              <h3>同步日志</h3>
              {wikiResult ? <CommandSteps steps={[wikiResult]} /> : <p className="hint">尚未执行同步。</p>}
            </section>
          </div>
        )}
        <HumanLog title="查看原始日志 / API 响应" data={rawLog} />
      </div>
    </section>
  );
}
