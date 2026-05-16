import { ChangeKind, GitChanges } from '../../types';
import { kb } from '../../utils';

const changeLabels: Record<ChangeKind, string> = { added: '新增', modified: '修改', deleted: '删除', renamed: '移动/改名' };

export function ChangeSummaryPanel({ data, onCopy }: { data: GitChanges; onCopy: () => void }) {
  const kinds: ChangeKind[] = ['added', 'modified', 'deleted', 'renamed'];
  return (
    <section className="changes-panel">
      <div className="changes-head">
        <div><p className="eyebrow">CHANGELOG DRAFT</p><h3>上个 tag → latest 摘要</h3><small>{data.from_ref} → latest</small></div>
        <div className="stats-grid">
          <span><b>{data.stats.total}</b>总计</span>
          <span><b>{data.stats.added}</b>新增</span>
          <span><b>{data.stats.modified}</b>修改</span>
          <span><b>{data.stats.deleted}</b>删除</span>
          <span><b>{data.stats.renamed}</b>移动</span>
        </div>
        <button onClick={onCopy}>复制更新摘要</button>
      </div>
      <div className="change-groups">
        {kinds.map(kind => (
          <div className={`change-group ${kind}`} key={kind}>
            <h4>【{changeLabels[kind]}】<em>{data.readable[kind]?.length || 0}</em></h4>
            {data.readable[kind]?.length ? data.readable[kind].map(item => (
              <div className="change-item" key={`${kind}-${item.old_path || ''}-${item.path}`}>
                <b>{item.title}</b>
                <span>{item.category || '根目录'} · {item.root}{typeof item.size === 'number' ? ` · ${kb(item.size)}` : ''}</span>
                <small>{kind === 'renamed' ? `${item.old_path} → ${item.path}` : item.path}</small>
              </div>
            )) : <p className="no-change">无</p>}
          </div>
        ))}
      </div>
      <details className="raw-details"><summary>查看原始 JSON</summary><pre className="log terminal-scroll">{JSON.stringify(data, null, 2)}</pre></details>
    </section>
  );
}
