import React, { useMemo, useState } from 'react';
import { Detail } from '../../types';
import { kb, words } from '../../utils';
import { parseDocument } from './parser';
import { Structured } from './Structured';

export function Reader({ detail }: { detail: Detail | null }) {
  const [mode, setMode] = useState<'structured' | 'raw'>('structured');
  const parsed = useMemo(() => detail ? parseDocument(detail.content) : null, [detail?.content]);
  const copy = (text: string) => navigator.clipboard?.writeText(text).catch(() => {});

  if (!detail) return (
    <article className="reader empty-reader">
      <div className="empty-sigil">∴</div>
      <h2>等待调阅序列档案</h2>
      <p>从左侧分类或顶部搜索结果选择资源。公开终端仅显示「序列库」，荣誉室记录已从前台隔离。</p>
      <div className="empty-tips">
        <small>提示</small>
        <ul>
          <li>支持拼音首字母（如 <code>bfz</code>）</li>
          <li>支持模糊（<code>强驱散</code> → 强制驱散）</li>
          <li>支持中英混搜（<code>Centurion 百夫长</code>）</li>
          <li>多个词用空格隔开做 AND 检索</li>
          <li>按 <code>Ctrl/⌘ + K</code> 唤起搜索</li>
        </ul>
      </div>
    </article>
  );

  return (
    <article className="reader terminal-scroll">
      <div className="reader-head">
        <h2>{detail.title}</h2>
        <div className="breadcrumbs">
          {detail.path.split('/').map((p, i, arr) => (
            <React.Fragment key={`${p}-${i}`}>
              <span>{p}</span>
              {i < arr.length - 1 && <i>/</i>}
            </React.Fragment>
          ))}
        </div>
      </div>
      <div className="detail-toolbar">
        <span>{detail.category || '根目录'}</span>
        <span>{kb(detail.size)}</span>
        <span>{words(detail.content)} 字</span>
        <span>{detail.encoding}</span>
        <div className="toolbar-spacer" />
        <div className="mode-toggle">
          <button type="button" className={mode === 'structured' ? 'active' : ''} onClick={() => setMode('structured')}>结构化</button>
          <button type="button" className={mode === 'raw' ? 'active' : ''} onClick={() => setMode('raw')}>原文</button>
        </div>
        <button type="button" onClick={() => copy(detail.path)}>复制路径</button>
        <button type="button" onClick={() => copy(detail.content)}>复制全文</button>
      </div>
      {mode === 'structured' && parsed
        ? <Structured blocks={parsed.blocks} />
        : <section className="document"><pre>{detail.content}</pre></section>}
    </article>
  );
}
