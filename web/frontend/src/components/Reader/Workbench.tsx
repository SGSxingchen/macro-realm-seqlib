import { useEffect, useLayoutEffect, useRef, useState } from 'react';
import type { ComponentProps, ReactNode } from 'react';
import { api, routePath } from '../../api';
import type { Detail } from '../../types';
import { closeTab, readOpenTabs, rememberTab, storeOpenTabs } from '../../workbench-state';
import { ActionPopover } from '../ui/ActionPopover';
import { Reader } from './index';
import { HistoryCompare } from './HistoryCompare';

type Props = ComponentProps<typeof Reader> & {
  openPath: string; onOpen: (path: string) => void;
  onFocusMode: (value: boolean) => void; revision: number; layoutActions?: ReactNode;
};
type Mode = 'read' | 'compare' | 'history';

export function ReaderWorkbench({ openPath, onOpen, onFocusMode, revision, layoutActions, compact = false, ...readerProps }: Props) {
  const [tabs, setTabs] = useState(() => rememberTab(readOpenTabs(), openPath, readerProps.detail?.title));
  const [mode, setMode] = useState<Mode>('read');
  const [referencePath, setReferencePath] = useState('');
  const [reference, setReference] = useState<Detail | null>(null);
  const [referenceError, setReferenceError] = useState('');
  const [referenceLoading, setReferenceLoading] = useState(false);
  const [referenceRetry, setReferenceRetry] = useState(0);
  const [referenceAnchor, setReferenceAnchor] = useState('');
  const [phonePane, setPhonePane] = useState<'primary' | 'reference'>('primary');
  const previousPath = useRef(openPath);
  const tabStrip = useRef<HTMLDivElement>(null);

  useEffect(() => {
    setTabs(current => rememberTab(current, openPath, readerProps.detail?.path === openPath ? readerProps.detail.title : undefined));
    if (openPath !== previousPath.current) {
      if (mode === 'compare' && openPath === referencePath) setReferencePath(previousPath.current);
      previousPath.current = openPath;
      setPhonePane('primary');
    }
    if (!openPath) setMode('read');
  }, [openPath, readerProps.detail?.title, mode, referencePath]);
  useEffect(() => { storeOpenTabs(tabs); }, [tabs]);
  useEffect(() => { onFocusMode(mode !== 'read'); return () => onFocusMode(false); }, [mode, onFocusMode]);
  useLayoutEffect(() => {
    const strip = tabStrip.current;
    const active = strip?.querySelector<HTMLElement>('[aria-current="page"]');
    if (strip && active) {
      const left = active.offsetLeft - strip.offsetLeft;
      if (left < strip.scrollLeft) strip.scrollLeft = left;
      else if (left + active.offsetWidth > strip.scrollLeft + strip.clientWidth) strip.scrollLeft = left + active.offsetWidth - strip.clientWidth;
    }
  }, [openPath, tabs.length, compact, mode]);

  useEffect(() => {
    const controller = new AbortController();
    setReference(null); setReferenceError(''); setReferenceAnchor('');
    if (mode !== 'compare' || !referencePath) { setReferenceLoading(false); return () => controller.abort(); }
    setReferenceLoading(true);
    api<Detail>('/api/resources/' + routePath(referencePath), { signal: controller.signal }).then(data => {
      if (!controller.signal.aborted) setReference(data);
    }).catch(error => { if (!controller.signal.aborted) setReferenceError(error instanceof Error ? error.message : '参照档案读取失败'); })
      .finally(() => { if (!controller.signal.aborted) setReferenceLoading(false); });
    return () => controller.abort();
  }, [referencePath, mode, revision, referenceRetry]);

  const remove = (path: string) => {
    const next = closeTab(tabs, path, openPath);
    setTabs(next.tabs);
    if (path === referencePath) { setReferencePath(''); setMode('read'); }
    if (next.active !== openPath) onOpen(next.active);
  };
  const compare = () => {
    if (mode === 'compare') { setMode('read'); return; }
    const path = referencePath && referencePath !== openPath && tabs.some(tab => tab.path === referencePath) ? referencePath : [...tabs].reverse().find(tab => tab.path !== openPath)?.path || '';
    setReferencePath(path); setMode('compare'); setPhonePane('primary');
  };
  const peers = tabs.filter(tab => tab.path !== openPath);
  const isSplit = mode === 'compare' && !!referencePath && referencePath !== openPath;

  return <section className={`realm-reader realm-reader-workbench${compact ? ' compact-workbench' : ''}`} aria-label="多资源阅读工作台">
    {(tabs.length > 0 || openPath || layoutActions) && <div className="workbench-chrome">
      <div className="workbench-tab-strip" ref={tabStrip} aria-label={`已打开的档案，共 ${tabs.length} 份`}>
        {tabs.map(tab => <div className={`workbench-tab${tab.path === openPath ? ' active' : ''}`} key={tab.path}>
          <button className="workbench-tab-title" data-resource-tab={tab.path} aria-current={tab.path === openPath ? 'page' : undefined} title={tab.title} onClick={() => onOpen(tab.path)}>{tab.title}</button>
          <button className="workbench-tab-close" aria-label={`关闭档案：${tab.title}`} onClick={() => remove(tab.path)}>×</button>
        </div>)}
      </div>
      <div className="workbench-controls">
        {openPath && <>
          <button aria-label={mode === 'compare' ? '退出并排' : '并排对照'} aria-pressed={mode === 'compare'} disabled={!peers.length && mode !== 'compare'} title={!peers.length ? '先打开另一份资源，它会保留在标签中' : '固定参照，切换主档案'} onClick={compare}>{mode === 'compare' ? '退出并排' : '并排对照'}</button>
          {isSplit && <ActionPopover label="对照设置" trigger="参照设置">
            <div className="workbench-compare-controls">
              <label>固定参照<select aria-label="选择参照档案" value={referencePath} onChange={event => setReferencePath(event.target.value)}>{peers.map(tab => <option key={tab.path} value={tab.path}>{tab.title}</option>)}</select></label>
              <button data-close-popover onClick={() => { const old = openPath; onOpen(referencePath); setReferencePath(old); }}>交换主 / 参照</button>
              <p>主档案跟随上方标签；参照保持固定。两侧独立滚动，不强行对齐不同模板。</p>
            </div>
          </ActionPopover>}
          <button aria-label={mode === 'history' ? '返回正文' : '历史版本'} aria-pressed={mode === 'history'} onClick={() => setMode(current => current === 'history' ? 'read' : 'history')}>{mode === 'history' ? '返回正文' : '历史版本'}</button>
        </>}
        {layoutActions}
      </div>
      {isSplit && <div className="workbench-phone-panes" role="group" aria-label="切换对照窗格"><button aria-pressed={phonePane === 'primary'} onClick={() => setPhonePane('primary')}>主档案</button><button aria-pressed={phonePane === 'reference'} onClick={() => setPhonePane('reference')}>参照档案</button></div>}
    </div>}
    <div className={`workbench-reading-panes ${isSplit ? 'split' : ''} phone-${phonePane}`} hidden={mode === 'history'}>
      <div className="workbench-primary">
        <Reader {...readerProps} compact={compact || isSplit} paneLabel={isSplit ? '主' : undefined} instance="primary" />
      </div>
      {isSplit && <div className="workbench-reference">
        <Reader detail={reference} loading={referenceLoading || (!reference && !referenceError)} error={referenceError} onRetry={() => setReferenceRetry(value => value + 1)} anchor={referenceAnchor} onAnchor={setReferenceAnchor} instance="reference" idPrefix="reference-" compact paneLabel="参照" />
      </div>}
    </div>
    {mode === 'history' && openPath && <HistoryCompare key={openPath} path={openPath} title={readerProps.detail?.title || tabs.find(tab => tab.path === openPath)?.title || openPath} />}
  </section>;
}
