import { useEffect, useId, useRef, useState } from 'react';
import type { ReactNode } from 'react';

/** A small native dialog: escapes overflow clipping and keeps keyboard focus
 * inside its controls. Opening settings never adds another layout row. */
export function ActionPopover({ label, children, trigger, className = '' }: {
  label: string; children: ReactNode; trigger?: ReactNode; className?: string;
}) {
  const id = useId();
  const button = useRef<HTMLButtonElement>(null);
  const dialog = useRef<HTMLDialogElement>(null);
  const [open, setOpen] = useState(false);

  const position = () => {
    const node = dialog.current;
    const anchor = button.current;
    if (!node?.open || !anchor) return;
    const viewport = window.visualViewport;
    const leftEdge = viewport?.offsetLeft || 0;
    const topEdge = viewport?.offsetTop || 0;
    const width = viewport?.width || window.innerWidth;
    const height = viewport?.height || window.innerHeight;
    node.style.maxHeight = `${Math.max(80, Math.min(520, height - 16))}px`;
    node.style.width = `${Math.min(350, width - 16)}px`;
    const rect = anchor.getBoundingClientRect();
    const own = node.getBoundingClientRect();
    node.style.left = `${Math.max(leftEdge + 8, Math.min(rect.right - own.width, leftEdge + width - own.width - 8))}px`;
    const below = rect.bottom + 6;
    node.style.top = `${Math.max(topEdge + 8, Math.min(below + own.height <= topEdge + height - 8 ? below : rect.top - own.height - 6, topEdge + height - own.height - 8))}px`;
  };
  const close = () => {
    dialog.current?.close();
    setOpen(false);
    if (button.current?.isConnected && button.current.getClientRects().length) button.current.focus({ preventScroll: true });
  };
  useEffect(() => {
    const node = dialog.current;
    const observer = new ResizeObserver(position);
    if (node) observer.observe(node);
    window.addEventListener('resize', position);
    window.visualViewport?.addEventListener('resize', position);
    return () => {
      observer.disconnect();
      window.removeEventListener('resize', position);
      window.visualViewport?.removeEventListener('resize', position);
      node?.close();
    };
  }, []);

  return <>
    <button ref={button} type="button" className={`reader-popover-trigger ${className}`} aria-label={label} aria-haspopup="dialog" aria-controls={id} aria-expanded={open} onClick={() => {
      if (dialog.current?.open) close();
      else { dialog.current?.showModal(); setOpen(true); position(); }
    }}>{trigger || label}</button>
    <dialog id={id} ref={dialog} className="reader-popover" aria-labelledby={`${id}-title`} onCancel={event => { event.preventDefault(); event.stopPropagation(); close(); }} onClose={() => setOpen(false)} onClick={event => {
      const rect = event.currentTarget.getBoundingClientRect();
      if (event.target === event.currentTarget && (event.clientX < rect.left || event.clientX > rect.right || event.clientY < rect.top || event.clientY > rect.bottom)) close();
      else if ((event.target as HTMLElement).closest('[data-close-popover]')) close();
    }}>
      <header><h2 id={`${id}-title`}>{label}</h2><button type="button" autoFocus aria-label={`关闭${label}`} onClick={close}>×</button></header>
      <div className="reader-popover-body">{children}</div>
    </dialog>
  </>;
}
