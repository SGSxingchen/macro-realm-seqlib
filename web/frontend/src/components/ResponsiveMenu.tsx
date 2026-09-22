import { useEffect, useId, useRef, useState } from 'react';
import type { ReactNode } from 'react';

/** Inline actions on desktop, one non-modal disclosure on small screens. */
export function ResponsiveMenu({ label, className, children, closeOnSelect = false }: {
  label: string; className: string; children: ReactNode; closeOnSelect?: boolean;
}) {
  const [open, setOpen] = useState(false);
  const id = useId();
  const root = useRef<HTMLDivElement>(null);
  const trigger = useRef<HTMLButtonElement>(null);
  useEffect(() => {
    const media = window.matchMedia('(max-width: 820px)');
    const reset = () => { if (!media.matches) setOpen(false); };
    media.addEventListener('change', reset);
    return () => media.removeEventListener('change', reset);
  }, []);
  useEffect(() => {
    if (!open) return;
    const outside = (event: PointerEvent) => {
      if (!root.current?.contains(event.target as Node)) setOpen(false);
    };
    const escape = (event: KeyboardEvent) => {
      if (event.key !== 'Escape') return;
      event.preventDefault(); event.stopPropagation();
      setOpen(false); trigger.current?.focus();
    };
    document.addEventListener('pointerdown', outside);
    root.current?.addEventListener('keydown', escape);
    const node = root.current;
    return () => { document.removeEventListener('pointerdown', outside); node?.removeEventListener('keydown', escape); };
  }, [open]);
  return <div ref={root} className={`realm-responsive-menu ${className}${open ? ' is-open' : ''}`} onBlur={event => {
    if (!event.currentTarget.contains(event.relatedTarget as Node | null)) setOpen(false);
  }}>
    <button ref={trigger} type="button" className="realm-menu-trigger" aria-expanded={open} aria-controls={id} onClick={() => setOpen(value => !value)}>{label} <span aria-hidden="true">⌄</span></button>
    <div id={id} className="realm-menu-panel" onClick={event => {
      if (open && closeOnSelect && (event.target as HTMLElement).closest('a, button')) {
        setOpen(false); trigger.current?.focus();
      }
    }}>{children}</div>
  </div>;
}
