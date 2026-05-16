import React from 'react';

export function StatusBadge({ ok, children, warn = false }: { ok: boolean; children: React.ReactNode; warn?: boolean }) {
  return <span className={ok ? 'status-badge ok' : warn ? 'status-badge warn' : 'status-badge bad'}>{children}</span>;
}
