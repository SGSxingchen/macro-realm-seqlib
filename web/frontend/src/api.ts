function parseErrorDetail(raw: string, status: number): string {
  const trimmed = raw.trim();
  if (!trimmed) return `HTTP ${status}`;
  try {
    const parsed = JSON.parse(trimmed) as { detail?: unknown; message?: unknown; error?: unknown };
    if (typeof parsed.detail === 'string' && parsed.detail.trim()) return parsed.detail.trim();
    if (typeof parsed.message === 'string' && parsed.message.trim()) return parsed.message.trim();
    if (typeof parsed.error === 'string' && parsed.error.trim()) return parsed.error.trim();
  } catch {
    /* keep raw text */
  }
  return trimmed;
}

export const api = async <T,>(url: string, init?: RequestInit): Promise<T> => {
  const hasBody = init?.body !== undefined && init?.body !== null;
  const headers = init?.body instanceof FormData
    ? init.headers
    : hasBody
      ? { 'Content-Type': 'application/json', ...(init?.headers || {}) }
      : { ...(init?.headers || {}) };
  const res = await fetch(url, { credentials: 'include', ...init, headers });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(parseErrorDetail(text, res.status));
  }
  return res.json();
};

export const routePath = (path: string) => encodeURIComponent(path).replaceAll('%2F', '/');

/** 把多值参数序列化为可重复 query。空值跳过。 */
export const buildQuery = (params: Record<string, string | string[] | number | boolean | undefined | null>): string => {
  const usp = new URLSearchParams();
  for (const [k, v] of Object.entries(params)) {
    if (v === undefined || v === null || v === '') continue;
    if (Array.isArray(v)) v.forEach(x => x !== '' && x !== undefined && usp.append(k, String(x)));
    else usp.set(k, String(v));
  }
  const s = usp.toString();
  return s ? `?${s}` : '';
};
