export const api = async <T,>(url: string, init?: RequestInit): Promise<T> => {
  const headers = init?.body instanceof FormData
    ? init.headers
    : { 'Content-Type': 'application/json', ...(init?.headers || {}) };
  const res = await fetch(url, { credentials: 'include', ...init, headers });
  if (!res.ok) {
    const text = await res.text();
    throw new Error(text || `HTTP ${res.status}`);
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
