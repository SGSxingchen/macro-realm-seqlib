export type Resource = {
  path: string;
  filename: string;
  title: string;
  root: string;
  category: string;
  mtime: number;
  size: number;
  side?: string;
  top_kind?: string;
  authors?: string[];
  score?: number;
  snippet?: string;
};

export type Detail = Resource & { content: string; encoding: string };

export type TreeNode = { name: string; path: string; count: number; children: TreeNode[] };

export type FacetItem = { name: string; count: number };
export type Facets = { kinds: FacetItem[]; sides: FacetItem[]; authors: FacetItem[] };

export type ResourceListResponse = {
  items: Resource[];
  count: number;
  total: number;
  limit: number;
  offset: number;
  tokens: string[];
  facets: Facets;
  engine?: { pinyin: boolean; opencc: boolean };
};

export type AdminState = { admin_configured: boolean; authenticated: boolean };

export type ChangeKind = 'added' | 'modified' | 'deleted' | 'renamed';
export type ChangeItem = {
  title: string;
  path: string;
  old_path?: string | null;
  category: string;
  root: string;
  size?: number | null;
  exists: boolean;
  score?: string | null;
};
export type ChangeStats = { added: number; modified: number; deleted: number; renamed: number; total: number };
export type GitChanges = {
  from_ref: string;
  to: string;
  stats: ChangeStats;
  readable: Record<ChangeKind, ChangeItem[]>;
  text: string;
  markdown: string;
  summary: unknown;
  raw: unknown;
};

export type CmdResult = { cmd: string[]; returncode: number; stdout: string; stderr: string; seconds?: number };
export type GitInfo = {
  latest_tag?: string | null;
  head_short?: string | null;
  head_full?: string | null;
  head_tags?: string[];
  branch_name?: string | null;
  status_short?: string;
  status_branch?: string;
  is_dirty?: boolean;
  tracking?: string | null;
  ahead_behind?: { ahead: number; behind: number; tracking: string } | null;
  remote_main?: string | null;
  admin_configured?: boolean;
  wiki_configured?: boolean;
  branch?: CmdResult;
  head?: CmdResult | string;
  status?: CmdResult;
};

export type PublishResult = { ok: boolean; version: string; steps: CmdResult[] };
export type AdminTab = 'overview' | 'edit' | 'package' | 'changes' | 'publish' | 'wiki';
export type HonorCategory = { category: string; count: number; next_number: number; next_prefix: string };
export type HonorCategoriesResponse = { items: HonorCategory[]; suggested_category?: string };
export type PackageFile = {
  path: string;
  filename: string;
  title: string;
  size: number;
  mtime: number;
  editable: boolean;
  extension: string;
  content?: string | null;
  encoding?: string | null;
};

export type SearchFilters = {
  q: string;
  category: string;
  kinds: string[];
  sides: string[];
  authors: string[];
};
