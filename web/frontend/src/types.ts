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

export type AppTab = 'read' | 'updates' | 'stats' | 'admin';

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
export type ChangeDiffRow = {
  type: 'context' | 'added' | 'removed' | 'gap';
  old_no?: number | null;
  new_no?: number | null;
  text?: string;
};
export type ChangeDetail = {
  from_ref: string;
  to: string;
  kind: ChangeKind;
  path: string;
  old_path?: string | null;
  title: string;
  old_exists: boolean;
  new_exists: boolean;
  old_encoding: string;
  new_encoding: string;
  old_line_count: number;
  new_line_count: number;
  additions: number;
  deletions: number;
  truncated: boolean;
  rows: ChangeDiffRow[];
};
export type ChangeStats = { added: number; modified: number; deleted: number; renamed: number; total: number };
export type GitChanges = {
  from_ref: string;
  to: string;
  public_only?: boolean;
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

export type NormalizationSignature = {
  signer: string;
  note?: string;
  signed_at: string;
};

export type NormalizationReview = {
  id: string;
  resource_path: string;
  title: string;
  original_content: string;
  normalized_content: string;
  note?: string;
  created_at: string;
  updated_at: string;
  required_signatures: number;
  signatures: NormalizationSignature[];
  status: 'pending' | 'approved';
};

export type NormalizationReviewSummary = {
  id: string;
  resource_path: string;
  title: string;
  status: 'pending' | 'approved';
  signature_count: number;
  required_signatures: number;
  created_at: string;
  updated_at: string;
};

export type SearchFilters = {
  q: string;
  category: string;
  kinds: string[];
  sides: string[];
  authors: string[];
};

export type SessionStatsOverview = {
  session_count: number;
  participant_count: number;
  total_game_hours: number;
  total_host_hours: number;
};

export type SessionStatsPlayerSort = 'hours' | 'games' | 'hosts' | 'name';

export type SessionStatsPlayer = {
  id: string | number;
  name: string;
  qq?: string | null;
  game_count: number;
  game_hours: number;
  reincarnation_count: number;
  host_count: number;
  host_hours: number;
};

export type SessionStatsPlayersResponse = {
  items: SessionStatsPlayer[];
  count: number;
  month: string;
};

export type SessionStatsSession = {
  id: string | number;
  title: string;
  duration_hours: number;
  source_filename: string;
  kp_name?: string | null;
  kp_qq?: string | null;
  pl_count: number;
  created_at?: string | null;
  confidence?: number | null;
  model_name?: string | null;
};

export type SessionStatsSessionsResponse = {
  items: SessionStatsSession[];
  count: number;
  month: string;
};

export type SessionStatsImportFileResult = {
  filename?: string;
  source_filename?: string;
  title?: string;
  success?: boolean;
  ok?: boolean;
  error?: string;
  reason?: string;
  detail?: string;
  message?: string;
  session_id?: string | number;
  [key: string]: unknown;
};

export type SessionStatsImportResponse = {
  success_count: number;
  failure_count: number;
  items: SessionStatsImportFileResult[];
};

export type SessionStatsImportJob = {
  job_id: string;
  status: 'running' | 'completed' | 'failed';
  month: string;
  total_count: number;
  processed_count: number;
  success_count: number;
  failure_count: number;
  current_filename?: string;
  items: SessionStatsImportFileResult[];
  error?: string;
  created_at?: string;
  updated_at?: string;
};

export type SessionStatsErrorItem = Record<string, unknown>;
export type SessionStatsErrorsResponse = {
  items: SessionStatsErrorItem[];
  count: number;
  month: string;
};
