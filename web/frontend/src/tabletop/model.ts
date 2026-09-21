export type Kind = 'token' | 'note' | 'circle' | 'rect' | 'cone' | 'line' | 'polygon';
export type Point = { x: number; y: number };
export type Piece = Point & {
  id: string; kind: Kind; name: string; width: number; height: number; radius: number;
  angle: number; rotation: number; points: [number, number][]; color: string; asset: string;
  owners: string[]; visibility: 'all' | 'gm' | 'selected'; viewers?: string[];
  locked: boolean; enabled: boolean; follow: string; elevation: string; footprint: string;
  note: string; gm_note?: string; counters: {name: string; value: string; maximum: string}[];
  statuses: string[]; source: string; v: number; editable?: boolean;
};
export type Scene = {
  id: string; name: string; grid: boolean; snap: boolean; grid_size: number; background: string;
  map_width: number; map_height: number; round: number; turn: string; order: string[];
  pieces: Record<string, Piece>; v: number;
};
export type Member = { id: string; name: string; role: 'gm' | 'player' | 'spectator'; online: boolean };
export type Snapshot = {
  id: string; name: string; revision: number; me: string; role: Member['role']; scene: Scene;
  scenes: {id: string; name: string}[]; members: Member[]; canUndo: boolean;
  ping?: Point & { scene: string; expires: number } | null;
};
export type Edit = { id: string; expected: number; value: Omit<Piece, 'editable'> | null };
export type Action = { kind: string; scene?: string; expected?: number; edits?: Edit[]; data?: Record<string, unknown> };
export const makeId = () => crypto.randomUUID();
export const rawPiece = ({editable: _editable, ...p}: Piece) => p;
export const newPiece = (kind: Kind, position: Point, me: string): Piece => ({
  id: makeId(), kind, name: kind === 'token' ? '新棋子' : kind === 'note' ? '文字标记' : '技能范围',
  ...position, width: 2, height: 2, radius: 5, angle: 60, rotation: 0, points: [],
  color: '#6dcfc1', asset: '', owners: [me], visibility: 'all', viewers: [], locked: false,
  enabled: true, follow: '', elevation: '', footprint: '', note: '', gm_note: '',
  counters: [], statuses: [], source: '', v: 0,
});
export const distance = (a: Point, b: Point) => Math.hypot(a.x-b.x, a.y-b.y);
export const pathLength = (points: Point[]) => points.slice(1).reduce((sum,p,i) => sum+distance(points[i],p),0);
export const positionOf = (p: Piece, pieces: Record<string, Piece>): Point => p.follow && pieces[p.follow]
  ? {x: p.x + pieces[p.follow].x, y: p.y + pieces[p.follow].y} : p;
export const snapPoint = (p: Point, scene: Scene): Point => scene.snap
  ? {x: Math.round(p.x/scene.grid_size)*scene.grid_size, y: Math.round(p.y/scene.grid_size)*scene.grid_size} : p;
export const sceneEdit = (scene: Scene, data: Record<string, unknown>): Action => ({kind:'scene',scene:scene.id,expected:scene.v,data});
