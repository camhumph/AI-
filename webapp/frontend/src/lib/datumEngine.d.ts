/**
 * Hand-written types for datum-engine.js.
 *
 * The engine is framework-free plain JS with a single export block at the
 * bottom. This shim exists for two reasons:
 *
 *  1. Without it the import needs a `@ts-expect-error`, and that directive
 *     inverts into a hard TS2578 ("unused directive") the moment anyone sets
 *     allowJs — a build failure in the *good* configuration.
 *
 *  2. It makes the engine's surface actually type-checked. `pl.ops` was `any`,
 *     which is why `qty` being a display STRING ("12 holes") rather than a
 *     number went unnoticed and silently dropped every quantity from the UI.
 *
 * EVERYTHING HERE IS MILLIMETRES. Volumes are mm^3, rho is g/cm^3, vc is
 * m/min, fz is mm/tooth. The adapter converts at the boundary.
 */

export interface DatumMaterial {
  id: string;
  name: string;
  cls: string;
  /** Density, g/cm^3. */
  rho: number;
  /** Stock cost, $/kg. */
  kg: number;
  /** Surface speed, m/min. */
  vc: number;
  /** Feed per tooth, mm. */
  fz: number;
  /** Machinability factor; 1.0 = 6061. */
  k: number;
  /** Cutter cost, $/litre removed. */
  tc: number;
}

export interface DatumTolerance {
  v: number;
  label: string;
  name: string;
  fin: number;
  passes: number;
  /** Inspection minutes. */
  insp: number;
  scrap: number;
  gauge: string;
}

export interface DatumAnalysis {
  nT: number;
  nV: number;
  /** Solid volume, mm^3, by signed tetrahedron sum. */
  vol: number;
  /** Surface area, mm^2. */
  area: number;
  bb: number[];
  reg: Int32Array | number[];
  regions: unknown[];
  pos: Float32Array;
  nrm: Float32Array;
  cen: Float32Array;
}

export interface DatumFeatures {
  /** Bounding-box extents [x, y, z], mm. */
  ext: number[];
  stock: number[];
  stockVol: number;
  holes: Array<{ d: number; depth: number; blind: boolean; purpose?: string; tap?: unknown }>;
  bosses: unknown[];
  fillets: unknown[];
  pockets: Array<{ depth: number; area: number; cbore?: boolean }>;
  faces: unknown[];
  chamfers: unknown[];
  /** Total chamfer edge length, mm. */
  chamferLen: number;
  /** Distinct machined directions with >= 2 features. */
  setups: number;
  frees: unknown[];
  angled: unknown[];
  curved: unknown[];
  dirs: unknown;
  /** Smallest internal corner radius, mm, or null if none found. */
  minInternalR: number | null;
  diag: unknown;
}

export interface DatumStock {
  /** [x, y, z] mm. */
  stock: number[];
  /** Axis indices that need facing. */
  faces: number[];
  mode: string;
  /** True when the supplied stock is SMALLER than the measured part. */
  short: boolean;
  plateT?: number;
}

export interface DatumOp {
  id: string;
  label: string;
  tool: string;
  detail: string;
  /** A DISPLAY STRING, not a count -- e.g. "12 holes", "466 x 403 mm". */
  qty: string;
  /** Minutes. */
  min: number;
  heat: number;
  regs?: number[];
}

export interface DatumPlan {
  ops: DatumOp[];
  /** Cycle minutes for one part. Raw sum of op minutes, NOT derated. */
  cycle: number;
  /** Setup minutes, off cycle. */
  setupMin: number;
  setups: number;
  toolChanges: number;
  /** Stock volume, mm^3. */
  stockVol: number;
  removeVol: number;
  openVol: number;
  pocketVol: number;
  holeMin: number;
  faceMin: number;
  facing: unknown;
  schedule: unknown[];
  regHeat: Float32Array;
  toolCost: number;
  cornerLimited: boolean;
  kfac: number;
  Dr: number;
  Df: number;
  reachPct: number;
  dirCount: number;
  st: DatumStock;
}

export interface DatumCost {
  lines: Array<{ id: string; label: string; sub: string; heat: number; v: number }>;
  subtotal: number;
  risk: number;
  /** One-part total, all lines plus risk. */
  total: number;
  stockKg: number;
  hours: { cycle: number; setup: number; total: number };
}

export function parseSTL(buf: ArrayBuffer): Float32Array;
export function analyze(pos: Float32Array): DatumAnalysis;
export function extract(an: DatumAnalysis): DatumFeatures;
export function makeStock(
  ext: number[],
  mode: "block" | "plate" | "custom",
  allow: number,
  custom?: number[],
  imperial?: boolean,
): DatumStock;
export function plan(
  an: DatumAnalysis,
  fx: DatumFeatures,
  mat: DatumMaterial,
  tol: DatumTolerance,
  rate: number,
  opts?: { st?: DatumStock; tap?: boolean },
): DatumPlan;
export function cost(
  pl: DatumPlan,
  an: DatumAnalysis,
  fx: DatumFeatures,
  mat: DatumMaterial,
  tol: DatumTolerance,
  rate: number,
): DatumCost;
export function materialByName(name?: string): DatumMaterial;
export function fmtMin(min: number): string;
export function money(v: number): string;

/** Present and true only in the placeholder stub, absent in the real engine. */
export const IS_STUB: boolean | undefined;

export const MATERIALS: DatumMaterial[];
export const TOLS: DatumTolerance[];
export const ALLOW: number;
/** Declared in the engine but never read by plan(). Do not present as applied. */
export const DUTY: number;
export const RPM_MAX: number;
