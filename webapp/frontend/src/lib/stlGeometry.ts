import type { BufferGeometry } from "three";
import { STLLoader } from "three/examples/jsm/loaders/STLLoader.js";

export interface GeomStats {
  sx: number;
  sy: number;
  sz: number; // bounding-box size
  bcx: number;
  bcy: number;
  bcz: number; // bounding-box center (world coords)
  comx: number;
  comy: number;
  comz: number; // center of mass, relative to bbox center
  vol: number; // volume, model units^3 (typically in^3)
  area: number;
  tris: number;
}

const EMPTY_STATS: GeomStats = {
  sx: 0, sy: 0, sz: 0, bcx: 0, bcy: 0, bcz: 0, comx: 0, comy: 0, comz: 0, vol: 0, area: 0, tris: 0,
};

export function emptyGeomStats(): GeomStats {
  return { ...EMPTY_STATS };
}

export const MM_PER_INCH = 25.4;

// A mold base is never this big in inches (16 ft across). Mirrors the macro's
// own MAX_SANE_MOLD_DIM_IN guard, which uses the same trick to spot a
// double-converted dimension.
export const MAX_SANE_MOLD_DIM_IN = 200;

/**
 * Force an STL's geometry into inches, in place.
 *
 * STL carries no unit tag, and SolidWorks writes whatever the document's export
 * unit happens to be -- on this shop's machine that was millimetres, so an
 * 18.000 in plate arrived as 457.20 and its volume was off by 25.4^3 (~16,387x),
 * which is how a clamping plate read 1.7 million lb.
 *
 * Two modes:
 *
 *  - `expectedMaxIn` given (PREFERRED): the caller knows the part's true largest
 *    dimension from the CAD export, which is always in inches. Compare and scale
 *    on the actual ratio. This is deterministic and works at any size.
 *
 *  - No hint: fall back to a magnitude test. This only catches parts bigger than
 *    MAX_SANE_MOLD_DIM_IN, so a 0.5 in dowel written as 12.7 mm slips through
 *    reading 12.7 in. Fine for plates and whole bases, unreliable for small
 *    hardware -- which is exactly why the hint exists.
 *
 * Scaling the geometry rather than the label means every consumer (camera
 * framing, computeGeomStats, the matching maths) agrees on one unit.
 * Idempotent: after scaling the extent is in inches, so a second call is a
 * no-op. That matters because three's useLoader caches geometry and hands the
 * same object to several components.
 *
 * Returns the scale factor applied (1 when the file was already in inches).
 */
export function normalizeGeometryToInches(
  geometry: BufferGeometry,
  expectedMaxIn?: number,
): number {
  geometry.computeBoundingBox();
  const b = geometry.boundingBox;
  if (!b) return 1;

  const maxExtent = Math.max(b.max.x - b.min.x, b.max.y - b.min.y, b.max.z - b.min.z);
  if (!Number.isFinite(maxExtent) || maxExtent <= 0) return 1;

  // Deterministic path: the CAD export tells us the truth.
  if (typeof expectedMaxIn === "number" && expectedMaxIn > 0.01) {
    const ratio = maxExtent / expectedMaxIn;
    // Only two plausible unit systems in this pipeline. Accept a generous band
    // because the STL is a tessellation of the solid, not the solid itself.
    if (ratio > 10) {
      const s = 1 / MM_PER_INCH;
      geometry.scale(s, s, s);
      geometry.computeBoundingBox();
      return s;
    }
    return 1;
  }

  // Fallback: magnitude only.
  if (maxExtent <= MAX_SANE_MOLD_DIM_IN) return 1;

  const s = 1 / MM_PER_INCH;
  geometry.scale(s, s, s);
  geometry.computeBoundingBox();
  return s;
}

// Ported 1:1 from the shop's Match Studio overlay viewer (geomStats) so
// volume / center-of-mass numbers here always agree with the shop's other
// tools. Centroid uses the signed-tetrahedron method, so it is the true
// center of mass of the solid -- not just the bounding-box center.
export function computeGeomStats(geometry: BufferGeometry): GeomStats {
  const posAttr = geometry.attributes.position;
  const arr = posAttr.array as ArrayLike<number>;
  const cnt = posAttr.count;
  const idxAttr = geometry.index;
  const idx = idxAttr ? (idxAttr.array as ArrayLike<number>) : null;
  const n = idx ? idx.length : cnt;

  let minx = Infinity, miny = Infinity, minz = Infinity;
  let maxx = -Infinity, maxy = -Infinity, maxz = -Infinity;
  for (let v = 0; v < cnt; v++) {
    const x = arr[v * 3], y = arr[v * 3 + 1], z = arr[v * 3 + 2];
    if (x < minx) minx = x;
    if (y < miny) miny = y;
    if (z < minz) minz = z;
    if (x > maxx) maxx = x;
    if (y > maxy) maxy = y;
    if (z > maxz) maxz = z;
  }

  let sv = 0, cx = 0, cy = 0, cz = 0, area = 0;
  for (let t = 0; t < n; t += 3) {
    const ia = (idx ? idx[t] : t) * 3;
    const ib = (idx ? idx[t + 1] : t + 1) * 3;
    const ic = (idx ? idx[t + 2] : t + 2) * 3;
    const ax = arr[ia], ay = arr[ia + 1], az = arr[ia + 2];
    const bx = arr[ib], by = arr[ib + 1], bz = arr[ib + 2];
    const cxx = arr[ic], cyy = arr[ic + 1], czz = arr[ic + 2];
    const vv = (-cxx * by * az + bx * cyy * az + cxx * ay * bz - ax * cyy * bz - bx * ay * czz + ax * by * czz) / 6;
    sv += vv;
    cx += (vv * (ax + bx + cxx)) / 4;
    cy += (vv * (ay + by + cyy)) / 4;
    cz += (vv * (az + bz + czz)) / 4;
    const ux = bx - ax, uy = by - ay, uz = bz - az;
    const wx = cxx - ax, wy = cyy - ay, wz = czz - az;
    const px = uy * wz - uz * wy, py = uz * wx - ux * wz, pz = ux * wy - uy * wx;
    area += 0.5 * Math.sqrt(px * px + py * py + pz * pz);
  }

  const bcx = (minx + maxx) / 2, bcy = (miny + maxy) / 2, bcz = (minz + maxz) / 2;
  const comx = Math.abs(sv) > 1e-9 ? cx / sv : bcx;
  const comy = Math.abs(sv) > 1e-9 ? cy / sv : bcy;
  const comz = Math.abs(sv) > 1e-9 ? cz / sv : bcz;

  return {
    sx: maxx - minx,
    sy: maxy - miny,
    sz: maxz - minz,
    bcx, bcy, bcz,
    comx: comx - bcx,
    comy: comy - bcy,
    comz: comz - bcz,
    vol: Math.abs(sv),
    area,
    tris: Math.round(n / 3),
  };
}

export interface MatchResult {
  pct: number;
  sizeDiffPct: number;
  massDiffPct: number;
  comDistIn: number;
}

// Same weighting the shop's Job Matching uses: 0.40 size + 0.25 mass (volume
// proxy) + 0.35 center of mass.
export function matchPercent(a: GeomStats, b: GeomStats): MatchResult {
  const da = [a.sx, a.sy, a.sz].sort((x, y) => x - y);
  const db = [b.sx, b.sy, b.sz].sort((x, y) => x - y);
  let sd = 0;
  for (let i = 0; i < 3; i++) {
    const m = Math.max(da[i], db[i]) || 1;
    sd += Math.abs(da[i] - db[i]) / m;
  }
  sd /= 3;

  const massd = Math.abs(a.vol - b.vol) / (Math.max(a.vol, b.vol) || 1);

  const dx = a.comx - b.comx, dy = a.comy - b.comy, dz = a.comz - b.comz;
  const dist = Math.sqrt(dx * dx + dy * dy + dz * dz);
  const scale = (Math.max(a.sx, a.sy, a.sz) + Math.max(b.sx, b.sy, b.sz)) / 2 || 1;
  const cd = Math.min(1, dist / scale);

  const diff = 0.4 * sd + 0.25 * massd + 0.35 * cd;
  return {
    pct: Math.max(0, Math.min(100, (1 - diff) * 100)),
    sizeDiffPct: sd * 100,
    massDiffPct: massd * 100,
    comDistIn: dist,
  };
}

export function matchTone(pct: number): "success" | "warning" | "danger" {
  if (pct >= 92) return "success";
  if (pct >= 80) return "warning";
  return "danger";
}

// Rough material densities (lb / in^3) for a plain-language mass estimate.
// Volume is exact (computed from the mesh); mass is only ever an estimate,
// since STL files carry no material data. Falls back to tool steel.
const DENSITY_LOOKUP: { match: RegExp; density: number }[] = [
  { match: /alum/i, density: 0.098 },
  { match: /stainless/i, density: 0.29 },
  { match: /brass/i, density: 0.307 },
  { match: /copper/i, density: 0.323 },
  { match: /(tool steel|p20|h13|s7|a2|d2|steel)/i, density: 0.284 },
];
const DEFAULT_STEEL_DENSITY = 0.284;

export function densityFor(material?: string): number {
  if (material) {
    for (const { match, density } of DENSITY_LOOKUP) {
      if (match.test(material)) return density;
    }
  }
  return DEFAULT_STEEL_DENSITY;
}

export function estimateMassLb(volumeIn3: number, density: number): number {
  return volumeIn3 * density;
}

// Off-screen STL parse (no Canvas/WebGL context needed) -- used to batch-analyze
// every file in a job without rendering them.
const _offscreenLoader = new STLLoader();

export async function loadStlGeometry(
  url: string,
  expectedMaxIn?: number,
): Promise<BufferGeometry> {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch ${url}: ${res.status} ${res.statusText}`);
  const buf = await res.arrayBuffer();
  const geometry = _offscreenLoader.parse(buf);
  normalizeGeometryToInches(geometry, expectedMaxIn);
  return geometry;
}
