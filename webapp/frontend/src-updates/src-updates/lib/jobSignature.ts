import { computeGeomStats, loadStlGeometry, type GeomStats } from "./stlGeometry";
import { classifyComponentFilename, type ComponentKind } from "./componentKind";

// The 6 component roles the shop's job-signature system tracks, in the exact
// order/list used by JOB_SIGNATURE_COMPONENTS in the Elgin backend.
export const SIGNATURE_KINDS: ComponentKind[] = ["TCP", "BCP", "ID HOLDER", "OD HOLDER", "ID POT", "OD POT"];

// Same defaults as ELGIN_SIGNATURE_SIZE_TOLERANCE / _MASS_TOLERANCE / _CENTER_TOLERANCE.
export const SIGNATURE_SIZE_TOLERANCE = 0.08;
export const SIGNATURE_MASS_TOLERANCE = 0.15;
export const SIGNATURE_CENTER_TOLERANCE = 0.15;

export function relativeDiff(a: number, b: number, floor = 0.001): number {
  const denom = Math.max((Math.abs(a) + Math.abs(b)) / 2, floor);
  return Math.abs(a - b) / denom;
}

// Forgive differences within tolerance; scale the remaining gap above the band.
export function tolerantDiff(rawDiff: number, tolerance: number): number {
  const raw = Math.max(0, rawDiff || 0);
  const tol = Math.max(0, tolerance || 0);
  if (tol <= 0) return raw;
  if (raw <= tol) return 0;
  return (raw - tol) / Math.max(1 - tol, 0.001);
}

export interface ComponentMatchResult {
  similarity: number;
  sizeDiffPct: number;
  massDiffPct: number;
  centerDistanceIn: number;
}

// Same formula as compare_signature_component, adapted for STL-derived
// signatures: dimensions are sorted smallest-to-largest before comparing
// (rather than compared as labeled length/width/thickness), since STL
// bounding boxes don't carry axis-labeled CAD metadata the way the shop's
// XT_Export_Job_Signature CSVs do. "Mass" uses raw mesh volume (in^3) --
// mathematically equivalent to comparing true mass as long as both sides
// would use the same material density, since a shared scale factor cancels
// out of a relative-difference ratio.
export function compareComponentSignature(current: GeomStats, candidate: GeomStats): ComponentMatchResult {
  const da = [current.sx, current.sy, current.sz].sort((a, b) => a - b);
  const db = [candidate.sx, candidate.sy, candidate.sz].sort((a, b) => a - b);

  let sizeDiff = 0;
  for (let i = 0; i < 3; i++) {
    sizeDiff += tolerantDiff(relativeDiff(da[i], db[i], 0.25), SIGNATURE_SIZE_TOLERANCE);
  }
  sizeDiff /= 3;

  const massDiff = tolerantDiff(relativeDiff(current.vol, candidate.vol, 0.1), SIGNATURE_MASS_TOLERANCE);

  const centerDistance = Math.sqrt(
    (current.comx - candidate.comx) ** 2 + (current.comy - candidate.comy) ** 2 + (current.comz - candidate.comz) ** 2
  );
  const curDiag = Math.sqrt(da[0] ** 2 + da[1] ** 2 + da[2] ** 2);
  const candDiag = Math.sqrt(db[0] ** 2 + db[1] ** 2 + db[2] ** 2);
  const centerDiff = tolerantDiff(centerDistance / Math.max((curDiag + candDiag) / 2, 1), SIGNATURE_CENTER_TOLERANCE);

  // Same weighting as the shop's Job Matching: 0.40 size + 0.25 mass + 0.35 center.
  const diffScore = 0.4 * sizeDiff + 0.25 * massDiff + 0.35 * centerDiff;
  const similarity = Math.max(0, Math.min(100, 100 * (1 - diffScore)));

  return {
    similarity: Math.round(similarity * 100) / 100,
    sizeDiffPct: Math.round(sizeDiff * 100 * 1000) / 1000,
    massDiffPct: Math.round(massDiff * 100 * 1000) / 1000,
    centerDistanceIn: Math.round(centerDistance * 10000) / 10000,
  };
}

export interface JobSignature {
  jobId: string;
  displayName: string;
  components: Map<ComponentKind, { model: { name: string; url: string }; stats: GeomStats }>;
}

function isStlFile(name: string): boolean {
  return name.toLowerCase().endsWith(".stl");
}

// Builds a job's signature by fetching only the (up to 6) STL files that
// matter -- one per SIGNATURE_KINDS role -- skipping everything else in the
// job's models list. Shared by the Matching page and the Comparison tab so
// both build signatures the exact same way.
export async function buildJobSignature(
  jobId: string,
  displayName: string,
  models: { name: string; url: string }[]
): Promise<JobSignature> {
  const components = new Map<ComponentKind, { model: { name: string; url: string }; stats: GeomStats }>();
  for (const kind of SIGNATURE_KINDS) {
    const model = models.find((m) => isStlFile(m.name) && classifyComponentFilename(m.name) === kind);
    if (!model) continue;
    try {
      const geometry = await loadStlGeometry(model.url);
      components.set(kind, { model, stats: computeGeomStats(geometry) });
    } catch {
      // skip this component if its file fails to load
    }
  }
  return { jobId, displayName, components };
}

export interface JobMatchComponent extends ComponentMatchResult {
  kind: ComponentKind;
}

export interface JobMatchResult {
  jobId: string;
  displayName: string;
  score: number;
  avgComponentSimilarity: number;
  matchedComponents: number;
  coveragePct: number;
  missingCurrent: ComponentKind[];
  missingCandidate: ComponentKind[];
  components: JobMatchComponent[];
}

// Port of compare_job_signature_to_previous's per-candidate scoring: average
// similarity across every matched component, then scale by how much of the
// 6-part signature was actually covered (0.75 + 0.25 * coverage), so a job
// that matches on all 6 parts outranks one that only had 2 parts to compare.
export function compareJobSignatures(current: JobSignature, candidate: JobSignature): JobMatchResult | null {
  const components: JobMatchComponent[] = [];
  const missingCurrent: ComponentKind[] = [];
  const missingCandidate: ComponentKind[] = [];

  for (const kind of SIGNATURE_KINDS) {
    const cur = current.components.get(kind);
    const cand = candidate.components.get(kind);
    if (!cur) {
      missingCurrent.push(kind);
      continue;
    }
    if (!cand) {
      missingCandidate.push(kind);
      continue;
    }
    components.push({ kind, ...compareComponentSignature(cur.stats, cand.stats) });
  }

  if (components.length === 0) return null;

  const avgSimilarity = components.reduce((s, c) => s + c.similarity, 0) / components.length;
  const coverage = components.length / SIGNATURE_KINDS.length;
  const finalScore = avgSimilarity * (0.75 + 0.25 * coverage);

  return {
    jobId: candidate.jobId,
    displayName: candidate.displayName,
    score: Math.round(finalScore * 100) / 100,
    avgComponentSimilarity: Math.round(avgSimilarity * 100) / 100,
    matchedComponents: components.length,
    coveragePct: Math.round(coverage * 1000) / 10,
    missingCurrent,
    missingCandidate,
    components: components.sort((a, b) => a.kind.localeCompare(b.kind)),
  };
}
