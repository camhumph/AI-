import { computeGeomStats, loadStlGeometry, type GeomStats } from "./stlGeometry";
import {
  classifyComponentFilename,
  classifyQuoteComponentName,
  MACHINED_PART_KINDS,
  type ComponentKind,
} from "./componentKind";

// The 6 component roles the shop's job-signature system tracks, in the exact
// order/list used by JOB_SIGNATURE_COMPONENTS in the Elgin backend. This is the
// BMS pot-block set specifically.
export const SIGNATURE_KINDS: ComponentKind[] = ["TCP", "BCP", "ID HOLDER", "OD HOLDER", "ID POT", "OD POT"];

/**
 * The kinds a signature may be built from.
 *
 * Every real part, not just the BMS six. A signature is a Map, so it has always
 * been variable-size -- it was only ever the LOOKUP that was pinned to the pot
 * block, which meant a standard or PCS base built an empty signature and could
 * never match anything on the /matching page.
 *
 * MACHINED_PART_KINDS is in stack order and starts with the standard plates, so
 * a pot-block job still resolves exactly TCP, BCP, ID HOLDER, OD HOLDER, ID POT,
 * OD POT, in that relative order -- BMS matching is unchanged.
 */
const SIGNATURE_CANDIDATE_KINDS: ComponentKind[] = MACHINED_PART_KINDS;

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
// XT_Export_Job_Signature CSVs do.
//
// The mass term prefers real measured lb from SolidWorks (CreateMassProperty,
// carried on the part rows as MassLb) and only falls back to mesh volume when
// a part has no measured mass. Volume is a fine proxy when both sides are the
// same material -- a shared density cancels out of a relative-difference ratio
// -- but it silently lies the moment two jobs use different steel.
export function compareComponentSignature(
  current: GeomStats,
  candidate: GeomStats,
  currentMassLb?: number,
  candidateMassLb?: number
): ComponentMatchResult {
  const da = [current.sx, current.sy, current.sz].sort((a, b) => a - b);
  const db = [candidate.sx, candidate.sy, candidate.sz].sort((a, b) => a - b);

  let sizeDiff = 0;
  for (let i = 0; i < 3; i++) {
    sizeDiff += tolerantDiff(relativeDiff(da[i], db[i], 0.25), SIGNATURE_SIZE_TOLERANCE);
  }
  sizeDiff /= 3;

  const haveMeasuredMass =
    typeof currentMassLb === "number" && currentMassLb > 0 &&
    typeof candidateMassLb === "number" && candidateMassLb > 0;
  const massA = haveMeasuredMass ? (currentMassLb as number) : current.vol;
  const massB = haveMeasuredMass ? (candidateMassLb as number) : candidate.vol;

  const massDiff = tolerantDiff(relativeDiff(massA, massB, 0.1), SIGNATURE_MASS_TOLERANCE);

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

export interface SignatureComponent {
  model: { name: string; url: string };
  stats: GeomStats;
  /** Measured lb from SolidWorks, when the job's part rows carry it. */
  massLb?: number;
}

export interface JobSignature {
  jobId: string;
  displayName: string;
  components: Map<ComponentKind, SignatureComponent>;
  /** The whole-base STL, so two complete mold bases can be overlaid. */
  base?: SignatureComponent;
}

export interface SignatureModel {
  name: string;
  url: string;
  size?: number;
}

function isStlFile(name: string): boolean {
  return name.toLowerCase().endsWith(".stl");
}

/**
 * The whole-base STL. Module6121 writes it as `{JobBaseName}.stl` with no plate
 * label, so it classifies as OTHER rather than FULL ASSEMBLY. Prefer an explicit
 * FULL ASSEMBLY, then fall back to the largest STL that isn't one of the six
 * signature parts -- the merged base is always by far the biggest file.
 */
export function findBaseModel(models: SignatureModel[]): SignatureModel | null {
  const stls = models.filter((m) => isStlFile(m.name));
  const explicit = stls.find((m) => classifyComponentFilename(m.name) === "FULL ASSEMBLY");
  if (explicit) return explicit;

  // Exclude EVERY known part kind, not just the pot-block six. On a standard
  // base the six were all absent, so all eleven plates counted as "leftovers"
  // and the largest plate could be picked as the whole base.
  const leftovers = stls.filter((m) => !SIGNATURE_CANDIDATE_KINDS.includes(classifyComponentFilename(m.name)));
  if (leftovers.length === 0) return null;
  return leftovers.reduce((biggest, m) => ((m.size ?? 0) > (biggest.size ?? 0) ? m : biggest));
}

// Builds a job's signature by fetching only the STL files that matter -- one
// per part role present on the job -- plus the whole-base STL for the overlay,
// skipping everything else in the job's models list. Six files on a pot block,
// eight to eleven on a standard or PCS base.
//
// massByKind carries measured SolidWorks lb per component kind; pass it and the
// mass term of the score becomes a real weight comparison instead of a volume
// proxy. Optional so callers without part rows still work.
export async function buildJobSignature(
  jobId: string,
  displayName: string,
  models: SignatureModel[],
  massByKind?: Map<ComponentKind, number>,
  maxDimByKind?: Map<ComponentKind, number>
): Promise<JobSignature> {
  const components = new Map<ComponentKind, SignatureComponent>();
  for (const kind of SIGNATURE_CANDIDATE_KINDS) {
    const model = models.find((m) => isStlFile(m.name) && classifyComponentFilename(m.name) === kind);
    if (!model) continue;
    try {
      // Pass the CAD's own largest dimension so mm-vs-inch is decided by a
      // measured ratio, not by whether the part happens to be big.
      const geometry = await loadStlGeometry(model.url, maxDimByKind?.get(kind));
      components.set(kind, { model, stats: computeGeomStats(geometry), massLb: massByKind?.get(kind) });
    } catch {
      // skip this component if its file fails to load
    }
  }

  let base: SignatureComponent | undefined;
  const baseModel = findBaseModel(models);
  if (baseModel) {
    try {
      // The whole base is at least as large as the biggest single plate.
      let baseHint = 0;
      if (maxDimByKind) {
        for (const v of maxDimByKind.values()) baseHint = Math.max(baseHint, v);
      }
      const geometry = await loadStlGeometry(baseModel.url, baseHint > 0 ? baseHint : undefined);
      base = { model: baseModel, stats: computeGeomStats(geometry) };
    } catch {
      // base overlay is optional
    }
  }

  return { jobId, displayName, components, base };
}

/**
 * Measured SolidWorks mass (lb) per component kind, from a job's part rows.
 * Parts are keyed by CAD index and models by filename, so the two are bridged
 * the same way everything else in the app bridges them -- by classifying the
 * part's role label / component name into a ComponentKind.
 */
export function massByKindFromParts(
  parts: { role_label?: string; Component?: string; MassLb?: string }[]
): Map<ComponentKind, number> {
  const out = new Map<ComponentKind, number>();
  for (const p of parts || []) {
    const lb = parseFloat(p.MassLb || "");
    if (!Number.isFinite(lb) || lb <= 0) continue;
    const kind = classifyQuoteComponentName(p.role_label || p.Component || "");
    if (kind === "OTHER") continue;
    // Several CAD parts can share a kind (rails); sum them so the kind's mass
    // reflects the same set of solids its merged STL contains.
    out.set(kind, (out.get(kind) ?? 0) + lb);
  }
  return out;
}

/**
 * Largest CAD dimension in inches per component kind, from the part rows.
 *
 * Handed to loadStlGeometry so unit detection is a measured ratio against known
 * truth rather than a guess about magnitude. The CAD export is always inches.
 */
export function maxDimByKindFromParts(
  parts: { role_label?: string; Component?: string; Thickness?: string; Width?: string; Length?: string }[]
): Map<ComponentKind, number> {
  const out = new Map<ComponentKind, number>();
  for (const p of parts || []) {
    const dims = [p.Thickness, p.Width, p.Length].map((d) => parseFloat(d || "")).filter(Number.isFinite);
    if (dims.length === 0) continue;
    const maxIn = Math.max(...dims);
    if (!(maxIn > 0)) continue;
    const kind = classifyQuoteComponentName(p.role_label || p.Component || "");
    if (kind === "OTHER") continue;
    // Keep the biggest, so a merged multi-part STL is never scaled down.
    out.set(kind, Math.max(out.get(kind) ?? 0, maxIn));
  }
  return out;
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
    components.push({
      kind,
      ...compareComponentSignature(cur.stats, cand.stats, cur.massLb, cand.massLb),
    });
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
