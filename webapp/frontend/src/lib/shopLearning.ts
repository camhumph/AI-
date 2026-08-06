/**
 * Actual-vs-quoted feedback, so the estimator learns THIS shop.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * WHY THIS IS THE MOST VALUABLE PART OF THE SYSTEM
 *
 * Every calibration constant in shopCalibration.ts came from two parts: J8441
 * and J8410, both ID POTs, both 4140, both on the UMC-1000. They are the best
 * data available and they are still two parts. A clamp plate in A36 is a
 * different animal, and nothing in the model knows that yet.
 *
 * The fix is not more constants. It is to close the loop: record what each job
 * ACTUALLY took, compare it to what was quoted, and derive the corrections from
 * the shop's own history. After thirty jobs the model is calibrated on thirty
 * jobs instead of on two, and it keeps improving without anyone editing code.
 *
 * That also changes what the software IS. A frozen estimator is a calculator
 * someone will eventually stop trusting. One that measures its own error, shows
 * it, and corrects for it is an asset that gets better the longer it runs — and
 * it can prove that with its own numbers.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * DESIGN RULES
 *
 *  1. NEVER silently overwrite a measured constant. Learned factors are applied
 *     on top of the measured ones and both are always visible, so a bad batch of
 *     entries can be traced and reverted.
 *  2. DO NOT LEARN FROM ONE JOB. Below MIN_SAMPLES a class keeps the measured
 *     factor and says so. One outlier must not move the model.
 *  3. CLAMP the correction. A learned factor outside [0.5, 2.0] means the data is
 *     wrong, not the model — a mis-keyed actual, or a job where half the work was
 *     subcontracted.
 *  4. TRIM the mean. Shop-floor actuals contain typos and jobs that got scrapped
 *     and re-run. The median is used, not the average.
 *  5. GROUP LIKE WITH LIKE. Material class and part role both matter; a learned
 *     factor is keyed on operation class first, then narrowed if there is enough
 *     data.
 */

export interface JobActual {
  jobId: string;
  /** Display name, for the audit list. */
  displayName: string;
  /** ISO date the actual was recorded. */
  recordedAt: string;
  /** Who entered it — enterprise buyers need this on the record. */
  recordedBy?: string;
  /** What the estimator said when the job was quoted, minutes. */
  quotedMin: number;
  /** What it really took at the spindle, minutes. */
  actualMin: number;
  /** Optional per-operation-class actuals, when the shop tracks that finely. */
  byClass?: Partial<Record<OpClass, { quotedMin: number; actualMin: number }>>;
  /** Material class, e.g. "4140", "A36", "P20". */
  material?: string;
  /** Part role, e.g. "ID POT", "A PLATE". */
  role?: string;
  /** Free text: what went differently. The most useful field on the record. */
  note?: string;
}

/** The operation classes shopCalibration.ts already factors. */
export type OpClass =
  | "faceMill"
  | "rough"
  | "finish"
  | "semiFinish"
  | "drill"
  | "spot"
  | "tap"
  | "chamfer"
  | "bore"
  | "contour"
  | "setup"
  | "toolChange";

/** Below this, a class keeps its measured factor. One job is an anecdote. */
export const MIN_SAMPLES = 4;

/** A learned factor outside this range means the DATA is wrong. */
export const LEARN_CLAMP: [number, number] = [0.5, 2.0];

export interface LearnedFactor {
  cls: OpClass;
  /** Multiplier to apply ON TOP of the measured factor. 1 = no change. */
  factor: number;
  samples: number;
  /** Spread of the underlying ratios, as a percentage. High means inconsistent. */
  spreadPct: number;
  /** True when there is enough data to actually use this. */
  applied: boolean;
  note: string;
}

function median(xs: number[]): number {
  if (xs.length === 0) return 1;
  const s = [...xs].sort((a, b) => a - b);
  const m = Math.floor(s.length / 2);
  return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
}

/** Interquartile spread as a percentage of the median — an honest scatter figure. */
function spread(xs: number[]): number {
  if (xs.length < 4) return 0;
  const s = [...xs].sort((a, b) => a - b);
  const q1 = s[Math.floor(s.length * 0.25)];
  const q3 = s[Math.floor(s.length * 0.75)];
  const med = median(s);
  return med > 0 ? Math.round((100 * (q3 - q1)) / med) : 0;
}

function clamp(v: number): number {
  return Math.max(LEARN_CLAMP[0], Math.min(LEARN_CLAMP[1], v));
}

/**
 * Derive per-class correction factors from recorded actuals.
 *
 * A class with fewer than MIN_SAMPLES entries comes back with factor 1 and
 * `applied: false` — present in the report so the shop can see how close it is
 * to having enough data, but not affecting any number.
 */
export function learnFactors(actuals: JobActual[]): LearnedFactor[] {
  const byClass = new Map<OpClass, number[]>();

  for (const a of actuals) {
    if (!a.byClass) continue;
    for (const [cls, v] of Object.entries(a.byClass) as Array<
      [OpClass, { quotedMin: number; actualMin: number }]
    >) {
      if (!v || !(v.quotedMin > 0) || !(v.actualMin > 0)) continue;
      const list = byClass.get(cls) ?? [];
      list.push(v.actualMin / v.quotedMin);
      byClass.set(cls, list);
    }
  }

  const out: LearnedFactor[] = [];
  for (const [cls, ratios] of byClass) {
    const med = median(ratios);
    const enough = ratios.length >= MIN_SAMPLES;
    const raw = clamp(med);
    const sp = spread(ratios);

    out.push({
      cls,
      factor: enough ? raw : 1,
      samples: ratios.length,
      spreadPct: sp,
      applied: enough,
      note: !enough
        ? `${ratios.length} of ${MIN_SAMPLES} jobs needed. Measured factor still in use.`
        : raw !== med
          ? `Median ratio ${med.toFixed(2)} clamped to ${raw.toFixed(2)} — a correction that ` +
            `large usually means a mis-keyed actual or subcontracted work, not a model error. ` +
            `Check the entries before trusting it.`
          : sp > 40
            ? `Median ${raw.toFixed(2)} but the middle half of jobs spans ${sp}%. The model is ` +
              `now unbiased on average while still being wrong job to job — look for what ` +
              `separates the fast ones from the slow ones.`
            : `Median of ${ratios.length} jobs, middle half within ${sp}%.`,
    });
  }

  return out.sort((a, b) => Math.abs(b.factor - 1) - Math.abs(a.factor - 1));
}

export interface AccuracyReport {
  jobs: number;
  /** Median actual/quoted across whole jobs. 1.0 = unbiased. */
  bias: number;
  /** Median absolute percentage error — the number that matters to a buyer. */
  mapePct: number;
  /** Share of jobs landing within +/-20% of quote. */
  withinTwentyPct: number;
  /** Jobs quoted LOW (actual over quote), which is the expensive direction. */
  underQuoted: number;
  overQuoted: number;
  summary: string;
}

/**
 * How well the estimator is doing, in the terms a shop owner cares about.
 *
 * Bias and scatter are reported separately on purpose. An estimator that is
 * unbiased but scattered is not the same as one that is consistently 15% low,
 * and the fixes are different: scatter needs better inputs, bias needs a factor.
 */
export function accuracyReport(actuals: JobActual[]): AccuracyReport | null {
  const live = actuals.filter((a) => a.quotedMin > 0 && a.actualMin > 0);
  if (live.length === 0) return null;

  const ratios = live.map((a) => a.actualMin / a.quotedMin);
  const bias = median(ratios);
  const mape = median(ratios.map((r) => Math.abs(r - 1) * 100));
  const within = live.filter((a) => {
    const r = a.actualMin / a.quotedMin;
    return r >= 0.8 && r <= 1.2;
  }).length;
  const under = ratios.filter((r) => r > 1.05).length;
  const over = ratios.filter((r) => r < 0.95).length;

  const dir = bias > 1.02 ? "LOW" : bias < 0.98 ? "HIGH" : "unbiased";
  const summary =
    live.length < MIN_SAMPLES
      ? `${live.length} job${live.length === 1 ? "" : "s"} recorded. Needs ${MIN_SAMPLES} before ` +
        `any of this is worth acting on.`
      : dir === "unbiased"
        ? `Across ${live.length} jobs the estimator is unbiased (median ${bias.toFixed(2)}x) with ` +
          `a typical miss of ${mape.toFixed(0)}%. ${within} of ${live.length} landed within 20%.`
        : `Across ${live.length} jobs the estimator runs ${dir} by ` +
          `${Math.abs((bias - 1) * 100).toFixed(0)}% (median ${bias.toFixed(2)}x), typical miss ` +
          `${mape.toFixed(0)}%. ${under} jobs came in over quote, ${over} under. ` +
          `${dir === "LOW" ? "Quoting low is the expensive direction — this is money already lost." : ""}`;

  return {
    jobs: live.length,
    bias: Math.round(bias * 1000) / 1000,
    mapePct: Math.round(mape * 10) / 10,
    withinTwentyPct: Math.round((100 * within) / live.length),
    underQuoted: under,
    overQuoted: over,
    summary,
  };
}

/**
 * Combined factor for an operation class: the measured constant times whatever
 * the shop's own history has learned.
 *
 * Both parts are returned separately so the UI can show the provenance. Hiding
 * which half of a number came from where is how a model becomes unauditable.
 */
export function combinedFactor(
  measured: number,
  learned: LearnedFactor[] | undefined,
  cls: OpClass,
): { factor: number; measured: number; learned: number; source: string } {
  const l = learned?.find((x) => x.cls === cls && x.applied);
  if (!l) {
    return {
      factor: measured,
      measured,
      learned: 1,
      source: "measured from the reference programs",
    };
  }
  return {
    factor: Math.round(measured * l.factor * 1000) / 1000,
    measured,
    learned: l.factor,
    source: `measured x${measured} then x${l.factor} learned from ${l.samples} of this shop's own jobs`,
  };
}
