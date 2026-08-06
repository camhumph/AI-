/**
 * Confidence bands, so a quote is a RANGE with named reasons.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * WHY
 *
 * A single number invites false precision. "4.7 hours" reads as though someone
 * measured it; in reality parts of it are measured off the mesh and parts of it
 * are modelled, and the modelled parts carry very different risk depending on
 * what the mesh actually supported. An estimator handed one number will quote
 * that number.
 *
 * This produces a low / expected / high band and, more importantly, the list of
 * things that widened it. That list is the deliverable — it tells whoever is
 * signing the quote exactly which assumptions to check, and it is what makes the
 * estimate auditable rather than merely plausible.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * HOW THE BAND IS BUILT
 *
 * Start from a base spread, then widen for each named risk. Nothing here is a
 * statistical confidence interval — there is no population to sample. It is an
 * engineering tolerance stack on the estimate itself, and every term is stated
 * so it can be argued with:
 *
 *   ASSUMED TIME.        Operations not measured off the mesh (setup, tool
 *                        changes) are modelled. The larger their share, the
 *                        wider the band.
 *   TOOL GAPS.           An op with no tool in the crib will be done some other
 *                        way, and that way is not priced.
 *   STOCK NOT FROM BOM.  Stock derived from finished size + allowance instead of
 *                        a real purchase row.
 *   CORNER LIMITED.      A tight internal corner drives finishing strategy and
 *                        is the least predictable part of any mold plate.
 *   LOW REACH.           Features the tool cannot get to from the analysed
 *                        directions imply extra setups nobody has planned yet.
 *   MESH DOUBT.          The finished-to-stock ratio is at the edge of credible.
 *
 * The asymmetry is deliberate: estimates miss LONG far more often than short,
 * because unplanned work gets discovered on the machine, not removed. So the
 * high side always opens wider than the low side.
 */

export interface ConfidenceInput {
  ops: Array<{ op: string; min: number; measured: boolean }>;
  toolGapCount: number;
  stockFromBom: boolean;
  stockShort: boolean;
  /** Finished volume as a percentage of stock volume. */
  remainingPct: number;
  cornerLimited: boolean;
  /** Datum's share of features reachable from the analysed directions, 0-100. */
  reachPct: number;
}

/**
 * What kind of risk a driver describes.
 *
 * This exists because the ROLL-UP has to aggregate drivers across plates, and it
 * cannot do that from prose. The previous version deduplicated by the first 40
 * characters of the sentence, which fails on every message that starts with its
 * own number: "5 operations have no tool in the crib" and "22 operations have no
 * tool in the crib" are different strings for the first 40 characters, so a
 * six-plate base printed the same driver five times with five different counts and
 * no indication of which plate each belonged to. The estimator then had five
 * lines that all said the same thing and an eight-line budget almost entirely
 * consumed by one risk, hiding the others.
 */
export type DriverKind =
  | "assumed-share"
  | "tool-gap"
  | "no-bom-stock"
  | "stock-short"
  | "corner"
  | "low-reach"
  | "heavy-removal";

export interface ConfidenceDriver {
  kind: DriverKind;
  /** The line as shown for a single plate. */
  text: string;
  /**
   * The number that parameterises this driver: op count for tool-gap, percent for
   * low-reach / assumed-share / heavy-removal. Undefined for the boolean kinds.
   * The roll-up aggregates on this rather than re-parsing `text`.
   */
  value?: number;
}

export interface ConfidenceBand {
  /** 0-100. High means the number rests mostly on measurement. */
  score: number;
  /** Multiplier on the expected minutes for the optimistic end. */
  lowFactor: number;
  /** Multiplier on the expected minutes for the pessimistic end. */
  highFactor: number;
  grade: "A" | "B" | "C" | "D";
  /**
   * One line per thing that widened the band, in plain shop language.
   *
   * Kept as plain strings because this is the rendered surface and the API mirror
   * in api/client.ts is deliberately structural -- a Python endpoint can populate
   * the same shape without knowing about DriverKind. `driverDetail` carries the
   * machine-readable form for the roll-up.
   */
  drivers: string[];
  /** Same drivers, with the kind and value the roll-up aggregates on. */
  driverDetail: ConfidenceDriver[];
  /** What the grade means for how the number should be used. */
  guidance: string;
}

/** Narrowest band worth quoting. Even an all-measured part is not exact. */
const BASE_LOW = 0.92;
const BASE_HIGH = 1.12;

/**
 * Risk weight at which a plate is grade D.
 *
 * Set so that roughly three significant problems on one plate takes it out of
 * quotable territory. Any single problem should downgrade it, not condemn it.
 */
const RISK_FULL = 0.8;

/**
 * Share of modelled time above which it is worth mentioning.
 *
 * Deliberately high. Setup and tool changes legitimately dominate a SMALL plate —
 * a 52-minute setup on a 20-minute cut is normal, not a warning sign, and an
 * earlier version that flagged anything over 15% modelled gave a perfectly clean
 * plate a grade D. Scoring the normal case as a failure trains people to ignore
 * the score, which defeats the entire purpose.
 */
const ASSUMED_SHARE_NOTABLE = 0.75;

export function plateConfidence(input: ConfidenceInput): ConfidenceBand {
  const { ops, toolGapCount, stockFromBom, stockShort, remainingPct, cornerLimited, reachPct } =
    input;

  const total = ops.reduce((s, o) => s + o.min, 0);
  const assumed = ops.filter((o) => !o.measured).reduce((s, o) => s + o.min, 0);
  const assumedShare = total > 0 ? assumed / total : 1;

  // Weighted risk, then one mapping to band and grade. Accumulating band widths
  // directly (the old approach) made the grade depend on the ORDER penalties were
  // applied and had no defined worst case.
  let risk = 0;
  const driverDetail: ConfidenceDriver[] = [];

  if (assumedShare > ASSUMED_SHARE_NOTABLE) {
    risk += 0.12;
    driverDetail.push({
      kind: "assumed-share",
      value: Math.round(100 * assumedShare),
      text:
        `${Math.round(100 * assumedShare)}% of the time is setup and tool changes rather than ` +
        `cutting. On a plate this small the fixed costs dominate, so the total moves with how ` +
        `you actually run it more than with the geometry.`,
    });
  }

  if (toolGapCount > 0) {
    // A MUCH SMALLER PENALTY THAN THIS USED TO BE, because most of what it was
    // counting is not a risk.
    //
    // It charged 0.13 risk per note on the premise that "no tool in the crib"
    // meant the op would be done some other, unpriced way. The crib is one
    // machine's carousel; the shop owns the tooling and the job runs wherever
    // the tool is. Penalising that was docking the quote's confidence for
    // nothing, four notes deep on an ordinary plate.
    //
    // What is left in these notes IS worth a little risk -- a 10xD hole needing
    // a peck strategy or a gun drill, a tap drill running over nominal -- so
    // this is reduced rather than removed, and capped low.
    risk += Math.min(0.12, 0.04 * toolGapCount);
    driverDetail.push({
      kind: "tool-gap",
      value: toolGapCount,
      text:
        `${toolGapCount} tooling note${toolGapCount === 1 ? "" : "s"} — reach limits or ` +
        `size calls worth confirming. Tools that live on another machine are quoted ` +
        `normally and are not counted here.`,
    });
  }

  if (!stockFromBom) {
    risk += 0.12;
    driverDetail.push({
      kind: "no-bom-stock",
      text:
        `No BOM row for this plate, so stock is finished size + 3/16" per side. If the real ` +
        `purchase is bigger, roughing goes up with it.`,
    });
  }

  if (stockShort) {
    risk += 0.25;
    driverDetail.push({
      kind: "stock-short",
      text:
        `The BOM stock is SMALLER than the measured part. One of the two is wrong and this ` +
        `estimate cannot tell you which.`,
    });
  }

  if (cornerLimited) {
    risk += 0.15;
    driverDetail.push({
      kind: "corner",
      text:
        `A tight internal corner sets the finishing strategy. Corner work is the least ` +
        `predictable part of a mold plate — small radius changes move finishing time a lot.`,
    });
  }

  if (Number.isFinite(reachPct) && reachPct > 0 && reachPct < 85) {
    risk += 0.25;
    driverDetail.push({
      kind: "low-reach",
      value: Math.round(reachPct),
      text:
        `Only ${Math.round(reachPct)}% of features are reachable from the analysed ` +
        `directions. The rest need an orientation or fixture that is not in this plan.`,
    });
  }

  if (remainingPct > 0 && remainingPct < 25) {
    risk += 0.15;
    driverDetail.push({
      kind: "heavy-removal",
      value: Math.round(remainingPct),
      text:
        `${remainingPct.toFixed(0)}% of the stock survives as finished part. That is a lot of ` +
        `metal to move and the roughing model carries most of the estimate here.`,
    });
  }

  const riskNorm = Math.min(1, risk / RISK_FULL);

  // Asymmetric on purpose: estimates miss LONG far more often than short,
  // because unplanned work gets discovered at the machine, not removed.
  const low = Math.max(0.6, BASE_LOW - 0.15 * riskNorm);
  const high = Math.min(2.2, BASE_HIGH + 0.9 * riskNorm);

  const score = Math.round(100 * (1 - riskNorm));

  const grade: ConfidenceBand["grade"] =
    score >= 80 ? "A" : score >= 60 ? "B" : score >= 40 ? "C" : "D";

  const guidance =
    grade === "A"
      ? "Quotable. The number rests on measured geometry and a complete tool list."
      : grade === "B"
        ? "Quotable with a look at the drivers below. Nothing here is structurally unknown."
        : grade === "C"
          ? "Budgetary. Use the high end, or resolve the drivers before committing a price."
          : "Rough order of magnitude only. Do not quote off this without a person checking the part.";

  return {
    score,
    lowFactor: Math.round(low * 1000) / 1000,
    highFactor: Math.round(high * 1000) / 1000,
    grade,
    drivers: driverDetail.map((d) => d.text),
    driverDetail,
    guidance,
  };
}

/** "B Plate, A Plate and 2 more" — names the plates without running off the line. */
function namePlates(labels: string[], max = 3): string {
  const named = labels.filter(Boolean);
  if (named.length === 0) return "";
  if (named.length <= max) {
    if (named.length === 1) return named[0];
    return `${named.slice(0, -1).join(", ")} and ${named[named.length - 1]}`;
  }
  return `${named.slice(0, max).join(", ")} and ${named.length - max} more`;
}

/**
 * One line per KIND of risk, aggregated across plates and attributed to them.
 *
 * The point of this list is that each line is an action: go and resolve this.
 * "5 operations have no tool in the crib" repeated five times with different
 * counts is not an action, because it does not say where to look. So each kind
 * collapses to a single line carrying the total, the worst plate, and which
 * plates are affected.
 *
 * Ordered by how much risk the kind actually contributed, worst first, so the
 * line worth reading is the first one.
 */
function aggregateDrivers(
  live: Array<{ minutes: number; band: ConfidenceBand; label?: string }>,
): ConfidenceDriver[] {
  // kind -> the per-plate hits for it
  const byKind = new Map<DriverKind, Array<{ label: string; d: ConfidenceDriver }>>();
  for (const p of live) {
    const label = p.label || "";
    // Fall back to the rendered strings when driverDetail is absent -- a band that
    // came from the Python endpoint has only `drivers`. Those cannot be aggregated,
    // so they pass through as-is rather than being dropped.
    const detail: ConfidenceDriver[] =
      p.band.driverDetail && p.band.driverDetail.length > 0
        ? p.band.driverDetail
        : (p.band.drivers || []).map((t) => ({ kind: "assumed-share" as DriverKind, text: t }));
    for (const d of detail) {
      const list = byKind.get(d.kind) || [];
      list.push({ label, d });
      byKind.set(d.kind, list);
    }
  }

  const nPlates = live.length;
  const out: Array<{ driver: ConfidenceDriver; weight: number }> = [];

  for (const [kind, hits] of byKind) {
    const labels = hits.map((h) => h.label).filter(Boolean);
    const values = hits.map((h) => h.d.value).filter((v): v is number => Number.isFinite(v));
    const where = namePlates(labels);
    const onN =
      hits.length === 1
        ? where
          ? `on ${where}`
          : "on 1 part"
        : `across ${hits.length} of ${nPlates} parts${where ? ` (${where})` : ""}`;

    // A single hit already has a well-written sentence for exactly this case.
    if (hits.length === 1 && kind !== "tool-gap") {
      out.push({
        driver: { kind, value: hits[0].d.value, text: where ? `${where}: ${hits[0].d.text}` : hits[0].d.text },
        weight: kind === "low-reach" || kind === "stock-short" ? 3 : 2,
      });
      continue;
    }

    switch (kind) {
      case "tool-gap": {
        const totalOps = values.reduce((s, v) => s + v, 0);
        const worst = hits.reduce((a, b) => ((b.d.value ?? 0) > (a.d.value ?? 0) ? b : a));
        out.push({
          driver: {
            kind,
            value: totalOps,
            text:
              `${totalOps} operation${totalOps === 1 ? "" : "s"} ${onN} have no tool in the ` +
              `crib${
                hits.length > 1 && worst.label ? `, worst on ${worst.label} with ${worst.d.value}` : ""
              }. Whatever gets done instead — a special, an extension, subbing it out — is not ` +
              `in this number.`,
          },
          // Tool gaps scale: 50 missing tools is a different problem from 5.
          weight: 3 + Math.min(2, totalOps / 20),
        });
        break;
      }
      case "low-reach": {
        const worstPct = Math.min(...values);
        out.push({
          driver: {
            kind,
            value: worstPct,
            text:
              `Reachability is short ${onN} — as low as ${worstPct}% of features on the worst ` +
              `one. The rest need an orientation or fixture that is not in this plan.`,
          },
          weight: 4,
        });
        break;
      }
      case "stock-short": {
        out.push({
          driver: {
            kind,
            text:
              `BOM stock is SMALLER than the measured part ${onN}. One of the two is wrong and ` +
              `this estimate cannot tell you which.`,
          },
          weight: 4,
        });
        break;
      }
      case "corner": {
        out.push({
          driver: {
            kind,
            text:
              `A tight internal corner sets the finishing strategy ${onN}. Corner work is the ` +
              `least predictable part of a mold plate — small radius changes move finishing ` +
              `time a lot.`,
          },
          weight: 2,
        });
        break;
      }
      case "no-bom-stock": {
        out.push({
          driver: {
            kind,
            text:
              `No BOM row ${onN}, so stock is finished size + 3/16" per side. If the real ` +
              `purchase is bigger, roughing goes up with it.`,
          },
          weight: 2,
        });
        break;
      }
      case "heavy-removal": {
        const worstPct = Math.min(...values);
        out.push({
          driver: {
            kind,
            value: worstPct,
            text:
              `Heavy stock removal ${onN} — as little as ${worstPct}% of the stock survives as ` +
              `finished part. The roughing model carries most of the estimate there.`,
          },
          weight: 2,
        });
        break;
      }
      case "assumed-share": {
        const worstPct = Math.max(...values, 0);
        out.push({
          driver: {
            kind,
            value: worstPct,
            text:
              worstPct > 0
                ? `Setup and tool changes rather than cutting dominate ${onN} — up to ` +
                  `${worstPct}% of the time. The total moves with how it is actually run more ` +
                  `than with the geometry.`
                : hits[0].d.text,
          },
          weight: 1,
        });
        break;
      }
    }
  }

  return out
    .sort((a, b) => b.weight - a.weight)
    .slice(0, 8)
    .map((o) => o.driver);
}

/**
 * Roll several plates' bands into one for the job.
 *
 * Risks do NOT simply add — a shop running six plates absorbs a surprise on one
 * of them more easily than a single-part job absorbs the same surprise. But they
 * do not cancel either, because the drivers are correlated: the same operator,
 * the same crib, the same mesh quality. Weighting each plate's band by its share
 * of the total minutes, then damping the spread by sqrt(n), sits between "they
 * all go wrong together" and "they are independent".
 */
export function rollUpConfidence(
  parts: Array<{ minutes: number; band: ConfidenceBand; label?: string }>,
): ConfidenceBand | null {
  const live = parts.filter((p) => p.minutes > 0 && p.band);
  if (live.length === 0) return null;

  const total = live.reduce((s, p) => s + p.minutes, 0);
  const wLow = live.reduce((s, p) => s + (p.minutes / total) * p.band.lowFactor, 0);
  const wHigh = live.reduce((s, p) => s + (p.minutes / total) * p.band.highFactor, 0);

  const damp = Math.sqrt(live.length);
  const low = 1 - (1 - wLow) / damp;
  const high = 1 + (wHigh - 1) / damp;

  const score = Math.round(
    live.reduce((s, p) => s + (p.minutes / total) * p.band.score, 0),
  );
  const grade: ConfidenceBand["grade"] =
    score >= 80 ? "A" : score >= 60 ? "B" : score >= 40 ? "C" : "D";

  const driverDetail = aggregateDrivers(live);

  return {
    score,
    lowFactor: Math.round(low * 1000) / 1000,
    highFactor: Math.round(high * 1000) / 1000,
    grade,
    drivers: driverDetail.map((d) => d.text),
    driverDetail,
    guidance:
      grade === "A"
        ? "Quotable as it stands."
        : grade === "B"
          ? "Quotable after a look at the drivers."
          : grade === "C"
            ? "Budgetary. Quote the high end or resolve the drivers first."
            : "Rough order of magnitude. Needs a person on the part before it goes out.",
  };
}
