/**
 * Cutting rates MEASURED off posted G-code, and the scaling model built on them.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * SOURCES — two part families now, not one
 *
 *   ID POT      J8441 ID POT 5X1.NC, J8410 ID POT 5X.NC       (the original pair)
 *   ID HOLDER   8471 x5 programs, 8472 x2 programs, plus
 *               FRONT LEFT BACK.NC and LAST OF IT.NC          (added 2026-08-05)
 *
 * 93,245 lines total, simulated line by line by geometry_classifier/
 * gcode_calibrate.py: arcs unwrapped from their I/J centres, canned cycles
 * expanded per hole, modal state tracked the way the control tracks it.
 * Raw output is checked in beside the script under data/calibration/.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * THE FINDING THAT MAKES EXTRAPOLATION POSSIBLE
 *
 * The shop runs LOCKED SPEEDS AND FEEDS PER TOOL. Every tool shared between the
 * two families came back at an identical spindle speed, on different parts, cut
 * months apart:
 *
 *     tool                     ID POT (2026-06)   ID HOLDER (2026-08)
 *     T24  2" Zenit HF mill        1308 rpm            1308 rpm
 *     T23  1" Zenit HF mill        2578 rpm            2578 rpm
 *     T7   0.625" spot drill        397 rpm             397 rpm
 *
 * That is a tooling library with saved parameters, not a programmer choosing
 * numbers per job. It means surface-speed physics is the WRONG model here --
 * Datum computes rpm from vc and diameter and gets answers the machine never
 * runs. The right model is a lookup:
 *
 *     minutes = cut distance / FEED(tool class)
 *
 * and the only thing that varies part to part is the DISTANCE. Distance is
 * geometry, which is exactly what the mesh analysis measures. So every rate
 * below is a constant of the shop, and everything part-specific is area, volume,
 * edge length or hole count.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * WHAT THE SECOND SOURCE CONFIRMED
 *
 *   Roughing, 0.75" Zenit HF   195 ipm asserted  ->  189.0 ipm measured   (-3%)
 *   Facing,   2"    Zenit HF   124.7 ipm         ->  128.9 ipm            (+3%)
 *   Finish feed, solid carbide 6.6 ipm           ->  5.4-6.8 ipm          (in band)
 *
 * Three independent confirmations from a different part. The high-feed constant-
 * table-feed finding holds, and FINISH_FEED_IPM in roughingCascade.ts is sound.
 *
 * WHAT IT CHANGED
 *
 *   Semi-finish is NOT the same speed as finish. Measured 7.3-9.1 ipm against
 *   5.4-6.8 for finish -- about 1.35x faster. The model reuses one feed for both
 *   and only doubles the stepover, so semi-finish was over-quoted.
 *
 *   Chamfering is not a rounding error. 70 min on 8472, 14% of the part.
 *
 *   Finisher diameter goes to 1.0" on holders. See FINISHER_DIA_OBSERVED.
 */

/** Where a number came from, so no rate in here is unattributable. */
export type RateSource = "id-pot" | "id-holder" | "both";

export interface MeasuredRate {
  /** Operation class as gcode_calibrate.py labels it. */
  cls: string;
  /** Representative tool diameter, inches. */
  diaIn: number;
  /** Measured cut distance / cut time, in/min. */
  ipm: number;
  /** Minutes of evidence behind it. Weight by this, not by op count. */
  minutes: number;
  source: RateSource;
  note: string;
}

/**
 * The table. Sorted by weight of evidence, because a rate backed by 100 minutes
 * of cutting deserves more trust than one backed by two.
 */
export const MEASURED_RATES: MeasuredRate[] = [
  {
    cls: "rough", diaIn: 0.75, ipm: 189.0, minutes: 100.5, source: "both",
    note: "0.75\" Zenit high-feed. Confirms the 195 ipm constant-table-feed finding to within 3%.",
  },
  {
    cls: "semiFinish", diaIn: 1.0, ipm: 168.1, minutes: 25.4, source: "id-holder",
    note: "1\" Zenit HF semi-finishing. High-feed indexables semi-finish far faster than solid carbide.",
  },
  {
    cls: "semiFinish", diaIn: 0.75, ipm: 156.3, minutes: 39.5, source: "id-holder",
    note: "0.75\" Zenit HF. Same tool as the 189 ipm rougher, 17% slower in the semi role.",
  },
  {
    cls: "faceMill", diaIn: 2.0, ipm: 128.9, minutes: 30.0, source: "both",
    note: "2\" Zenit HF facing. Independently 124.7 ipm on the pots -- 3% apart.",
  },
  {
    cls: "finish", diaIn: 3.0, ipm: 20.0, minutes: 39.9, source: "id-holder",
    note: "3\" indexable face mill, FINISH facing pass. Not a wall finisher.",
  },
  {
    cls: "semiFinish", diaIn: 1.0, ipm: 9.1, minutes: 29.0, source: "id-holder",
    note: "1\" solid flat endmill.",
  },
  {
    cls: "semiFinish", diaIn: 0.75, ipm: 8.3, minutes: 87.7, source: "id-holder",
    note: "0.75\" solid flat endmill. The heaviest single semi-finish measurement in the set.",
  },
  {
    cls: "semiFinish", diaIn: 0.5, ipm: 7.3, minutes: 55.3, source: "id-holder",
    note: "0.5\" Dapra.",
  },
  {
    cls: "finish", diaIn: 0.5, ipm: 6.8, minutes: 48.7, source: "id-holder",
    note: "0.5\" Dapra finishing walls.",
  },
  {
    cls: "finish", diaIn: 0.375, ipm: 6.0, minutes: 15.0, source: "both",
    note: "0.375\" solid finisher. Pots measured 7.6-8.1 ipm on the same size.",
  },
  {
    cls: "finish", diaIn: 0.75, ipm: 5.6, minutes: 117.7, source: "id-holder",
    note: "0.75\" long finish endmill -- the single largest measurement in the whole set.",
  },
  {
    cls: "finish", diaIn: 1.0, ipm: 5.4, minutes: 43.8, source: "id-holder",
    note: "1\" long finish endmill. Slowest of the finishers, and the biggest.",
  },
  {
    cls: "chamfer", diaIn: 1.0, ipm: 13.3, minutes: 7.4, source: "id-holder",
    note: "1\" chamfer mill.",
  },
  {
    cls: "chamfer", diaIn: 0.45, ipm: 6.1, minutes: 19.1, source: "both",
    note: "0.45\" chamfer mill, the workhorse. Pots: 14.7 min over 8 ops.",
  },
  {
    cls: "drill", diaIn: 0.265, ipm: 2.0, minutes: 8.3, source: "both",
    note: "Peck drilling is 2.0 ipm across EVERY drill measured, 0.154\" to 0.640\". " +
      "Independent of diameter -- it is a peck cycle, not a surface-speed-limited cut.",
  },
  {
    cls: "tap", diaIn: 0.3125, ipm: 6.8, minutes: 2.8, source: "both",
    note: "Rigid tapping. 4.6-8.9 ipm across 1/4-20 to 1/2-13, rising with pitch.",
  },
  {
    cls: "spot", diaIn: 0.625, ipm: 1.3, minutes: 2.3, source: "both",
    note: "Spotting is slow and gets quoted as though it were fast. 397 rpm on both families.",
  },
];

/** Best measured feed for a class, weighted by minutes of evidence. */
export function measuredIpm(cls: string, diaIn?: number): number | null {
  const hits = MEASURED_RATES.filter((r) => r.cls === cls);
  if (hits.length === 0) return null;
  if (diaIn && diaIn > 0) {
    // Nearest diameter wins; ties go to the heavier evidence.
    const best = [...hits].sort(
      (a, b) => Math.abs(a.diaIn - diaIn) - Math.abs(b.diaIn - diaIn) || b.minutes - a.minutes,
    );
    return best[0].ipm;
  }
  const wsum = hits.reduce((s, r) => s + r.minutes, 0);
  return wsum > 0 ? hits.reduce((s, r) => s + r.ipm * r.minutes, 0) / wsum : hits[0].ipm;
}

/**
 * Semi-finish runs faster than finish with the SAME tool.
 *
 * 8.3 ipm against 5.6 on the 0.75", 7.3 against 6.8 on the 0.5", 9.1 on the 1"
 * solid. Weighted across the solid-carbide sizes that matters to 1.35x.
 * roughingCascade.ts prices both passes off FINISH_FEED_IPM and only widens the
 * stepover, so semi-finish carried the finishing feed and came out slow.
 */
export const SEMI_FINISH_SPEEDUP = 1.35;

/**
 * Finisher diameters the shop ACTUALLY uses, by part family.
 *
 * ⚠ THIS IS A KNOWN GAP, LEFT DELIBERATELY UNPATCHED. Read before changing
 * FINISH_MAX_DIA_IN in roughingCascade.ts.
 *
 *     ID POT     T174 0.499", T13 0.375", T19 0.1875"      max 0.5"
 *     ID HOLDER  T173 1.0", T190 0.75", T25 0.5", T13 .375" max 1.0"
 *
 * FINISH_MAX_DIA_IN is 0.5, which fits the pots and is half what the holders
 * use. Finishing minutes go as 1/D, so a holder finished on paper with a 0.5"
 * cutter when the shop ran a 1" is quoted at roughly 2x.
 *
 * It has NOT been raised here because FINISH_AREA_FRACTION (0.053) was FITTED at
 * D = 0.5 to reproduce the pot's measured 62.8 min. The two constants are one
 * fit. Raising the cap alone would halve finishing on pot blocks as a side
 * effect and silently break the calibration point it was derived from -- the
 * exact failure this codebase has documented twice already.
 *
 * The correct fix needs the pair re-anchored together, on a part whose finisher
 * diameter and finish minutes are both known. Re-run gcode_calibrate.py on
 * J8441 and refit; do not adjust one without the other.
 */
export const FINISHER_DIA_OBSERVED = {
  idPotMaxIn: 0.5,
  idHolderMaxIn: 1.0,
  modelCapIn: 0.5,
  overQuoteFactorOnHolders: 2.0,
} as const;

/**
 * ─────────────────────────────────────────────────────────────────────────────
 * EXTRAPOLATING TO A PART THAT IS NOT THESE
 *
 * With feeds fixed per tool, every class scales with a geometric quantity the
 * mesh analysis already measures. Bigger part, smaller part, more slots, no
 * slots -- all of it reduces to changing one of these inputs.
 *
 *   class        scales with              exponent on a uniform size change
 *   faceMill     face area x passes       L²
 *   rough        volume removed           L³
 *   semiFinish   wall + floor area        L²
 *   finish       wall + floor area        L²
 *   chamfer      edge length              L¹
 *   drill/tap    hole count x depth       independent of L
 *   spot         hole count               independent of L
 *
 * So DOUBLING a plate in every direction does NOT double the time: roughing
 * goes up 8x, the two finishing classes 4x, chamfer 2x, and the hole work not
 * at all. On these two holders roughing is 26% and finishing 41%, so the same
 * part at 2x scale lands near 4.4x the minutes, not 2x and not 8x. Quoting a
 * big plate by scaling a small one linearly is wrong in both directions
 * depending on which class dominates.
 *
 * FEATURES ARE ADDITIVE, AND THE DATA GIVES A PRICE FOR ONE
 *
 * 8471 has NO key pocket. 8472 has two. That is a clean natural experiment, and
 * the key-pocket operations in 8472 are separable from everything else:
 *
 *     DYNAMIC ROUGH FRONT SIDE KEY POCKET   T12 0.375"     2.24 min
 *     DYNAMIC ROUGH BACK KEY POCKET         T12 0.375"     0.58 min
 *     KEY POCKET FINISH FRONTSIDE           T13 0.375"     8.75 min
 *     KEY POCKET CONTOUR FIN FRONT SIDE     T13 0.375"     4.36 min
 *     KEY POCKET CONTOUR FIN BACK SIDE      T5  0.45"     14.48 min
 *     KEY POCKET CONTOUR FIN BACK SIDE      T13 0.375"     1.86 min
 *                                                         ─────────
 *                                           2 pockets      32.3 min
 *
 * See SLOT_MINUTES. Note the split: 9% roughing, 91% finishing. A slot is not
 * expensive because of the metal in it -- there is almost none -- it is
 * expensive because it is a precision feature with walls to contour. Any model
 * that prices slots by removed volume will under-quote them by an order of
 * magnitude.
 */
export const SLOT_MINUTES = {
  /** Per key pocket / slot, measured on 8472: 32.3 min for two. */
  perSlotMin: 16.2,
  roughShare: 0.09,
  finishShare: 0.91,
  evidence: "8472 ID Holder, 2 key pockets, 6 operations, 32.3 min. 8471 has none.",
} as const;

/**
 * Where the minutes actually go, as a share of feed time.
 *
 * Useful as a smell test on any estimate: if a quote for a holder says
 * chamfering is 1% or finishing is 15%, something upstream is wrong.
 *
 * 8471 is finish-heavy because the roughing program for it is not in the set --
 * its 360.9 min is finishing and semi-finishing almost end to end. 8472's two
 * programs cover the part from face to tap, so its shares are the ones to
 * generalise from.
 */
export const CLASS_SHARE_ID_HOLDER = {
  rough: 0.26,
  finish: 0.22,
  semiFinish: 0.19,
  chamfer: 0.14,
  faceMill: 0.06,
  drill: 0.05,
  tap: 0.05,
  spot: 0.02,
  bore: 0.01,
  totalMinMeasured: 504.0,
  note:
    "8472 ID Holder, 2 programs, 504.0 min of FEED time. Tool changes, rapids, " +
    "setup and probing are on top -- see setupMinutes() in shopCalibration.ts. " +
    "The part is also missing its centre bore, so the real figure is higher.",
} as const;
