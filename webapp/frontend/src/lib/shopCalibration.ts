/**
 * Calibration against a real posted program.
 *
 * SOURCE
 *   J8410 ID POT 5X.NC — Mastercam 2026, post MPPOSTABILITY_HAAS_UMC-1000.PST,
 *   dated JUN-10-2026. An actual ID POT for BMS-833200084, ~20,000 lines, so
 *   this is what the shop really does rather than what a generic model assumes.
 *
 * Datum's engine is calibrated for a generic 3-axis machining centre with one
 * surface speed per material. This shop runs a 5-axis UMC-1000 with high-feed
 * indexable mills and comparatively slow drilling. Two of its assumptions are
 * badly wrong here, and both are measurable.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * FINDING 1 — SETUPS. Datum counts a "setup" per distinct machined direction,
 * and my adapter charged 52 min each. The real program indexes the B and C axes
 * to FIVE orientations:
 *
 *     B0.  C0.      B90. C0.     B90. C90.     B90. C180.    B90. C270.
 *
 * That is ONE setup. The part is clamped once and the machine rotates to it —
 * `M11 / B90. C180. / M10` is a few seconds of brake-release, index, brake-set,
 * not a re-fixture. Charging 3 x 52 = 156 min of setup on the ID POT was the
 * single largest error in the estimate.
 *
 * FINDING 2 — SPEEDS BY OPERATION CLASS. Datum uses one vc (140 m/min for
 * 4140) for everything. Observed, converted to surface speed:
 *
 *   tool                        dia      real rpm   implied vc   Datum vc
 *   T24  2" Zenit HF mill       2.000"      1308     209 m/min      140
 *   T23  1" Zenit HF mill       1.000"      2578     206 m/min      140
 *   T14  0.5" flat EM Widin     0.500"      4250     169 m/min      140
 *   T5   0.45" chamfer mill     0.450"      2200      80 m/min      140
 *   T152 0.5" drill             0.500"      1757      70 m/min      140
 *   T35  0.625" drill           0.625"       764      47 m/min      140
 *   T7   0.625" spot drill      0.625"       397      20 m/min      140
 *   T151 0.406" drill           0.406"       395      13 m/min      140
 *
 * So high-feed indexable mills run ~205 m/min (Datum is too SLOW, overstating
 * time), while drilling runs 13–70 m/min and spotting ~20 (Datum is far too
 * FAST, understating time). Those errors run in opposite directions, which is
 * why a total can look plausible while every line is wrong.
 *
 * Time scales inversely with spindle speed at fixed chip load, so each factor
 * below is Datum_vc / real_vc.
 * ─────────────────────────────────────────────────────────────────────────────
 *
 * These are ONE part on ONE machine. Treat them as a starting point: re-derive
 * from two or three more programs and the model tightens considerably.
 */

/**
 * Multiply Datum's minutes by these to land on observed reality.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * REVISION 2 — measured, not inferred.
 *
 * The first pass derived these from spindle speeds in the program header. That
 * was better than nothing but still indirect. These come from simulating the
 * actual toolpath of J8441 ID POT 5X1.NC: all 40,645 lines, every G01, every arc
 * unwrapped from its I/J centre, canned cycles expanded by peck depth.
 *
 *     simulated feed time      266 min
 *     + tool changes & rapids  281 min
 *     estimator said           311 min
 *
 * The TOTAL was close, which is exactly why the errors survived: they cancel.
 * Underneath, three factors were badly wrong and two were wrong in opposite
 * directions.
 *
 *   operation        claimed   actual   old factor   new factor
 *   helical boring     12.3     47.3       1.0          3.8
 *   semi + finish     124.6     62.8       0.85         0.43
 *   spot drilling      16.5      3.5       5.5          1.15
 *   chamfering          4.1     14.7       1.75         6.3
 *
 * The boring error has a clean explanation. Datum reported "148 helical laps"
 * where the program cuts 559 — a Ø2.750 x 6.70 deep bore at 0.012 in per lap.
 * 559/148 = 3.78, and 12.3 x 3.78 = 46.5, which lands on the measured 47.3. Its
 * helical pitch is roughly 0.045 in/lap against the 0.012 the shop actually
 * runs. Op 43 alone (.062R INSERT FINISH BORE SIZE 2.75, T42 2" bull-nose) is
 * 47.3 min, 18% of all cutting -- the single longest operation in the program.
 *
 * The finishing error is a wrong tool AND a wrong algorithm: Datum picked a
 * 0.123" endmill (not in the crib) and rastered 1086 cm2 at 0.60 mm stepover.
 * The program finishes with T19, T13 and T174 as profile contours per Z step,
 * not as an area raster. The 0.43 factor compensates for the symptom; the tool
 * cascade is the real fix.
 * ─────────────────────────────────────────────────────────────────────────────
 */
export const TIME_FACTORS = {
  /** Indexable high-feed mills (Zenit/Dijet, >= 0.625"). Real is faster. */
  faceMill: 0.67,
  /** Big-tool roughing with an indexable mill. */
  rough: 0.7,
  /**
   * SUPERSEDED for mesh-path estimates, kept for the server's coarse path.
   *
   * This 0.43 was a patch for two stacked errors: the engine sized the finisher
   * to the tightest internal corner, and it rastered the part's ENTIRE mesh
   * surface (~3000 cm² on a pot block) as though all of it needed finishing.
   * A single multiplier cannot fix both, and it only ever worked on parts shaped
   * like the one it was fitted to -- on the OD holder the same factor still left
   * 698 min of finishing against 95 on the ID holder, two halves of one tool.
   *
   * planFinishPasses() in roughingCascade.ts now prices finishing from the area
   * with a real finisher plus a corner pass, and those ops are marked
   * preCalibrated so this factor is NOT applied on top of them.
   */
  finish: 0.43,
  /** Same tooling and same algorithm error as finish. Also superseded. */
  semiFinish: 0.43,
  /** Drilling. Observed 13-70 m/min against an assumed 140. */
  drill: 2.5,
  /**
   * Spot drilling. The old 5.5 was the worst factor in the model and it ran the
   * wrong way: 16.5 min claimed against 3.5 measured. Spotting is a peck to
   * depth, not a surface-speed-limited cut, so scaling it by a vc ratio was
   * simply the wrong physics.
   */
  spot: 1.15,
  /** Tapping: Datum's 676 rpm matches what the post emits. Leave alone. */
  tap: 1.0,
  /** Eight separate chamfer ops (T5, T6) totalling 14.7 min against 4.1 claimed. */
  chamfer: 6.3,
  /**
   * Helical boring and interpolation. See the header -- Datum's helical pitch is
   * ~3.8x coarser than the program's 0.012 in/lap.
   */
  bore: 3.8,
  /** Ball-nose contouring, no direct measurement yet. */
  contour: 1.0,
} as const;

/**
 * Deburr is ON the machine, not beside it.
 *
 * The model books deburr as bench work running parallel to the next part, so it
 * costs nothing in cycle. Op 27 of the reference program is a real DEBURR cycle
 * with T10 on the spindle: 11.7 min of actual spindle time. Some hand work does
 * happen at the bench afterwards, but the machine time is real and has to be in
 * the cycle.
 */
export const DEBURR_ON_MACHINE = true;

/**
 * Setup model for a 5-axis UMC-1000.
 *
 * One fixturing, then the machine indexes. Datum's "setups" are really tool
 * orientations, so they cost an index and a re-prove, not a re-clamp.
 */
export const SETUP_MODEL = {
  /** Clamp, indicate, set work offsets. Once per part. */
  firstSetupMin: 52,
  /** Each additional B/C orientation: prove-out and first-part check. */
  perOrientationMin: 8,
  /** True if the shop's mill can index. Set false for a 3-axis machine. */
  fiveAxis: true,
};

/**
 * Setup minutes for a part, given how many distinct orientations the plan found.
 *
 * On a 5-axis machine the orientations are indexes off one setup. On a 3-axis
 * machine each one really is a re-fixture, so it falls back to Datum's model.
 */
export function setupMinutes(orientations: number): {
  min: number;
  detail: string;
} {
  const n = Math.max(1, Math.round(orientations || 1));

  if (!SETUP_MODEL.fiveAxis) {
    return {
      min: n * SETUP_MODEL.firstSetupMin,
      detail: `${n} re-fixture(s) at ${SETUP_MODEL.firstSetupMin} min (3-axis)`,
    };
  }

  const extra = Math.max(0, n - 1);
  const min = SETUP_MODEL.firstSetupMin + extra * SETUP_MODEL.perOrientationMin;
  return {
    min,
    detail:
      n === 1
        ? `1 setup, ${SETUP_MODEL.firstSetupMin} min`
        : `1 setup (${SETUP_MODEL.firstSetupMin} min) + ${extra} B/C orientation(s) ` +
          `at ${SETUP_MODEL.perOrientationMin} min — UMC-1000 indexes, it does not re-clamp`,
  };
}

/** Pick the right factor for an op label. */
export function timeFactorFor(opLabel: string): { f: number; cls: string } {
  const s = opLabel.toLowerCase();
  if (/^face/.test(s)) return { f: TIME_FACTORS.faceMill, cls: "faceMill" };
  if (/^spot/.test(s)) return { f: TIME_FACTORS.spot, cls: "spot" };
  if (/^drill/.test(s)) return { f: TIME_FACTORS.drill, cls: "drill" };
  if (/^tap/.test(s)) return { f: TIME_FACTORS.tap, cls: "tap" };
  if (/chamfer/.test(s)) return { f: TIME_FACTORS.chamfer, cls: "chamfer" };
  if (/semi-finish/.test(s)) return { f: TIME_FACTORS.semiFinish, cls: "semiFinish" };
  if (/finish walls/.test(s)) return { f: TIME_FACTORS.finish, cls: "finish" };
  if (/^rough/.test(s)) return { f: TIME_FACTORS.rough, cls: "rough" };
  if (/bore|interpolate|ream|counterbore/.test(s)) return { f: TIME_FACTORS.bore, cls: "bore" };
  if (/contour/.test(s)) return { f: TIME_FACTORS.contour, cls: "contour" };
  return { f: 1.0, cls: "none" };
}

/**
 * What the reference program actually used, for a sanity line in the UI.
 *
 * 46 tool changes over 28 distinct tools — tools get recalled for later
 * orientations (T24 appears 5 times, once per B/C position). An estimate that
 * reports "29 tools" is counting distinct tools, not changes, and will
 * understate tool-change time by roughly 40%.
 */
export const REFERENCE_JOB = {
  name: "J8441 ID POT 5X1 (UMC-1000, Mastercam 2026)",
  distinctTools: 28,
  toolChanges: 46,
  /**
   * 575 rotary index moves across 318 distinct B positions, ~16,000 deg of
   * travel, brakes released once per operation. "3 setups, 2 orientations" does
   * not describe this program at all -- it is one clamp with hundreds of
   * indexes, and the real indexing cost is roughly 5-8 min.
   */
  orientations: 5,
  /** Roughing was done with a 2" indexable, then stepped down. */
  roughingTool: 'T24 2" Zenit HF mill',
  /** Simulated feed time over all 40,645 lines. */
  measuredFeedMin: 266,
  /** With tool changes and rapids. */
  measuredTotalMin: 281,
  note:
    "Roughed with a 2 inch indexable high-feed mill, then cascaded down through " +
    "1 inch, 0.75, 0.5, 0.375, 0.25 and 0.1875. Datum picks a single rougher and " +
    "then shrinks it to the smallest internal corner radius, which is how it " +
    "ended up roughing a pot block with a 0.123 inch cutter that is not even in " +
    "the crib.",
  longestOp:
    'Op 43 ".062R INSERT FINISH BORE SIZE 2.75" on T42 (2" bull-nose): 47.3 min, ' +
    "18% of all cutting. 559 helical laps, Ø2.750 x 6.70 deep at 0.012 in/lap, F28.",
};
