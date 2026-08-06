/**
 * Roughing as a TOOL CASCADE, priced from cutting parameters.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * THE PROBLEM THIS FIXES
 *
 * Datum picks a single rougher and, when the part has a tight internal corner,
 * shrinks that one cutter to fit the corner — then removes ALL the bulk with it.
 * On J8441 that meant hogging a pot block with a Ø0.123" endmill, a tool this
 * crib does not own. The resulting time was so wrong that a blanket x0.43 factor
 * was thrown at finishing just to drag the total back to something plausible.
 * That factor was papering over a modelling error and only worked on parts
 * shaped like the one it was fitted to.
 *
 * No shop works that way. You hog with the biggest cutter that fits, then step
 * down, each smaller tool clearing only what the last could not reach. The
 * corner radius is the FINISHER's problem — a finisher removes a skin, not bulk.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * WHY THIS DOES NOT SCALE OFF DATUM'S NUMBER
 *
 * The obvious shortcut is to take Datum's roughing minutes and rescale them for
 * a bigger cutter. That inherits its error: its minutes come from a tool that
 * does not exist and a rate law that does not hold for high-feed indexables. A
 * ratio built on a wrong number is still wrong.
 *
 * So the volume comes from Datum (mesh volume is the part it does well) and the
 * RATE comes from cutting parameters listed below, one row per ladder tool, each
 * auditable by anyone who runs the machine. If a machinist disagrees with a
 * number they can point at the row and change it.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * THE MODEL
 *
 *   1. VOLUME.  A cutter of diameter D clears bulk in proportion to D², so step
 *      i takes V_i = V · D_i² / ΣD_j². The big cutter does most of the work; each
 *      smaller one picks up rest material in the corners it could not enter.
 *
 *   2. RATE.    MRR_i = ae · ap · vf for that specific tool, from the table
 *      below — NOT one surface speed for everything.
 *
 *   3. REALITY. Ideal MRR assumes the cutter is always in full engagement. Real
 *      3+2 roughing retracts, repositions, ramps in, cuts air at the ends of
 *      passes and steps over. TOOLPATH_EFFICIENCY is that derate, and it is the
 *      one constant here still pinned to a single part — see its comment.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * VALIDATION — measured, from two real posted programs
 *
 * Every roughing operation in each program, simulated line by line: arcs
 * unwrapped from their I/J centres, canned cycles expanded by peck depth.
 *
 *     tool        J8441 ID POT 5X1      J8410 ID POT 5X
 *     2.000"      18.2 min   30.8%      17.9 min   30.4%
 *     1.000"      18.5 min   31.3%      18.0 min   30.6%
 *     0.750"      19.3 min   32.6%      19.9 min   33.7%
 *     0.500"       3.1 min    5.3%       3.1 min    5.3%
 *     TOTAL       59.1 min              59.0 min
 *
 * ⚠ NOTE ON THAT TABLE. Those four rows come from grouping by the word ROUGH in
 * the operation name, which also catches "RGH FACE". Properly separated, the 2"
 * Zenit on these parts is FACING (18.2 min, costed separately) and the actual
 * pocket/bore roughing is 40.9 min on a 1 / 0.75 / 0.5 ladder. Both figures are
 * in the calibration output; the cascade models the second.
 *
 * WHAT IS ESTABLISHED, AND WHAT IS NOT
 *
 *   ✓ SHAPE. Equal time per cascade step, for the high-feed indexables.
 *     18.5 vs 19.3 min on J8441, 18.0 vs 19.9 on J8410 — within 6% on two jobs
 *     cut months apart. The model reproduces this.
 *
 *   ✓ FEED RATES. Measured, not assumed. Every HF indexable roughing pass runs
 *     at exactly 195.0 ipm regardless of diameter; the dynamic 0.5" pass runs
 *     102.6. High-feed milling is programmed at constant table feed.
 *
 *   ✓ TOOL SELECTION. Reach-limited, and it now picks what the shop picked.
 *
 *   ✗ ABSOLUTE LEVEL. Time = volume / MRR, and the removal volume for the
 *     reference part is unknown — only the G-code was available, not the STL. So
 *     the model is internally consistent and shape-validated, but whether it is
 *     right in absolute minutes is genuinely open. That is the question the
 *     learning loop in shopLearning.ts answers from real job actuals, and it is
 *     why TOOLPATH_EFFICIENCY is 1.0 with a hook rather than a fitted constant
 *     that would hide the uncertainty.
 *
 * Source: J8441 ID POT 5X1.NC, J8410 ID POT 5X.NC — Mastercam 2026,
 * MPPOSTABILITY_HAAS_UMC-1000.PST. Extracted by
 * geometry_classifier/gcode_calibrate.py, which is the tool to re-run on every
 * new program so this table stops resting on two parts.
 *
 * ─────────────────────────────────────────────────────────────────────────────
 * UPDATE 2026-08-05 — it no longer rests on two parts.
 *
 * Nine more programs, two ID HOLDERS (8471, 8472), 93,245 lines. A different
 * part in the same family, and it confirms the two load-bearing numbers:
 *
 *     0.75" Zenit roughing   195 ipm asserted -> 189.0 measured   (-3%)
 *     2"    Zenit facing     124.7 measured   -> 128.9 measured   (+3%)
 *
 * The ✗ above still stands — absolute level is still open, because the holders
 * came without STLs too, so removal volume is still unknown. What is now settled
 * is the RATE law. Full second-source table, and what it changed, in
 * lib/measuredRates.ts.
 */

import { bestTool, END_MILLS, FACE_MILLS, type ShopTool } from "./shopTools";
import { SEMI_FINISH_SPEEDUP } from "./measuredRates";

/**
 * Cutting parameters per ladder tool, for 4140 at this shop's SFM.
 *
 * ap = axial depth per pass, ae = radial engagement, vf = table feed (in/min).
 * MRR = ae · ap · vf, cubic inches per minute.
 *
 * The three biggest are Zenit high-feed indexables: shallow ap, wide ae, very
 * high feed. That is why the top of the cascade moves so much metal, and why
 * modelling them at the same surface speed as a solid carbide endmill — which is
 * what the generic engine does — gets roughing badly wrong.
 *
 * The 0.5" row is trochoidal/dynamic milling: narrow ae, nearly full-flute ap.
 * Different strategy, comparable MRR, which is exactly why the shop reaches for
 * it instead of stepping down further.
 *
 * EVERY NUMBER HERE IS MEANT TO BE ARGUED WITH. They are ordinary shop values,
 * not measurements off the reference program, because a posted G-code file does
 * not state ae or ap.
 */
export interface LadderTool {
  diaIn: number;
  apIn: number;
  aeIn: number;
  vfIpm: number;
  /** Shorthand for the UI. */
  strategy: string;
}

/**
 * Radial engagement as a fraction of cutter diameter.
 *
 * One constant for the whole ladder, because with MEASURED feed rates that is
 * what the data supports — see the validation block. A per-step engagement
 * schedule was an earlier attempt to fit a shape that turned out to be an
 * artefact of assumed feed rates.
 */
export const AE_FRACTION_OF_DIA = 0.45;

/**
 * The ladder, with feed rates MEASURED from the posted programs.
 *
 * The measurement that mattered: every high-feed indexable roughing pass in both
 * programs runs at exactly 195.0 ipm — the 1" and the 0.75" alike. High-feed
 * milling is programmed at a constant table feed, NOT at a feed proportional to
 * diameter. The first version assumed 200 / 150 / 120 by diameter and that single
 * wrong assumption is what forced a fitted engagement schedule and a fitted
 * efficiency factor to paper over it.
 *
 * Extracted by geometry_classifier/gcode_calibrate.py as cut distance / cut time
 * per tool, so these are not catalogue numbers — they are what the machine ran.
 */
export const ROUGHING_LADDER: LadderTool[] = [
  // 195 ipm by consistency with the two smaller indexables, NOT the 124.7 ipm
  // this tool measures on the reference parts — because there it is only ever
  // FACING, which is a different operation with different parameters and is
  // costed separately. Using the facing rate here made the 2" step take 1.6x the
  // time of the 1" step, contradicting the measured equal-time result. The shop's
  // high-feed rate is a machine constant across diameters; that is the finding.
  { diaIn: 2.0, apIn: 0.04, aeIn: 0.9, vfIpm: 195, strategy: "high-feed indexable" },
  // 195.0 ipm measured, 3 operations across 2 programs.
  { diaIn: 1.0, apIn: 0.04, aeIn: 0.45, vfIpm: 195, strategy: "high-feed indexable" },
  // 195.0 ipm measured, 2 operations across 2 programs.
  //
  // CONFIRMED BY A SECOND PART FAMILY, 2026-08-05. The 8472 ID Holder runs this
  // tool for 100.5 min of roughing across 7 operations and measures 189.0 ipm --
  // 3% off, on a different part, months later. That is the strongest single
  // check this table has: the constant-table-feed finding was derived from pot
  // blocks and it holds on holders. See lib/measuredRates.ts.
  { diaIn: 0.75, apIn: 0.04, aeIn: 0.3375, vfIpm: 195, strategy: "high-feed indexable" },
  // 102.6 ipm measured (DYNAMIC ROUGH POCKET). Trochoidal: narrow engagement,
  // deep axial. ap is the least-supported figure on this row — a posted program
  // states feed but never states depth of cut.
  { diaIn: 0.5, apIn: 0.5, aeIn: 0.05, vfIpm: 103, strategy: "dynamic / trochoidal" },
  // No direct measurement below 0.5"; scaled from the 0.5" row.
  { diaIn: 0.375, apIn: 0.375, aeIn: 0.04, vfIpm: 100, strategy: "dynamic / trochoidal" },
  { diaIn: 0.25, apIn: 0.25, aeIn: 0.03, vfIpm: 90, strategy: "dynamic / trochoidal" },
];

/**
 * How roughing volume splits across the cascade: V_i proportional to D_i^1.
 *
 * This exponent was WRONG in the first version (2, i.e. area-like) and the error
 * was masked by the assumed feed rates. Once feed is measured as constant across
 * the indexables, the two candidates separate cleanly against the measured
 * per-step split of 1.00 / 0.96:
 *
 *     V ∝ D    predicts  1.00 / 0.99      <-- matches
 *     V ∝ D²   predicts  1.00 / 0.75
 *
 * Physically: each successive smaller cutter clears a BAND of rest material along
 * the walls and corners the last one could not enter, and the volume of that band
 * scales with the cutter that clears it, not with its square.
 */
export const VOLUME_EXPONENT = 1;

/** Ideal metal removal rate for a ladder tool, in³/min at full engagement. */
export function idealMrr(t: LadderTool): number {
  return t.apIn * t.aeIn * t.vfIpm;
}

/**
 * Real toolpath versus ideal engagement.
 *
 * Ideal MRR assumes the cutter never leaves the metal. A real 3+2 roughing
 * program retracts between regions, ramps in, runs out past the end of each
 * pass, steps over, and re-approaches at every one of the hundreds of B/C
 * positions in these programs.
 *
 * NOW 1.0, AND THAT IS THE POINT.
 *
 * The first version carried 0.313 here. That number existed only because the feed
 * rates were assumed: efficiency and removal volume were fitted jointly against
 * one measured total, which made them individually unidentifiable — two unknowns,
 * one equation. Any reviewer asking "why 0.313?" would have got a circular answer.
 *
 * Measuring feed from the G-code removed the need for it. The model now predicts
 * 40.9 min of roughing at a removal volume of 157 in³ with no fitted derate, and
 * 157 in³ is 37% of a 9.3 x 6.5 x 7" block — a plausible, falsifiable prediction
 * rather than a back-solved constant. Integrating the actual toolpath
 * (ae x ap x cut distance) over the roughing operations gives ~120 in³, so the
 * model is within about 25% on the volume it implies.
 *
 * The hook is kept, set to 1.0, because the learning loop should own this: the
 * "rough" class in shopLearning.ts multiplies straight into it. If a shop's real
 * jobs run consistently longer than quoted, that is where the correction lands —
 * as a measured fact about that shop, not a constant fitted here to one part.
 */
export const TOOLPATH_EFFICIENCY = 1.0;

/**
 * Smallest diameter that still counts as ROUGHING.
 *
 * Measured: in both reference programs every cutter below 0.5" is doing
 * semi-finish or finish work. Below this you are not clearing bulk, and
 * finishing is already its own operation. This is the line the generic engine
 * fails to draw — it lets a corner radius drag the one and only rougher down to
 * Ø0.123", so a skin cut's tool ends up removing cubic inches.
 */
export const MIN_ROUGHING_DIA_IN = 0.5;

/**
 * Largest cutter that fits, as a fraction of the part's smaller footprint side.
 *
 * Much beyond a third of the width and the cutter has nowhere to run out. On the
 * reference pot block (9.3 x 6.5") this gives 6.5/3 = 2.17" and selects the 2"
 * indexable — the tool the program actually uses.
 */
export const MAX_TOOL_FRACTION_OF_WIDTH = 1 / 3;

export interface CascadeStep {
  diaIn: number;
  tool: ShopTool;
  min: number;
  /** Share of roughing volume this step clears, 0-1. */
  volFraction: number;
  volIn3: number;
  /** Effective in³/min after the toolpath derate. */
  mrr: number;
  strategy: string;
  role: "bulk" | "rest";
  /** True when the crib tool cannot reach the depth required. */
  reachShort: boolean;
}

export interface CascadeResult {
  steps: CascadeStep[];
  totalMin: number;
  /** What Datum charged for the same removal with one small cutter. */
  singleToolMin: number;
  singleToolDiaIn: number;
  /** Volume the cascade is pricing, in³. */
  removeVolIn3: number;
  wasCornerLimited: boolean;
  /** True when the cascade was not applied and Datum's number stands. */
  passthrough: boolean;
  note: string;
}

/**
 * Names that are the wrong TYPE of tool for roughing, whatever their diameter.
 *
 * A bull-nose is a boring tool and a "finish" endmill is ground for finishing —
 * putting either on a hogging pass misrepresents what the shop does. This matters
 * because the first version searched face mills and endmills as one pool sorted
 * by reach, and the 2" bull-nose (0.5" reach) beat the 2" Zenit (0.375" reach).
 * The estimate then also flagged REACH SHORT against a tool nobody would have
 * used, which is worse than a wrong tool — it is a wrong tool with a warning
 * attached.
 */
const NOT_A_ROUGHER = /bull-?nose|finish|raptor|snap ring|dovetail|grease/i;

/**
 * Crib tool for a ladder size.
 *
 * High-feed and face mills are preferred whenever one exists at that diameter —
 * those are the Zenits and Dijets the shop actually hogs with. Solid endmills are
 * the fallback for the small end of the ladder where no indexable exists.
 */
function cribToolFor(diaIn: number, depthIn: number): { tool: ShopTool; reachShort: boolean } {
  const atSize = (list: ShopTool[]) =>
    list.filter((t) => Math.abs(t.dia - diaIn) <= 0.01 && !NOT_A_ROUGHER.test(t.name));

  // Preference order, not one merged pool: an indexable at the right size always
  // beats a solid endmill at the right size, regardless of reach.
  for (const candidates of [atSize(FACE_MILLS), atSize(END_MILLS)]) {
    if (candidates.length === 0) continue;
    const byReach = [...candidates].sort((a, b) => reachOf(b) - reachOf(a));
    const reaches = byReach.find((t) => reachOf(t) >= depthIn - 1e-6);
    return reaches ? { tool: reaches, reachShort: false } : { tool: byReach[0], reachShort: true };
  }
  // Nothing at this size on THIS machine's list, so the rougher comes off the
  // machine that has it. The ladder diameter is chosen from the part's own
  // geometry (see the header) and is the right cutter for the cut, so the step
  // is priced at that size rather than dropped. bestTool synthesises it.
  return { tool: bestTool(END_MILLS, diaIn, depthIn, "rougher"), reachShort: false };
}

/**
 * How deep a ROUGHER can actually get to — which is not its `depth` field.
 *
 * `ShopTool.depth` is the usable cutting depth PER PASS: 0.375" on the 2" Zenit,
 * because a high-feed indexable takes a shallow bite at very high feed. A rougher
 * reaches a 2" deep pocket by stepping down in Z, so comparing total pocket depth
 * against that per-pass figure is a category error. The first version did exactly
 * that and flagged REACH SHORT against all three Zenits on the reference part —
 * the correct tools, warned about for the wrong reason. False warnings on correct
 * answers are worse than no warnings, because people learn to click past them.
 *
 * The real limit is flute/overall length minus clearance for the holder.
 */
const HOLDER_CLEARANCE_IN = 0.5;

function reachOf(t: ShopTool): number {
  return Math.max(t.depth, t.oal - HOLDER_CLEARANCE_IN);
}

/* ══════════════════════════════════════════════════════════════════════════
   FINISHING
   ══════════════════════════════════════════════════════════════════════════ */

/**
 * Smallest cutter that should be doing a full-area finish pass.
 *
 * THE BUG THIS FIXES, in the user's own numbers:
 *
 *              ID HOLDER            OD HOLDER
 *   area       2470 cm²             3028 cm²      (+23%)
 *   tool       0.25" (T182)         0.123" (T176)
 *   stepover   1.05 mm              0.60 mm
 *   FINISH     94.9 min             698.1 min     (+636%)
 *
 * A 23% larger part took SEVEN TIMES longer to finish. Nothing about the parts
 * justifies that — they are two halves of the same tool. The whole difference is
 * that the engine sized the finisher to the part's tightest internal corner, and
 * on the OD holder that corner is smaller, so it rastered 3028 cm² with a
 * 0.123" cutter at a 0.6 mm stepover.
 *
 * That is not how anyone finishes a plate. You finish the open walls and floors
 * with a proper endmill, then take the corners the big tool could not reach with
 * a small one, over a small area. Same argument as the roughing cascade: a tight
 * corner is a LOCAL problem and must not set the tool for a GLOBAL operation.
 */
export const MIN_FINISH_DIA_IN = 0.375;

/**
 * Share of wall area that is genuinely corner, needing the small cutter.
 *
 * Corners are a length of blend along an edge, not a share of surface, so this
 * is a shape-dependent guess. 15% is deliberately generous — the failure mode to
 * avoid is under-quoting corner work, and at 15% the corner pass on the OD holder
 * still lands around 100 min, which is a real amount of time rather than a
 * rounding error.
 *
 * The honest position: this number is not measured. It is the one input to the
 * finish model that a shop should challenge first, and the learning loop's
 * "finish" class is what corrects it.
 */
export const CORNER_AREA_FRACTION = 0.15;

/** Biggest cutter anyone finishes mold-plate walls with. */
export const FINISH_MAX_DIA_IN = 0.5;

/** Finishing stepover as a share of cutter diameter. Matches the engine's own ~0.17·D. */
export const FINISH_STEPOVER_FRACTION = 0.17;

/**
 * Finish feed, in/min. MEASURED from the reference program by gcode_calibrate:
 * T174 0.499" at 6.6 ipm, T12/T13 0.375" at 7.6-8.1 ipm. Near enough constant
 * across the finishing sizes to use one figure.
 */
export const FINISH_FEED_IPM = 6.6;

/**
 * Share of the engine's reported surface area that actually gets a finish pass.
 *
 * ⚠ THIS IS THE REAL BUG, and it is not a tool-choice problem.
 *
 * The engine reports ~3000 cm² of area to finish on a pot block — 4.7 square
 * feet. Working backwards from the reference program, which spends a measured
 * 62.8 min on semi-finish plus finish, at the measured 6.6 ipm and a 0.17·D
 * stepover on a 0.5" cutter, the area that really gets finished is about
 * 227 cm². The engine is using TOTAL MESH SURFACE — every facet, including the
 * outside of the block, the drilled holes and the faces that only ever get
 * faced — as though all of it needed a finish raster.
 *
 * 227 / 3000 = 0.076 for a single pass. But the engine emits a SEMI-finish op and
 * a FINISH op, and this function is applied to both, so the fraction has to cover
 * the pair: 0.076 / 1.42 = 0.053 puts semi + finish together on the measured
 * 62.8 min rather than 42% over it. Getting this wrong is easy and quiet — each
 * op looked individually reasonable while the total was half again too high.
 *
 * No amount of picking a better cutter fixes a 13x area error, which is why the
 * first attempt at this still left the OD holder 5.8x the ID holder: it scaled
 * off the engine's minutes, and those minutes carry the area error. Pricing from
 * the area directly, with this fraction applied, is what makes two halves of the
 * same tool come out at comparable times.
 *
 * Anchored on one measured part, so it is a prime candidate for the learning
 * loop's "finish" class to refine.
 */
export const FINISH_AREA_FRACTION = 0.053;

/** Tools that are not finishers whatever their size. Bull-nose is for boring. */
const NOT_A_FINISHER = /bull-?nose|dovetail|snap ring|grease|raptor/i;

export interface FinishStep {
  diaIn: number;
  tool: ShopTool;
  min: number;
  role: "bulk" | "corner";
  areaFraction: number;
  reachShort: boolean;
}

export interface FinishResult {
  steps: FinishStep[];
  totalMin: number;
  /** What the engine charged for the same area with one small cutter. */
  singleToolMin: number;
  singleToolDiaIn: number;
  passthrough: boolean;
  note: string;
}

export interface FinishInput {
  /** Engine's minutes for this finish op. Reported for contrast, NOT scaled from. */
  singleToolMin: number;
  /** Diameter the engine chose, INCHES. */
  singleToolDiaIn: number;
  /** Engine's reported surface area, cm². This is what actually drives the time. */
  areaCm2: number;
  /** Smallest internal corner radius, INCHES, or null. */
  minCornerRadIn: number | null;
  /** Part footprint [long, short], INCHES. */
  footprintIn: [number, number];
  /** Deepest wall, INCHES. */
  depthIn: number;
  /** Semi-finish rather than final finish — same geometry, coarser stepover. */
  semi?: boolean;
}

/**
 * Split one finish operation into a bulk pass plus a corner pass.
 *
 * TIME MODEL. For a raster finish, time = area / (stepover x feed), and stepover
 * scales with cutter diameter — the engine itself uses ~0.17 x D (1.05 mm on a
 * 0.25" tool, 0.60 mm on a 0.123"). So for a fixed area, time is inversely
 * proportional to diameter:
 *
 *     bulk_min   = engine_min x (D_engine / D_bulk) x (1 - cornerFraction)
 *     corner_min = engine_min x cornerFraction          [small tool, unchanged]
 *
 * Only the tool choice changes; the engine's area measurement and feed model are
 * kept. That keeps this a targeted fix rather than a second, competing estimate.
 */
export function planFinishPasses(input: FinishInput): FinishResult {
  const { singleToolMin, singleToolDiaIn, areaCm2, minCornerRadIn, footprintIn, depthIn, semi } =
    input;

  const bail = (note: string): FinishResult => ({
    steps: [],
    totalMin: singleToolMin,
    singleToolMin,
    singleToolDiaIn,
    passthrough: true,
    note,
  });

  if (!(areaCm2 > 0)) {
    return bail("No usable surface area reported, so the engine's figure stands.");
  }

  const width = Math.min(footprintIn[0], footprintIn[1]);
  const maxByPart = width * MAX_TOOL_FRACTION_OF_WIDTH;

  // Largest real finisher in the sensible size band that fits the part.
  const inBand = END_MILLS.filter(
    (t) =>
      !NOT_A_FINISHER.test(t.name) &&
      t.dia >= MIN_FINISH_DIA_IN - 1e-6 &&
      t.dia <= Math.min(FINISH_MAX_DIA_IN, maxByPart) + 1e-6,
  ).sort((a, b) => b.dia - a.dia);

  if (inBand.length === 0) {
    return bail(
      `No endmill between Ø${MIN_FINISH_DIA_IN}" and Ø${FINISH_MAX_DIA_IN}" fits a ` +
        `${width.toFixed(2)}" wide part, so the engine's figure stands.`,
    );
  }

  // Prefer one that reaches; otherwise take the largest and FLAG the reach.
  //
  // Refusing outright would mean keeping the engine's number, and that number is
  // known to be ~13x too high because of the area error. A correctly-priced
  // estimate with "this tool does not reach" attached is more useful than a
  // silently wrong one. Note the reach bar here is the part's full depth, which
  // over-states what a WALL finish pass actually has to reach into.
  const reaching = inBand.filter((t) => Math.max(t.depth, t.oal - 0.5) >= depthIn - 1e-6);
  const bulk = reaching[0] ?? inBand[0];
  const bulkReachShort = reaching.length === 0;
  // The corner tool is the one the engine sized to the tightest radius. That
  // choice is right for the corners; it was only ever wrong for the whole area.
  const cornerDia = singleToolDiaIn > 0 ? singleToolDiaIn : bulk.dia;
  const cornerTool = cribToolFor(cornerDia, depthIn);

  // Area that genuinely gets a finish pass, in mm².
  const finishAreaMm2 = areaCm2 * 100 * FINISH_AREA_FRACTION;

  // SEMI-FINISH RUNS FASTER THAN FINISH, NOT JUST COARSER.
  //
  // This used one feed for both passes and separated them only by stepover. The
  // ID Holder programs measure the same tools in both roles and they are not the
  // same speed: 0.75" solid at 8.3 ipm semi against 5.6 finish, 0.5" Dapra at
  // 7.3 against 6.8, 1" solid at 9.1. Weighted, semi runs 1.35x faster -- which
  // makes sense, a semi pass is not holding a finish tolerance.
  //
  // See SEMI_FINISH_SPEEDUP in lib/measuredRates.ts for the measurements.
  const feedMmMin = FINISH_FEED_IPM * (semi ? SEMI_FINISH_SPEEDUP : 1) * 25.4;
  // Semi-finish runs a coarser stepover, so roughly half the passes.
  const stepFactor = FINISH_STEPOVER_FRACTION * (semi ? 2 : 1);

  const minutesFor = (diaIn: number, areaMm2: number) => {
    const stepoverMm = stepFactor * diaIn * 25.4;
    if (!(stepoverMm > 0)) return 0;
    return areaMm2 / (stepoverMm * feedMmMin);
  };

  const bulkArea = finishAreaMm2 * (1 - CORNER_AREA_FRACTION);
  const cornerArea = finishAreaMm2 * CORNER_AREA_FRACTION;

  const steps: FinishStep[] = [
    {
      diaIn: bulk.dia,
      tool: bulk,
      min: Math.round(minutesFor(bulk.dia, bulkArea) * 10) / 10,
      role: "bulk",
      areaFraction: 1 - CORNER_AREA_FRACTION,
      reachShort: bulkReachShort,
    },
    {
      diaIn: cornerDia,
      tool: cornerTool.tool,
      min: Math.round(minutesFor(cornerDia, cornerArea) * 10) / 10,
      role: "corner",
      areaFraction: CORNER_AREA_FRACTION,
      reachShort: cornerTool.reachShort,
    },
  ];

  const total = steps.reduce((s, x) => s + x.min, 0);
  const ratio = total > 0 ? singleToolMin / total : 1;

  const note =
    `${(areaCm2 * FINISH_AREA_FRACTION).toFixed(0)} cm² of wall and floor gets a finish ` +
    `pass (${Math.round(100 * FINISH_AREA_FRACTION)}% of the ${areaCm2.toFixed(0)} cm² total ` +
    `mesh surface — the rest is outside faces, drilled holes and faced surfaces that never ` +
    `see a finish raster). ` +
    `${Math.round(100 * (1 - CORNER_AREA_FRACTION))}% of it with Ø${bulk.dia}" (T${bulk.num}) ` +
    `at a ${(stepFactor * bulk.dia * 25.4).toFixed(2)} mm stepover and a measured ` +
    `${FINISH_FEED_IPM} ipm; the remaining ${Math.round(100 * CORNER_AREA_FRACTION)}% of ` +
    `corner blends with Ø${cornerDia.toFixed(3)}". ` +
    (minCornerRadIn
      ? `The Ø${(2 * minCornerRadIn).toFixed(3)}" internal corner sets the SMALL tool only. `
      : "") +
    (singleToolMin > 0
      ? `The engine charged ${singleToolMin.toFixed(0)} min by rastering the FULL mesh ` +
        `surface with the corner-sized cutter, ${ratio.toFixed(1)}x this.`
      : "");

  return {
    steps,
    totalMin: total,
    singleToolMin,
    singleToolDiaIn,
    passthrough: false,
    note,
  };
}

export interface CascadeInput {
  /** Datum's roughing minutes, for the before/after contrast only. */
  singleToolMin: number;
  /** Datum's rougher diameter, INCHES. Reported, not used to scale. */
  singleToolDiaIn: number;
  /** Volume to clear, CUBIC INCHES. This is what actually drives the time. */
  removeVolIn3: number;
  /** Smallest internal corner radius, INCHES, or null. */
  minCornerRadIn: number | null;
  /** Part footprint [long, short], INCHES — caps the biggest usable cutter. */
  footprintIn: [number, number];
  /** Deepest cut, INCHES. */
  depthIn: number;
  /** Learned multiplier on TOOLPATH_EFFICIENCY, from this shop's own actuals. */
  learnedRoughFactor?: number;
}

/**
 * Turn one roughing operation into the cascade a machinist would run.
 *
 * Falls through to Datum's own number (`passthrough: true`) when the volume is
 * unusable, so a bad mesh cannot be laundered into a confident-looking cascade.
 */
export function planRoughingCascade(input: CascadeInput): CascadeResult {
  const {
    singleToolMin,
    singleToolDiaIn,
    removeVolIn3,
    minCornerRadIn,
    footprintIn,
    depthIn,
    learnedRoughFactor,
  } = input;

  const cornerDia = minCornerRadIn && minCornerRadIn > 0 ? 2 * minCornerRadIn : Infinity;
  const wasCornerLimited = cornerDia < MIN_ROUGHING_DIA_IN;

  const bail = (note: string): CascadeResult => ({
    steps: [],
    totalMin: singleToolMin,
    singleToolMin,
    singleToolDiaIn,
    removeVolIn3,
    wasCornerLimited,
    passthrough: true,
    note,
  });

  if (!(removeVolIn3 > 0)) {
    return bail(
      "No usable removal volume, so the roughing cascade was not applied and the " +
        "engine's own figure stands.",
    );
  }

  const width = Math.min(footprintIn[0], footprintIn[1]);
  const maxByPart = width * MAX_TOOL_FRACTION_OF_WIDTH;

  // Roughing stops at MIN_ROUGHING_DIA_IN regardless of how tight the corner is.
  const smallestNeeded = Math.max(Math.min(cornerDia, MIN_ROUGHING_DIA_IN), MIN_ROUGHING_DIA_IN);

  let ladder = ROUGHING_LADDER.filter(
    (t) => t.diaIn <= maxByPart + 1e-6 && t.diaIn >= smallestNeeded - 1e-6,
  );

  // REACH EXCLUDES A TOOL, it does not merely annotate it.
  //
  // This is what makes the model pick the tools the shop picks. On the reference
  // pot the bore is 6.7" deep; the 2" Zenit reaches 5.5", so it cannot be the
  // bulk rougher there — and indeed the program roughs that bore with the 1" and
  // finishes it by helical boring with a long bull-nose. An earlier version left
  // the 2" in the ladder and just printed REACH beside it, which gave the big
  // cutter 47% of the volume it could not physically get to.
  const reachable = ladder.filter((t) => {
    const c = cribToolFor(t.diaIn, depthIn);
    return c.tool !== null && !c.reachShort;
  });

  let reachNote = "";
  if (reachable.length > 0 && reachable.length < ladder.length) {
    const dropped = ladder
      .filter((t) => !reachable.includes(t))
      .map((t) => `${t.diaIn}"`)
      .join(", ");
    reachNote =
      `${dropped} excluded — nothing that size in the crib reaches ` +
      `${depthIn.toFixed(2)}" deep. `;
    ladder = reachable;
  } else if (reachable.length === 0 && ladder.length > 0) {
    // Nothing reaches at all. Keep the deepest-reaching option so a number still
    // comes out, but say plainly that this is not a roughing problem.
    const deepest = [...ladder].sort(
      (a, b) => (cribToolFor(b.diaIn, 0).tool?.depth ?? 0) - (cribToolFor(a.diaIn, 0).tool?.depth ?? 0),
    );
    reachNote =
      `NO rougher in the crib reaches ${depthIn.toFixed(2)}" deep. A feature this ` +
      `deep is bored or plunged, not cascade-roughed — treat this figure as a ` +
      `placeholder and check the operation plan. `;
    ladder = [deepest[0]];
  }

  // Narrow part: nothing on the ladder fits at roughing size. Use the largest
  // that does fit, even if it is below MIN_ROUGHING_DIA_IN — a 1" wide rail is
  // still roughed by something.
  if (ladder.length === 0) {
    const fits = ROUGHING_LADDER.filter((t) => t.diaIn <= maxByPart + 1e-6);
    if (fits.length === 0) {
      return bail(
        `Part is only ${width.toFixed(2)}" across — smaller than any rougher in the ` +
          `ladder, so the engine's figure stands.`,
      );
    }
    ladder = [fits[0]];
  }

  // DIVIDED, not multiplied. A learned factor is actual/quoted, so 1.4 means
  // jobs really take 40% LONGER than quoted — which means the shop achieves LESS
  // engagement than modelled, so efficiency must come DOWN and time must go UP.
  // Multiplying here would have made every correction push the estimate the
  // wrong way, and the feedback loop would have diverged instead of converged.
  const learn = learnedRoughFactor && learnedRoughFactor > 0 ? learnedRoughFactor : 1;
  const efficiency = TOOLPATH_EFFICIENCY / learn;
  const weight = (d: number) => Math.pow(d, VOLUME_EXPONENT);
  const sumW = ladder.reduce((s, t) => s + weight(t.diaIn), 0);

  const steps: CascadeStep[] = ladder.map((t, i) => {
    const frac = weight(t.diaIn) / sumW;
    const vol = removeVolIn3 * frac;
    const mrr = idealMrr(t) * efficiency;
    const { tool, reachShort } = cribToolFor(t.diaIn, depthIn);
    return {
      diaIn: t.diaIn,
      tool,
      min: Math.round((vol / Math.max(mrr, 1e-6)) * 10) / 10,
      volFraction: frac,
      volIn3: Math.round(vol * 100) / 100,
      mrr: Math.round(mrr * 100) / 100,
      strategy: t.strategy,
      role: i === 0 ? "bulk" : "rest",
      reachShort,
    };
  });

  const totalMin = steps.reduce((s, x) => s + x.min, 0);
  const ratio = totalMin > 0 ? singleToolMin / totalMin : 1;

  const note =
    reachNote +
    `${removeVolIn3.toFixed(1)} in³ roughed with ${ladder.length} cutter` +
    `${ladder.length === 1 ? "" : "s"}, ` +
    `${ladder[0].diaIn}"${ladder.length > 1 ? ` down to ${ladder[ladder.length - 1].diaIn}"` : ""}. ` +
    `The biggest takes ${Math.round(100 * steps[0].volFraction)}% of the volume at ` +
    `${steps[0].mrr.toFixed(1)} in³/min; each smaller one clears only the rest material ` +
    `the last could not reach. ` +
    (wasCornerLimited
      ? `The Ø${cornerDia.toFixed(3)}" internal corner is left to the finisher, which ` +
        `removes a skin — it is not a roughing tool. `
      : "") +
    (singleToolMin > 0
      ? `The engine charged ${singleToolMin.toFixed(0)} min for the same metal with one ` +
        `Ø${singleToolDiaIn.toFixed(3)}" cutter, ${ratio.toFixed(1)}x this. `
      : "") +
    `Rates are ae·ap·vf per tool, derated x${efficiency.toFixed(2)} for real toolpath ` +
    `engagement.`;

  return {
    steps,
    totalMin,
    singleToolMin,
    singleToolDiaIn,
    removeVolIn3,
    wasCornerLimited,
    passthrough: false,
    note,
  };
}
