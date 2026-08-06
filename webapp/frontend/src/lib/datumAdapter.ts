/**
 * Datum engine -> CMS mold-base adapter.
 *
 * Runs the Datum feature-recognition engine over the SIX per-plate STLs that
 * Module6121 exports (ExportPlateStlsForComparison), and returns the result in
 * the same shape the MachiningPanel already renders.
 *
 * WHY THIS IS BETTER THAN THE PYTHON ESTIMATOR
 *   machining.py guesses. Finished volume is mass/density or "82% of stock",
 *   and hole counts come from a per-role pattern table because the CAD export
 *   carries no hole data. Datum measures: it welds the mesh, segments it into
 *   surfaces, fits cylinders, and finds the actual holes -- diameter, depth,
 *   through vs blind by ray cast, tapped vs clearance, counterbores, pockets.
 *   Every number that used to be "assumed" becomes "measured".
 *
 * THE FOUR THINGS THAT HAD TO BE RECONCILED
 *
 *   1. UNITS. Datum is millimetre-native throughout -- stock in mm, an.vol in
 *      mm^3, rho in g/cm^3, vc in m/min, ALLOW = 3.0 mm. Our pipeline is
 *      inches: Module6121 now forces inch STL export, and the viewer normalises
 *      to inches. So every mesh is scaled to mm on the way in and every result
 *      is converted back on the way out. Get this wrong by 25.4 and cycle time
 *      is wrong by 25.4^3.
 *
 *   2. STOCK. Datum's default is bounding box + allowance, because it cannot
 *      know what is in your rack. We DO know -- the quote row carries the real
 *      BOM stock size. So we pass mode "custom" with the actual T/W/L, which is
 *      strictly more accurate and makes facing time real rather than notional.
 *      HANDOFF measured this as a big lever: "+3 mm block -> 6:57 facing,
 *      4in plate -> 18:54 facing and $43 more per part."
 *
 *   3. MATERIALS. Datum ships 6061/7075/brass/1018/4140/304/17-4/Ti/Inconel.
 *      A mold shop runs 4140, 4130, P20, A-2, D-2, H13, A36, 420SS. 4140 is
 *      already there; the rest map to the closest match by machinability and
 *      the substitution is reported, never silent.
 *
 *   4. ONE PART vs SIX. Datum costs a single part and lands setup and
 *      programming in full -- deliberately, per HANDOFF. For a mold base that
 *      is actually right: each plate is its own program and its own setup. But
 *      it means programming is charged six times, so it is broken out as its
 *      own line rather than buried, because that is the number a shop will
 *      want to argue with first.
 */

// Types come from datumEngine.d.ts, so the engine surface is genuinely checked
// rather than silently `any`. Copy the implementation in with
// install_datum_engine.py -- tsc resolves the .d.ts either way, but the bundle
// needs the real file.
import * as Datum from "./datumEngine.js";
import type { DatumFeatures, DatumPlan } from "./datumEngine.js";

import { classifyComponentFilename, MACHINED_PART_KINDS, type ComponentKind } from "./componentKind";
import { planFinishPasses, planRoughingCascade } from "./roughingCascade";
import { plateConfidence, rollUpConfidence, type ConfidenceBand } from "./quoteConfidence";
import { MM_PER_INCH } from "./stlGeometry";
import {
  BALL_MILLS,
  CHAMFER_MILLS,
  DATUM_TOL_INDEX_FOR_SHOP,
  DRILLS,
  END_MILLS,
  FACE_MILLS,
  METRIC_TO_UNC,
  SPOT_DRILLS,
  TAPS,
  bestDrill,
  bestTool,
  drillExists,
  drillNear,
  finishFaceMill,
  impliedTapFromDrill,
  pickDrill,
  pickTool,
  toolLabel,
  type ShopTool,
} from "./shopTools";
import {
  DEBURR_ON_MACHINE,
  REFERENCE_JOB,
  setupMinutes,
  timeFactorFor,
} from "./shopCalibration";
import type {
  AssetRef,
  GeoBore,
  GeometryPart,
  GeometryReport,
  GeoHole,
  GeoPocket,
  MachiningEstimate,
  MachiningPart,
  QuoteLineItem,
} from "../api/client";

/**
 * Every part kind Module6121 exports an STL for, in stack order.
 *
 * This used to be the six BMS pot-block roles, hardcoded. That silently capped
 * the whole Geometry tab and machining estimate to pot-block jobs: a standard
 * or PCS base classified its plates correctly and then had every one of them
 * filtered out here, so an A/B/T/AX/5X/6X job reported no holes, no pockets
 * and no machining time. Now it comes from KIND_ORDER, so a kind added to the
 * classifier is picked up automatically.
 */
export const PLATE_KINDS: ComponentKind[] = MACHINED_PART_KINDS;

/**
 * Shop material -> nearest Datum material, with the substitution recorded.
 *
 * Matched on machinability first, not on name. P20 and 4140 cut close enough to
 * the same that sharing a rate table is defensible; A-2 and D-2 are harder and
 * sit nearer stainless than alloy steel.
 */
const MATERIAL_MAP: Record<string, { datum: string; exact: boolean; note?: string }> = {
  "4140": { datum: "4140 PH", exact: true },
  "4130": { datum: "4140 PH", exact: false, note: "4130 rated as 4140 (near-identical machinability)" },
  P20: { datum: "4140 PH", exact: false, note: "P20 rated as 4140 (both pre-hard ~30 Rc)" },
  A36: { datum: "1018", exact: false, note: "A36 rated as 1018 mild steel" },
  "1030": { datum: "1018", exact: false, note: "1030 rated as 1018 mild steel" },
  "A-2": { datum: "304", exact: false, note: "A-2 tool steel rated as 304 (closest k)" },
  "D-2": { datum: "17-4 PH H900", exact: false, note: "D-2 rated as 17-4 PH (closest k)" },
  H13: { datum: "304", exact: false, note: "H13 rated as 304 (closest k)" },
  S7: { datum: "304", exact: false, note: "S7 rated as 304 (closest k)" },
  STAINLESS: { datum: "304", exact: true },
  "420SS": { datum: "304", exact: false, note: "420SS rated as 304" },
  "6061": { datum: "6061-T6", exact: true },
};

function mapMaterial(raw: string): { datum: string; exact: boolean; note?: string } {
  const s = (raw || "").toUpperCase();
  for (const key of Object.keys(MATERIAL_MAP)) {
    if (s.includes(key)) return MATERIAL_MAP[key];
  }
  // Everything in a mold base is some flavour of tool steel. 4140 is the safest
  // default and by far the most common.
  return { datum: "4140 PH", exact: false, note: `"${raw || "unknown"}" not recognised — rated as 4140` };
}

/**
 * Tolerance band.
 *
 * The shop holds ±.001 in on most features. That is ±0.0254 mm, i.e. Datum's
 * "Precision" band (±0.025), not the "Close" band (±0.05) this used to default
 * to. One band too loose under-counts finish passes, inspection minutes and
 * scrap risk on every plate.
 */
const DEFAULT_TOL_INDEX = DATUM_TOL_INDEX_FOR_SHOP;

/**
 * Rewrite Datum's invented tooling as the real crib, and report what the crib
 * cannot do.
 *
 * Datum sizes cutters from continuous maths, so it happily calls out a "Ø7
 * 4-flute" or a "Ø14.29 carbide drill". Neither exists here. Snapping to actual
 * tool numbers is what makes the op list auditable at the machine -- and the
 * gaps it surfaces (a hole with no drill, a pocket deeper than anything reaches,
 * a metric tap in a UNC-only crib) are the most useful output of the whole pass.
 */
function applyShopTools(
  ops: Array<{ op: string; min: number; measured: boolean; detail: string }>,
): { ops: typeof ops; gaps: string[] } {
  const gaps: string[] = [];
  const mm2in = (mm: number) => mm / MM_PER_INCH;

  const label = (t: ShopTool | null, fallback: string) =>
    t ? `T${t.num} ${t.name}` : fallback;

  const out = ops.map((o) => {
    const name = o.op;

    // Datum embeds the diameter it chose in the op label / detail as "Ø<mm>".
    const dMatch = /Ø\s*([\d.]+)/.exec(name) || /Ø\s*([\d.]+)/.exec(o.detail);
    const dIn = dMatch ? mm2in(parseFloat(dMatch[1])) : 0;

    // ---- drilling ----------------------------------------------------
    if (/^Drill/i.test(name) && dIn > 0) {
      const depthMatch = /([\d.]+)\s*mm\b/.exec(o.detail);
      const depthIn = depthMatch ? mm2in(parseFloat(depthMatch[1])) : 0;
      const drill = pickDrill(dIn, depthIn);
      if (drill) return { ...o, detail: `${label(drill, "")} · ${o.detail}` };

      // Right size, wrong reach — a different problem with a different answer.
      const anySize = drillExists(dIn);
      if (anySize) {
        gaps.push(
          `Ø${dIn.toFixed(3)}" x ${depthIn.toFixed(2)}" deep (${(depthIn / Math.max(dIn, 0.001)).toFixed(1)}xD): ` +
            `T${anySize.num} is the right diameter but only reaches ${anySize.depth.toFixed(2)}". ` +
            `Needs an extension, a step drill, or gun drilling.`,
        );
        return {
          ...o,
          op: `${name} — NO REACH`,
          detail: `T${anySize.num} is Ø${anySize.dia.toFixed(3)}" but only ${anySize.depth.toFixed(2)}" deep · ${o.detail}`,
        };
      }

      // No drill that size. If the hole is a tap-drill diameter, say which tap
      // it is for and which drill the crib actually stocks -- shops routinely
      // run a slightly oversize tap drill in 4140 to cut tap breakage, so
      // "NO TOOL — needs Ø0.201"" is technically true and practically useless
      // when T3 (Ø0.213") is sitting in the carousel for exactly this thread.
      const tap = impliedTapFromDrill(dIn);
      if (tap) {
        const stocked = DRILLS.filter(
          (d) => d.dia >= dIn - 0.001 && d.dia <= dIn + 0.02 && d.depth >= depthIn,
        ).sort((a, b) => a.dia - b.dia)[0];
        if (stocked) {
          gaps.push(
            `Ø${dIn.toFixed(3)}" is the ${tap.spec} tap drill. The crib stocks ` +
              `T${stocked.num} at Ø${stocked.dia.toFixed(3)}" — ` +
              `${((stocked.dia - dIn) * 1000).toFixed(0)} thou over nominal, which is normal ` +
              `practice in 4140 but gives a looser thread. Confirm that is intended.`,
          );
          return {
            ...o,
            detail: `T${stocked.num} ${stocked.name} (${tap.spec} tap drill, ` +
              `${((stocked.dia - dIn) * 1000).toFixed(0)} thou over) · ${o.detail}`,
          };
        }
      }

      // Not on THIS machine's list, so it runs on the machine that has it, at
      // the size the hole actually needs. Priced, not refused -- see bestDrill.
      // An oversize substitution is still never made: that would be a different
      // hole, which is why the tool is synthesised at nominal rather than
      // snapped to whatever is closest here.
      const elsewhere = bestDrill(dIn, depthIn);
      return {
        ...o,
        detail: `${toolLabel(elsewhere)} · ${o.detail}`,
      };
    }

    // ---- tapping -----------------------------------------------------
    if (/^Tap/i.test(name)) {
      // Datum names the THREAD in the op label: "Tap M12 x5", "Tap 1/2-13 x4".
      // Validate that callout against the taps actually in the crib.
      const metric = /\b(M\d+)/i.exec(name);
      if (metric) {
        // RE-READ it as the UNC equivalent rather than flagging it and stopping.
        //
        // The reference program settles this: J8441 ID POT has exactly two taps,
        // T31 1/4-20 UNC and T30 1/2-13 UNC. There is no M6 or M10 anywhere in
        // it. Datum matched Ø5.11 mm against the metric table (M6 tap drill is
        // 5.0 mm) when it is really 0.2012 in, the 1/4-20 tap drill, to four
        // decimal places. On a UNC-only shop the metric reading is always the
        // misread, so translating is more useful than refusing.
        const unc = METRIC_TO_UNC[metric[1].toUpperCase()];
        const owned = unc ? TAPS.find((t) => t.spec.startsWith(unc)) : null;

        if (owned) {
          return {
            ...o,
            op: name.replace(metric[1], owned.spec.replace(" UNC", "")),
            detail:
              `T${owned.num} ${owned.spec} · tap drill Ø${owned.tapDrill.toFixed(3)}" · ` +
              `re-read from ${metric[1]} — this crib is UNC only and the hole size ` +
              `matches the UNC tap drill · ${o.detail}`,
          };
        }

        gaps.push(
          `${name} — no metric taps in this crib, and no clean UNC equivalent. ` +
            `Available: ${TAPS.map((t) => t.spec.replace(" UNC", "")).join(", ")}.`,
        );
        return {
          ...o,
          op: `${name} — NOT IN CRIB`,
          detail: `no metric taps here · ${o.detail}`,
        };
      }

      // Imperial, and a size this shop plausibly runs even when it is not on
      // THIS machine's rack. A UNC tap is a stock item; the metric branch above
      // is different and stays a warning, because there the finding is that the
      // mesh READING is wrong, not that a drawer is empty.
      const spec = /Tap\s+([\d/]+-\d+)/i.exec(name)?.[1];
      const owned = spec ? TAPS.find((t) => t.spec.startsWith(spec)) : null;
      if (spec && !owned) {
        return { ...o, detail: `${spec} tap (another machine) · ${o.detail}` };
      }
      if (owned) {
        return {
          ...o,
          detail: `T${owned.num} ${owned.spec} · tap drill Ø${owned.tapDrill.toFixed(3)}" · ${o.detail}`,
        };
      }
      return o;
    }

    // ---- facing ------------------------------------------------------
    // ALWAYS the 2" face mill. This used to take the widest face mill on the
    // list (the 3" finisher) while datumEngine timed the pass at 63 mm, so the
    // callout, the maths and the machine were three different cutters.
    if (/^Face/i.test(name)) {
      return { ...o, detail: `${toolLabel(finishFaceMill())} · ${o.detail}` };
    }

    // ---- roughing / finishing / semi-finish --------------------------
    if (/Rough|finish walls|Semi-finish/i.test(name) && dIn > 0) {
      const em = bestTool(END_MILLS, dIn, 0, "endmill");
      return { ...o, detail: `${toolLabel(em)} · ${o.detail}` };
    }

    // ---- ball nose ---------------------------------------------------
    if (/Contour/i.test(name) && dIn > 0) {
      const bm = bestTool(BALL_MILLS, dIn, 0, "ball nose");
      return { ...o, detail: `${toolLabel(bm)} · ${o.detail}` };
    }

    // ---- chamfer -----------------------------------------------------
    if (/Chamfer/i.test(name)) {
      const cm = pickTool(CHAMFER_MILLS, Infinity, 0);
      return { ...o, detail: `${label(cm, "")} · ${o.detail}` };
    }

    // ---- spotting ----------------------------------------------------
    if (/Spot drill/i.test(name)) {
      const sd = pickTool(SPOT_DRILLS, Infinity, 0);
      return { ...o, detail: `${label(sd, "")} · ${o.detail}` };
    }

    // ---- deburr ------------------------------------------------------
    // Datum books this as bench work running parallel to the next part, so it
    // costs nothing in cycle. Op 27 of the reference program is a real DEBURR
    // cycle with T10 on the spindle -- 11.7 min of machine time. Some hand work
    // still happens at the bench, but the spindle time is real.
    if (/Deburr/i.test(name) && DEBURR_ON_MACHINE) {
      const dt = pickTool(END_MILLS, 0.25, 0);
      return {
        ...o,
        detail:
          `${label(dt, "")} · ON THE MACHINE — the reference program runs a ` +
          `deburr cycle with a tool in the spindle, so this is cycle time, not ` +
          `parallel bench work · ${o.detail}`,
      };
    }

    // ---- reaming: only real if a drill that size exists ---------------
    if (/^Ream/i.test(name)) {
      const sizes = [...o.detail.matchAll(/Ø([\d.]+)/g)].map((m) => mm2in(parseFloat(m[1])));
      const missing = sizes.filter((s) => !drillNear(s));
      if (missing.length > 0) {
        gaps.push(
          `Ream called out at Ø${missing.map((s) => s.toFixed(3)).join(", Ø")}" but there is ` +
            `no matching drill in the crib, so the hole cannot be pre-drilled to size.`,
        );
      }
      return o;
    }

    return o;
  });

  return { ops: out, gaps };
}

type AdapterOp = {
  op: string;
  min: number;
  measured: boolean;
  detail: string;
  /**
   * Already priced from real cutting parameters, so the blanket per-class
   * calibration factor must NOT be applied on top.
   *
   * Cascade steps are named "Rough 1/4 …", which matches the /^rough/ test in
   * timeFactorFor and would pick up the x0.7 speed correction a second time. But
   * that factor exists to fix the generic engine using ONE surface speed for
   * everything, and the cascade already prices each tool at its own ae·ap·vf
   * derated to measured reality. Applying both would knock 30% off a number that
   * already reproduces the reference program to 2.5%.
   */
  preCalibrated?: boolean;
};

/**
 * Replace Datum's single roughing operation with the cascade the shop runs.
 *
 * Datum emits one "Rough" op whose cutter has been sized down to the tightest
 * internal corner, then charges the whole bulk removal to it. See
 * roughingCascade.ts for why that is wrong and what the two reference programs
 * actually do.
 *
 * Non-roughing ops pass through untouched.
 */
function expandRoughingCascade(
  ops: AdapterOp[],
  pl: DatumPlan,
  fx: DatumFeatures,
  an: { area?: number } | null,
  removeVolIn3: number,
  notes: string[],
): AdapterOp[] {
  const out: AdapterOp[] = [];

  // Footprint = the two largest extents; depth = the deepest pocket, falling
  // back to the smallest extent (a through feature on a plate).
  const extIn = [...fx.ext].map((v) => v / MM_PER_INCH).sort((a, b) => b - a);
  const footprintIn: [number, number] = [extIn[0] || 1, extIn[1] || 1];
  const deepestPocketMm = (fx.pockets || []).reduce(
    (m, p) => Math.max(m, Number(p.depth) || 0),
    0,
  );
  const depthIn = (deepestPocketMm || fx.ext[2] || 0) / MM_PER_INCH;

  const minCornerRadIn =
    typeof fx.minInternalR === "number" && fx.minInternalR > 0
      ? fx.minInternalR / MM_PER_INCH
      : null;

  // ONE cascade for the whole part, not one per roughing operation.
  //
  // The engine can emit several roughing ops (a pocket rough and a profile
  // rough, say). Expanding each of them into a full cascade charged the ENTIRE
  // part's removal volume to each one, so a part with two roughing ops
  // double-counted every cubic inch and every minute. On the ID HOLDER that
  // showed up as the four cascade steps printed twice, ~47 min of invented
  // roughing on one plate.
  //
  // Roughing is one job with one volume, so the cascade replaces the whole set:
  // the first roughing op becomes the cascade, the rest are dropped.
  const roughIdx = ops
    .map((o, i) => ({ o, i }))
    .filter(({ o }) => /^Rough/i.test(o.op) && o.min > 0)
    .map(({ i }) => i);
  const cascadeAt = roughIdx.length > 0 ? roughIdx[0] : -1;
  const collapsed = roughIdx.length - 1;

  for (let i = 0; i < ops.length; i++) {
    const o = ops[i];
    // Only true roughing. "Semi-finish" and "finish walls" are skin cuts with
    // their own tools and must not be cascaded.
    if (i !== cascadeAt) {
      // A second/third roughing op is folded into the single cascade above, so
      // it must not also appear on its own.
      if (roughIdx.includes(i)) continue;

      // Finishing gets the same treatment as roughing, for the same reason: a
      // tight internal corner must not choose the tool for a full-area pass.
      // Left unfixed this was the single worst number in the estimate — 698 min
      // of finishing on the OD holder against 95 on the ID holder.
      if (/finish walls|semi-finish/i.test(o.op) && o.min > 0) {
        const fm = /Ø\s*([\d.]+)/.exec(o.op) || /Ø\s*([\d.]+)/.exec(o.detail);
        const finDia = fm ? parseFloat(fm[1]) / MM_PER_INCH : 0;
        // The engine prints the surface area it is rastering, e.g. "3028 cm²".
        // That number drives the time; its minutes do not, because they carry a
        // ~13x area error. Fall back to the mesh area from the analysis.
        const am = /([\d.]+)\s*cm²/.exec(o.detail);
        const areaCm2 = am
          ? parseFloat(am[1])
          : (Number(an?.area) || 0) / 100; // an.area is mm²
        const fin = planFinishPasses({
          singleToolMin: o.min,
          singleToolDiaIn: finDia,
          areaCm2,
          minCornerRadIn,
          footprintIn,
          depthIn,
          semi: /semi/i.test(o.op),
        });
        if (fin.passthrough) {
          out.push(o);
        } else {
          notes.push(fin.note);
          const label = /semi/i.test(o.op) ? "Semi-finish" : "Finish";
          fin.steps.forEach((s) => {
            out.push({
              op:
                s.role === "bulk"
                  ? `${label} walls and floors — Ø${s.diaIn}"`
                  : `${label} corner blends — Ø${s.diaIn.toFixed(3)}"`,
              min: s.min,
              measured: true,
              preCalibrated: true,
              detail:
                `${toolLabel(s.tool)} · ` +
                `${Math.round(100 * s.areaFraction)}% of wall area · ` +
                `${s.role === "bulk" ? "open walls and floors" : "corners the bulk finisher cannot enter"}` +
                (s.reachShort
                  ? ` · REACH: needs ${depthIn.toFixed(2)}", this tool cuts ${s.tool.depth.toFixed(2)}"`
                  : ""),
            });
          });
        }
        continue;
      }

      out.push(o);
      continue;
    }

    // Sum the engine's roughing minutes across all of its roughing ops, so the
    // before/after comparison in the note is against the real total.
    const singleToolMinTotal = roughIdx.reduce((s, k) => s + ops[k].min, 0);

    const dMatch = /Ø\s*([\d.]+)/.exec(o.op) || /Ø\s*([\d.]+)/.exec(o.detail);
    const singleToolDiaIn = dMatch
      ? parseFloat(dMatch[1]) / MM_PER_INCH
      : (Number(pl.Dr) || 0) / MM_PER_INCH;

    if (!(singleToolDiaIn > 0)) {
      out.push(o);
      continue;
    }

    const cascade = planRoughingCascade({
      singleToolMin: singleToolMinTotal,
      singleToolDiaIn,
      removeVolIn3,
      minCornerRadIn,
      footprintIn,
      depthIn,
    });

    // Volume unusable or part too small — keep the engine's own figure rather
    // than dressing a bad input up as a cascade.
    if (cascade.passthrough) {
      out.push({ ...o, detail: `${o.detail} · ${cascade.note}` });
      continue;
    }

    notes.push(
      collapsed > 0
        ? `${cascade.note} (${collapsed + 1} engine roughing operations merged into one ` +
          `cascade — roughing is one job with one volume, so expanding each separately ` +
          `would charge the same metal ${collapsed + 1} times.)`
        : cascade.note,
    );

    cascade.steps.forEach((s, i) => {
      out.push({
        op: `Rough ${i + 1}/${cascade.steps.length} — Ø${s.diaIn}" ${s.role === "bulk" ? "bulk" : "rest"}`,
        min: s.min,
        measured: true,
        preCalibrated: true,
        detail:
          `${toolLabel(s.tool)} · ` +
          `${s.strategy} · ${s.volIn3} in³ ` +
          `(${Math.round(100 * s.volFraction)}% of roughing) at ${s.mrr} in³/min · ` +
          `${s.role === "bulk" ? "bulk removal" : "rest material the previous cutter could not reach"}` +
          (s.reachShort
            ? ` · REACH: needs ${depthIn.toFixed(2)}" deep, this tool cuts ${s.tool.depth.toFixed(2)}"`
            : ""),
      });
    });
  }

  return out;
}

/**
 * Real tool CHANGES, not distinct tools.
 *
 * Datum's `toolChanges` is a count of distinct tools. On a 3+2 job that badly
 * understates the carousel work, because a tool gets recalled at every
 * orientation it is needed in — the 2" Zenit appears five separate times in the
 * reference program, once per B/C position. Measured on the two reference
 * programs: 48 M06 changes over 24 distinct tools, and 46 over 24. So roughly
 * two changes per tool, scaling with the number of orientations.
 *
 * Modelled as: every tool is loaded once, plus one extra load for each
 * additional orientation that needs it, capped so a 12-orientation part does not
 * imply twelve loads of every cutter.
 */
export const TOOL_CHANGE_MIN = 0.4;

export function realToolChanges(distinctTools: number, orientations: number): number {
  const n = Math.max(0, Math.round(distinctTools || 0));
  const o = Math.max(1, Math.round(orientations || 1));
  if (n === 0) return 0;
  // Not every tool is needed at every orientation. The reference programs give
  // 48/24 = 2.0 changes per tool at 5 orientations, so the share of tools
  // recalled per extra orientation is (2.0 - 1) / (5 - 1) = 0.25.
  const RECALL_SHARE = 0.25;
  return Math.round(n * (1 + RECALL_SHARE * (o - 1)));
}

export interface DatumOptions {
  /** $/hr. Falls through to the Python estimator's rate so both agree. */
  shopRate?: number;
  /** Index into Datum.TOLS. 0 = as-milled … 4 = high precision. */
  tolIndex?: number;
  /** Treat diameter matches to tap-drill sizes as tapped holes. */
  assumeTaps?: boolean;
}

/**
 * Scale a parsed STL to millimetres, in place.
 *
 * Deterministic, not a magnitude guess: compare the mesh's own largest extent
 * against the stock size the BOM says this plate is cut from. A ratio near 1
 * means the file is already inches (Module6121 forces inch export now); a ratio
 * near 25.4 means it is an older millimetre file. Anything else and we bail out
 * rather than silently scale by the wrong factor.
 */
function toMillimetres(
  pos: Float32Array,
  expectedMaxIn: number,
): { ok: boolean; note: string } {
  let min = [Infinity, Infinity, Infinity];
  let max = [-Infinity, -Infinity, -Infinity];
  for (let i = 0; i < pos.length; i += 3) {
    for (let a = 0; a < 3; a++) {
      const v = pos[i + a];
      if (v < min[a]) min[a] = v;
      if (v > max[a]) max[a] = v;
    }
  }
  const extent = Math.max(max[0] - min[0], max[1] - min[1], max[2] - min[2]);
  if (!Number.isFinite(extent) || extent <= 0) {
    return { ok: false, note: "mesh has no measurable extent" };
  }

  // No stock hint. Magnitude alone can only resolve the extremes, because the
  // two unit systems overlap in the middle: a BMS pot block at 6 in and a small
  // plate at 6 mm are both "6". Guessing wrong here is a 25.4x error on every
  // coordinate, so 16,387x on volume -- and it would be reported as a success.
  // Refuse the ambiguous band instead.
  if (!(expectedMaxIn > 0.01)) {
    if (extent > 200) return { ok: true, note: "already mm (by magnitude)" };
    if (extent > 40) {
      return {
        ok: false,
        note:
          `no BOM stock row to check against, and ${extent.toFixed(1)} units is ambiguous ` +
          `(${(extent / MM_PER_INCH).toFixed(2)} in or ${extent.toFixed(0)} mm) — refusing to guess`,
      };
    }
    for (let i = 0; i < pos.length; i++) pos[i] *= MM_PER_INCH;
    return { ok: true, note: "scaled in -> mm (by magnitude)" };
  }

  const ratio = extent / expectedMaxIn;

  if (ratio > 0.5 && ratio < 2.5) {
    for (let i = 0; i < pos.length; i++) pos[i] *= MM_PER_INCH;
    return { ok: true, note: `scaled in -> mm (mesh/stock ratio ${ratio.toFixed(2)})` };
  }
  if (ratio > 12 && ratio < 60) {
    return { ok: true, note: `already mm (mesh/stock ratio ${ratio.toFixed(1)})` };
  }
  return {
    ok: false,
    note: `mesh is ${extent.toFixed(1)} units against ${expectedMaxIn.toFixed(2)} in of stock ` +
      `(ratio ${ratio.toFixed(1)}) — neither inches nor mm, refusing to guess`,
  };
}

/** Find the quote row that belongs to a plate kind. */
/** Every STL for each part kind, in the order the macro wrote them. */
function groupModelsByKind(models: AssetRef[]): Map<ComponentKind, AssetRef[]> {
  const byKind = new Map<ComponentKind, AssetRef[]>();
  for (const m of models) {
    if (!m.name.toLowerCase().endsWith(".stl")) continue;
    const k = classifyComponentFilename(m.name);
    if (!PLATE_KINDS.includes(k)) continue;
    const list = byKind.get(k);
    if (list) list.push(m);
    else byKind.set(k, [m]);
  }
  return byKind;
}

/**
 * Which of a kind's STLs to actually measure.
 *
 * The macro writes one file per physical solid — Rails 1, Rails 2 — while a
 * priced quote row covers all of them at once with a quantity. Measuring every
 * file AND applying the row quantity counts the same steel twice, so a PRICED
 * kind measures one representative file and lets qty do the multiplying.
 *
 * An UNPRICED kind has no row and therefore no quantity to apply, which is
 * exactly the case for the extra steel parts Module6121 now sweeps up. Each of
 * those files is its own distinct solid, so each is measured on its own.
 */
function modelsToAnalyse(found: AssetRef[], row: QuoteLineItem | null): AssetRef[] {
  return row ? found.slice(0, 1) : found;
}

function rowForKind(items: QuoteLineItem[], kind: ComponentKind): QuoteLineItem | null {
  for (const r of items) {
    const label = String(r.role_label || r.component || "");
    if (classifyComponentFilename(label) === kind) return r;
  }
  return null;
}

/**
 * Estimate one plate.
 *
 * Returns a MachiningPart, so the existing panel renders it with no changes.
 * Every operation is flagged `measured: true` -- unlike the Python path, none of
 * this is a pattern guess.
 */
async function estimateOnePlate(
  kind: ComponentKind,
  model: AssetRef,
  row: QuoteLineItem | null,
  opts: DatumOptions,
  toolGaps: string[],
): Promise<MachiningPart> {
  const material = String(row?.material || "");
  const mapped = mapMaterial(material);

  const tIn = Number(row?.thickness) || 0;
  const wIn = Number(row?.width) || 0;
  const lIn = Number(row?.length) || 0;
  const qty = Math.max(1, Number(row?.qty) || 1);
  const haveStock = tIn > 0 && wIn > 0 && lIn > 0;
  const expectedMaxIn = Math.max(tIn, wIn, lIn);

  const base: MachiningPart = {
    component: model.name,
    role_label: kind,
    material: mapped.datum,
    qty,
    skipped: true,
    total_min: 0,
    total_hours: 0,
  };

  let pos: Float32Array;
  try {
    const res = await fetch(model.url);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    pos = Datum.parseSTL(await res.arrayBuffer());
  } catch (e) {
    return { ...base, reason: `Could not read STL: ${e instanceof Error ? e.message : "unknown"}` };
  }

  const units = toMillimetres(pos, expectedMaxIn);
  if (!units.ok) return { ...base, reason: `Unit check failed — ${units.note}` };

  try {
    const an = Datum.analyze(pos);
    const fx = Datum.extract(an);
    const mat = Datum.materialByName(mapped.datum);
    const tol = Datum.TOLS[opts.tolIndex ?? DEFAULT_TOL_INDEX];
    const rate = opts.shopRate ?? 85;

    // ---- STOCK -------------------------------------------------------
    //
    // Built by hand rather than through makeStock, because makeStock does
    // `Math.max(custom[i], ext[i])` -- it lets the mesh bounding box OVERRIDE
    // the BOM upward, and there is no way to opt out.
    //
    // That is not a theoretical concern. On BMS-851100038 the TCP plate came
    // back with 10.4 x 18.4 x 13.65 in of stock for a plate that is 1.375 in
    // thick: the mesh bbox spanned most of the mold base (stray bodies in the
    // merged STL), Math.max took it, and the "removed" volume went from ~25 in3
    // to 2,242 in3 -- 4,557 minutes of roughing on one clamp plate.
    //
    // The BOM is the authority on what steel gets bought. If the mesh disagrees
    // by more than the machining allowance, that is a DATA FAULT to report, not
    // a bigger block to quote.
    //
    // Shop rule for stock: finished size + 1/8 to 1/4 per side. Mid-range is
    // 3/16 per side = 0.375 in on each axis.
    const PER_SIDE_IN = 0.1875;
    const ALLOW_MM = PER_SIDE_IN * 2 * MM_PER_INCH;

    const extDesc = [...fx.ext].sort((a, b) => b - a);
    let stockMm: number[];
    let stockBasis: string;
    let short = false;

    if (haveStock) {
      // Pair by size rank: biggest BOM dimension onto the mesh's biggest axis.
      const bomDesc = [lIn, wIn, tIn].map((v) => v * MM_PER_INCH).sort((a, b) => b - a);
      const axesByExtent = [0, 1, 2].sort((a, b) => fx.ext[b] - fx.ext[a]);
      stockMm = [0, 0, 0];
      axesByExtent.forEach((axis, rank) => {
        stockMm[axis] = bomDesc[rank];
      });

      // Sanity gate. Compare like with like -- both sorted descending -- so an
      // axis-order difference is not mistaken for a size difference.
      const worstOversize = Math.max(
        ...extDesc.map((e, i) => e - bomDesc[i]),
      );
      short = worstOversize > 0.01;

      if (worstOversize > ALLOW_MM) {
        return {
          ...base,
          reason:
            `STL does not match the BOM row. Mesh is ` +
            `${extDesc.map((v) => (v / MM_PER_INCH).toFixed(2)).join(" x ")} in but the row says ` +
            `${bomDesc.map((v) => (v / MM_PER_INCH).toFixed(2)).join(" x ")} in ` +
            `(over by ${(worstOversize / MM_PER_INCH).toFixed(2)} in). Either the per-plate STL ` +
            `contains more than this plate, or the wrong quote row was matched — ` +
            `quoting either one would be guesswork.`,
        };
      }
      stockBasis = "stock from BOM row";
    } else {
      // No BOM row: derive stock the way the shop does, from the finished size
      // plus 3/16 per side.
      stockMm = fx.ext.map((v: number) => v + ALLOW_MM);
      stockBasis = `stock = finished + ${PER_SIDE_IN}" per side (no BOM row)`;
    }

    // Shape matches makeStock's return so plan() is happy.
    const st = { stock: stockMm, faces: [0, 1, 2], mode: "custom", short };

    const pl = Datum.plan(an, fx, mat, tol, rate, { st, tap: opts.assumeTaps !== false });

    // Full cost, not just machine time. cost() adds stock ($/kg), CAM
    // programming, inspection, cutter wear, deburr and a tolerance risk
    // multiplier. Computing hours x rate by hand drops roughly $640 a plate --
    // about $3,800 on a six-plate base -- which is exactly the kind of quiet
    // understatement someone would quote off.
    const cs = Datum.cost(pl, an, fx, mat, tol, rate);

    const stockVolIn3 = (pl.stockVol || 0) / Math.pow(MM_PER_INCH, 3);
    const finishedVolIn3 = (an.vol || 0) / Math.pow(MM_PER_INCH, 3);
    const removedIn3 = Math.max(0, stockVolIn3 - finishedVolIn3);

    // A finished part cannot be bigger than the stock it came out of, and it
    // cannot be a sliver of it either. Both mean the mesh volume is not this
    // plate: an unwelded or non-manifold STL gives a meaningless signed volume,
    // and a merged STL containing only a fragment gives a tiny one.
    //
    // On BMS-851100038 the ID POT read 29.7 in3 finished out of 1,177 in3 stock
    // -- 98% removed. A pot block is a solid chunk with a bore; it does not lose
    // 98% of itself. Quoting 16 hours off that is worse than saying "check this".
    const remainingPct = stockVolIn3 > 0 ? (100 * finishedVolIn3) / stockVolIn3 : 0;
    if (finishedVolIn3 > stockVolIn3 * 1.02) {
      return {
        ...base,
        reason:
          `Mesh volume (${finishedVolIn3.toFixed(1)} in3) exceeds its own stock ` +
          `(${stockVolIn3.toFixed(1)} in3). The STL is not watertight, so the computed ` +
          `volume is meaningless — re-export this plate.`,
      };
    }
    if (remainingPct < 12) {
      return {
        ...base,
        reason:
          `Mesh is only ${remainingPct.toFixed(1)}% of its stock ` +
          `(${finishedVolIn3.toFixed(1)} of ${stockVolIn3.toFixed(1)} in3). A mold plate does not ` +
          `lose ~90% of itself, so this STL is probably a fragment or not watertight. ` +
          `Re-export before trusting a time.`,
      };
    }

    // Datum's op list is already per-operation with tool, feed and pass counts.
    // Map it straight across; everything here is measured off the mesh.
    // NOTE: ops[].qty is a DISPLAY STRING in the engine ("12 holes",
    // "466 x 403 mm", "9 starts"), not a count. Treating it as a number made
    // `o.qty > 1` coerce to NaN > 1 -- always false -- so every quantity was
    // silently dropped from the detail line.
    const operations = (pl.ops || []).map((o: {
      label?: string; min?: number; tool?: string; detail?: string; qty?: string;
    }) => ({
      op: String(o.label || "Operation"),
      min: Math.round((Number(o.min) || 0) * 10) / 10,
      measured: true,
      detail: [o.tool, o.detail, o.qty ? String(o.qty) : ""].filter(Boolean).join(" · "),
    }));

    // Snap every op to the real crib and collect what it cannot do.
    const shopped = applyShopTools(operations);
    toolGaps.push(...shopped.gaps.map((g) => `${kind}: ${g}`));

    const cascadeNotes: string[] = [];

    // Expand the single rougher into the cascade the shop actually runs, BEFORE
    // calibration so each step still picks up the high-feed speed correction.
    // The two corrections are independent: the cascade fixes which cutter moves
    // the metal, the factor fixes how fast that cutter runs.
    const withCascade = expandRoughingCascade(
      shopped.ops, pl, fx, an, removedIn3, cascadeNotes,
    );

    // Recalibrate against the reference program. Datum uses one surface speed
    // for every operation; this shop runs high-feed indexable mills fast and
    // drills slowly, so the errors run in opposite directions and a plausible
    // total can hide wrong lines. See shopCalibration.ts for the observed data.
    const ops = withCascade.map((o) => {
      // Cascade steps are already priced from real cutting parameters. Applying
      // the blanket class factor on top would double-correct — see AdapterOp.
      if (o.preCalibrated) {
        return { ...o, detail: `${o.detail} · priced from ae·ap·vf, not class-calibrated` };
      }
      const { f, cls } = timeFactorFor(o.op);
      if (f === 1) return o;
      const adj = Math.round(o.min * f * 10) / 10;
      const dir = f > 1 ? "slower" : "faster";
      return {
        ...o,
        min: adj,
        detail:
          `${o.detail} · calibrated x${f} (${cls} runs ${dir} here than the ` +
          `generic model assumes)`,
      };
    });

    // Tool changes are real spindle time and were previously not charged at all,
    // only mentioned. Datum's count is DISTINCT TOOLS; the machine reloads a
    // cutter at every orientation that needs it, so the real count is roughly
    // double on a 5-orientation part. See realToolChanges().
    const distinctTools = Number(pl.toolChanges) || 0;
    const orientations = pl.setups ?? 1;
    const changes = realToolChanges(distinctTools, orientations);
    if (changes > 0) {
      ops.push({
        op: "Tool changes",
        min: Math.round(changes * TOOL_CHANGE_MIN * 10) / 10,
        measured: false,
        detail:
          `${changes} changes at ${TOOL_CHANGE_MIN} min · ${distinctTools} distinct tools ` +
          `recalled across ${orientations} B/C orientation(s). The reference programs ` +
          `run 48 and 46 M06 changes over 24 distinct tools — counting distinct tools ` +
          `understates carousel time by about half.`,
      });
    }

    // Setup: the UMC-1000 indexes B/C off ONE fixturing. Datum's "setups" are
    // orientations, and charging a full re-clamp for each was the biggest single
    // error in the first run (3 x 52 = 156 min on a part that clamps once).
    const setup = setupMinutes(orientations);
    ops.push({
      op: "Setup + fixturing",
      min: Math.round(setup.min * 10) / 10,
      measured: false,
      detail: setup.detail,
    });

    // Cycle from the CALIBRATED ops, not pl.cycle -- pl.cycle is the sum of the
    // uncalibrated minutes and would ignore every correction above.
    const cycleMin = ops
      .filter((o) => o.op !== "Setup + fixturing")
      .reduce((s, o) => s + o.min, 0);
    const perPartMin = cycleMin + setup.min;

    // Confidence band. A single number implies a precision this cannot have, and
    // an estimator who is handed one will quote it. See quoteConfidence.ts.
    const confidence = plateConfidence({
      ops,
      toolGapCount: toolGaps.length,
      stockFromBom: haveStock,
      stockShort: Boolean(st?.short),
      remainingPct,
      cornerLimited: Boolean(pl.cornerLimited),
      reachPct: Number(pl.reachPct),
    });

    return {
      component: model.name,
      role_label: kind,
      material: mapped.datum,
      qty,
      skipped: false,
      stock: { thickness: tIn, width: wIn, length: lIn },
      stock_volume_cuin: Math.round(stockVolIn3 * 100) / 100,
      finished_volume_cuin: Math.round(finishedVolIn3 * 100) / 100,
      finished_basis: `measured from mesh · ${stockBasis} · ${units.note}`,
      removed_cuin: Math.round(removedIn3 * 100) / 100,
      removed_pct: stockVolIn3 > 0 ? Math.round((1000 * removedIn3) / stockVolIn3) / 10 : 0,
      operations: ops,
      cut_min: Math.round(cycleMin * 10) / 10,
      // Deliberately NOT reporting Datum.DUTY here. It is declared in the engine
      // but never read by plan() -- pl.cycle is the raw sum of op minutes, not
      // derated -- so showing it as an efficiency factor would be a number the
      // model does not actually apply.
      per_part_min: Math.round(perPartMin * 10) / 10,
      total_min: Math.round(perPartMin * qty * 10) / 10,
      total_hours: Math.round(((perPartMin * qty) / 60) * 100) / 100,
      cost_usd: Math.round((Number(cs?.total) || 0) * qty * 100) / 100,
      stock_short: Boolean(st?.short),
      confidence: confidence,
      low_min: Math.round(perPartMin * confidence.lowFactor * qty * 10) / 10,
      high_min: Math.round(perPartMin * confidence.highFactor * qty * 10) / 10,
      cascade_note: cascadeNotes[0],
    };
  } catch (e) {
    return { ...base, reason: `Datum analysis failed: ${e instanceof Error ? e.message : "unknown"}` };
  }
}

/**
 * Feature list for one plate: every hole, pocket, bore, tap and chamfer the mesh
 * analysis found, with sizes.
 *
 * Separate from the time estimate on purpose. The estimate answers "how long";
 * this answers "what is actually on the part", which is what someone checking
 * the work needs. Nothing here is calibrated or derated -- these are measurements.
 */
async function geometryForPlate(
  kind: ComponentKind,
  model: AssetRef,
  row: QuoteLineItem | null,
): Promise<GeometryPart> {
  const tIn = Number(row?.thickness) || 0;
  const wIn = Number(row?.width) || 0;
  const lIn = Number(row?.length) || 0;
  const expectedMaxIn = Math.max(tIn, wIn, lIn);

  const base: GeometryPart = {
    role_label: kind,
    component: model.name,
    material: String(row?.material || "unknown"),
    holes: [],
    pockets: [],
    bores: [],
    taps: [],
    chamfer_count: 0,
    chamfer_len_in: 0,
    radii_count: 0,
    orientations: 0,
    distinct_tools: 0,
    skipped: true,
  };

  let pos: Float32Array;
  try {
    const res = await fetch(model.url);
    if (!res.ok) throw new Error(`${res.status} ${res.statusText}`);
    pos = Datum.parseSTL(await res.arrayBuffer());
  } catch (e) {
    return { ...base, reason: `Could not read STL: ${e instanceof Error ? e.message : "unknown"}` };
  }

  const units = toMillimetres(pos, expectedMaxIn);
  if (!units.ok) return { ...base, reason: `Unit check failed — ${units.note}` };

  try {
    const an = Datum.analyze(pos);
    const fx = Datum.extract(an);
    const mm2in = (mm: number) => mm / MM_PER_INCH;

    // Group identical holes so the list reads like a drill schedule rather than
    // 60 near-duplicate rows.
    const holeMap = new Map<string, GeoHole>();
    for (const h of fx.holes || []) {
      const dIn = mm2in(Number(h.d) || 0);
      const depIn = mm2in(Number(h.depth) || 0);
      const key = `${dIn.toFixed(4)}|${depIn.toFixed(3)}|${h.blind ? 1 : 0}|${h.purpose || ""}`;
      const existing = holeMap.get(key);
      if (existing) {
        existing.count += 1;
        continue;
      }
      const tap = impliedTapFromDrill(dIn);
      const drill = pickDrill(dIn, depIn);
      const anySize = drillExists(dIn);
      holeMap.set(key, {
        dia_in: Math.round(dIn * 10000) / 10000,
        depth_in: Math.round(depIn * 1000) / 1000,
        count: 1,
        blind: Boolean(h.blind),
        purpose: String(h.purpose || "fit"),
        thread: tap ? tap.spec : undefined,
        ld: dIn > 0 ? Math.round((depIn / dIn) * 10) / 10 : 0,
        // Reach is still worth calling out -- a 10xD hole needs a peck strategy
        // or a gun drill whatever machine it runs on. Not owning the drill is
        // not: the shop has it, on the machine this job would go to.
        tool: drill
          ? `T${drill.num} ${drill.name}`
          : anySize
            ? `T${anySize.num} right size, only ${anySize.depth.toFixed(2)}" reach`
            : toolLabel(bestDrill(dIn, depIn)),
      });
    }

    const pocketMap = new Map<string, GeoPocket>();
    for (const p of fx.pockets || []) {
      const depIn = mm2in(Number(p.depth) || 0);
      const areaIn = (Number(p.area) || 0) / (MM_PER_INCH * MM_PER_INCH);
      const key = `${depIn.toFixed(3)}|${areaIn.toFixed(2)}|${p.cbore ? 1 : 0}`;
      const ex = pocketMap.get(key);
      if (ex) {
        ex.count += 1;
        continue;
      }
      pocketMap.set(key, {
        depth_in: Math.round(depIn * 1000) / 1000,
        area_sqin: Math.round(areaIn * 100) / 100,
        count: 1,
        cbore: Boolean(p.cbore),
      });
    }

    // Bores and taps come off the plan, which is where Datum decides strategy.
    const mat = Datum.materialByName(mapMaterial(String(row?.material || "")).datum);
    const tol = Datum.TOLS[DEFAULT_TOL_INDEX];
    const st = Datum.makeStock(fx.ext, "block", Datum.ALLOW);
    const pl = Datum.plan(an, fx, mat, tol, 85, { st, tap: true });

    const bores: GeoBore[] = [];
    const taps: { spec: string; count: number; tool: string }[] = [];
    for (const o of pl.ops || []) {
      const lab = String(o.label || "");
      const bore = /Interpolate and bore Ø([\d.]+)\s*×?(\d+)?/i.exec(lab);
      if (bore) {
        bores.push({
          dia_in: Math.round(mm2in(parseFloat(bore[1])) * 10000) / 10000,
          count: bore[2] ? parseInt(bore[2], 10) : 1,
          method: "helical interpolation + boring head",
        });
        continue;
      }
      const tap = /Tap\s+([#\d/]+-\d+|M\d+)\s*×?(\d+)?/i.exec(lab);
      if (tap) {
        const spec = tap[1];
        const owned = TAPS.find((t) => t.spec.startsWith(spec));
        taps.push({
          spec,
          count: tap[2] ? parseInt(tap[2], 10) : 1,
          tool: owned
            ? `T${owned.num} ${owned.spec} (drill Ø${owned.tapDrill.toFixed(3)}")`
            : `NOT IN CRIB${/^M/i.test(spec) ? " — metric, likely " + (METRIC_TO_UNC[spec.toUpperCase()] || "a UNC size") : ""}`,
        });
      }
    }

    return {
      role_label: kind,
      component: model.name,
      material: String(row?.material || "unknown"),
      stock_in: tIn > 0 ? { thickness: tIn, width: wIn, length: lIn } : undefined,
      finished_volume_cuin:
        Math.round(((an.vol || 0) / Math.pow(MM_PER_INCH, 3)) * 100) / 100,
      holes: [...holeMap.values()].sort((a, b) => b.dia_in - a.dia_in),
      pockets: [...pocketMap.values()].sort((a, b) => b.depth_in - a.depth_in),
      bores: bores.sort((a, b) => b.dia_in - a.dia_in),
      taps,
      chamfer_count: (fx.chamfers || []).length,
      chamfer_len_in: Math.round(mm2in(Number(fx.chamferLen) || 0) * 10) / 10,
      radii_count: (fx.fillets || []).length,
      orientations: Number(fx.setups) || 1,
      distinct_tools: Number(pl.toolChanges) || 0,
      skipped: false,
    };
  } catch (e) {
    return { ...base, reason: `Analysis failed: ${e instanceof Error ? e.message : "unknown"}` };
  }
}

/** Feature report across the six plates. */
export async function geometryFromStls(
  models: AssetRef[],
  quoteItems: QuoteLineItem[],
  engravingByKind?: Map<ComponentKind, string[]>,
): Promise<GeometryReport> {
  const byKind = groupModelsByKind(models);

  const parts: GeometryPart[] = [];
  for (const kind of PLATE_KINDS) {
    const found = byKind.get(kind);
    if (!found) continue;
    const row = rowForKind(quoteItems, kind);
    for (const model of modelsToAnalyse(found, row)) {
      const p = await geometryForPlate(kind, model, row);
      p.engraving = engravingByKind?.get(kind);
      parts.push(p);
    }
  }

  const live = parts.filter((p) => !p.skipped);
  const totals = {
    holes: live.reduce((s, p) => s + p.holes.reduce((a, h) => a + h.count, 0), 0),
    tapped: live.reduce((s, p) => s + p.taps.reduce((a, t) => a + t.count, 0), 0),
    pockets: live.reduce((s, p) => s + p.pockets.reduce((a, k) => a + k.count, 0), 0),
    bores: live.reduce((s, p) => s + p.bores.reduce((a, b) => a + b.count, 0), 0),
    chamfer_len_in: Math.round(live.reduce((s, p) => s + p.chamfer_len_in, 0) * 10) / 10,
  };

  return {
    parts,
    totals,
    notes: [
      "Every size here is MEASURED off the mesh — diameters fitted from the facets, " +
        "depths measured, through-vs-blind resolved by ray cast. Nothing is calibrated " +
        "or derated, so these are the numbers to check the model against.",
      "Holes are grouped by diameter, depth and purpose, the way a drill schedule reads.",
      "L/D above about 8 needs a deep-hole strategy: through coolant and reduced feed. " +
        "Above 5 the crib may not have the reach — those rows say so.",
      "ENGRAVING: the letters are built by the macro from job data (customer number, " +
        "job number, weight, date) and cut as a separate DXF. They are far too fine to " +
        "read back out of an STL, so the text shown is what the macro stamped, not " +
        "something recovered from the mesh.",
    ],
  };
}

/**
 * Machining estimate for a whole mold base, measured from the six plate STLs.
 *
 * Shaped exactly like the Python endpoint's MachiningEstimate, so the panel can
 * render either one.
 */
export async function estimateFromStls(
  models: AssetRef[],
  quoteItems: QuoteLineItem[],
  opts: DatumOptions = {},
): Promise<MachiningEstimate> {
  const byKind = groupModelsByKind(models);

  const parts: MachiningPart[] = [];
  const substitutions = new Set<string>();
  const toolGaps: string[] = [];

  for (const kind of PLATE_KINDS) {
    const found = byKind.get(kind);
    if (!found) continue;
    const row = rowForKind(quoteItems, kind);
    const mapped = mapMaterial(String(row?.material || ""));
    if (mapped.note) substitutions.add(mapped.note);
    for (const model of modelsToAnalyse(found, row)) {
      parts.push(await estimateOnePlate(kind, model, row, opts, toolGaps));
    }
  }

  const priced = parts.filter((p) => !p.skipped);
  const totalMin = priced.reduce((s, p) => s + p.total_min, 0);

  // Weight by qty, or these two will visibly fail to add up to Total time on any
  // plate quoted more than once (rails, ejector pairs).
  const measuredMin = priced.reduce(
    (s, p) =>
      s + p.qty * (p.operations || []).filter((o) => o.measured).reduce((a, o) => a + o.min, 0),
    0,
  );
  const assumedMin = priced.reduce(
    (s, p) =>
      s + p.qty * (p.operations || []).filter((o) => !o.measured).reduce((a, o) => a + o.min, 0),
    0,
  );

  const rate = opts.shopRate ?? 85;
  const hours = totalMin / 60;
  const tol = Datum.TOLS[opts.tolIndex ?? DEFAULT_TOL_INDEX];

  // Full cost from Datum.cost(): machine + setup + stock + programming +
  // inspection + cutter wear + deburr, with the tolerance risk multiplier.
  const fullCost = priced.reduce((s, p) => s + (p.cost_usd || 0), 0);
  const shortPlates = priced.filter((p) => p.stock_short).map((p) => p.role_label);

  // Job-level band. Damped rather than summed -- a six-plate base absorbs one
  // plate's surprise better than a single part does. See rollUpConfidence().
  // The label matters: a driver that does not say WHICH plate is not an action.
  // See aggregateDrivers() in quoteConfidence.ts.
  const jobBand: ConfidenceBand | null = rollUpConfidence(
    priced
      .filter((p) => p.confidence)
      .map((p) => ({
        minutes: p.total_min,
        band: p.confidence as ConfidenceBand,
        label: p.role_label || p.component || "",
      })),
  );
  const lowMin = jobBand ? totalMin * jobBand.lowFactor : totalMin;
  const highMin = jobBand ? totalMin * jobBand.highFactor : totalMin;

  const cascadeNotes = Array.from(
    new Set(priced.map((p) => p.cascade_note).filter((n): n is string => Boolean(n))),
  );

  const notes = [
    "Holes, pockets, counterbores and chamfers are FOUND IN THE MESH, not " +
      "assumed from a pattern table — diameter and depth are fitted, and " +
      "through-vs-blind is a ray cast.",
    `Every operation is quoted with the BEST TOOL FOR THE JOB. Where that tool is on ` +
      `this machine's list it is called out by carousel number; where it is not, it is ` +
      `named and marked "another machine", because the shop's tooling is spread across ` +
      `machines and the job runs on the one that has it. Facing is always the 2" face ` +
      `mill. Nothing is ever downgraded to a smaller cutter that would quote a slower cut.`,
    `Stock is the BOM row, capped so a bad mesh cannot inflate it. Where there is no ` +
      `row, stock = finished + 3/16" per side (mid of your 1/8–1/4 rule).`,
    `Tolerance band ${tol?.label ?? "±0.025"} (${tol?.name ?? "Precision"}) — set from the ` +
      `shop's ±.001 in working tolerance. Applied to the whole plate, because an STL ` +
      `carries no per-feature tolerance.`,
    "Cost is the full Datum breakdown — machine time, setup, stock at $/kg, CAM " +
      "programming, inspection, cutter wear and deburr, with a tolerance risk " +
      "multiplier. It is NOT just hours x rate.",
    "Setup and CAM programming land in full on every plate, because each plate " +
      "is its own program. On a six-plate base that is six setups and six " +
      "programming charges — check that against how you actually run them.",
    `Times are calibrated against ${REFERENCE_JOB.name}: ${REFERENCE_JOB.distinctTools} ` +
      `distinct tools, ${REFERENCE_JOB.toolChanges} tool changes, ` +
      `${REFERENCE_JOB.orientations} B/C orientations off ONE setup. High-feed indexable ` +
      `mills run faster here than the generic model assumes; drilling and spotting run ` +
      `much slower. Both corrections are applied per operation and shown on each line.`,
    "Setup is one fixturing plus B/C indexes, NOT one setup per direction — the " +
      "UMC-1000 rotates to the part rather than being re-clamped. That single " +
      "correction removed ~100 min per plate.",
    "Cycle time still excludes probing and chip clearing, so add your own factor.",
    `Factors are MEASURED, not inferred: the whole toolpath of ${REFERENCE_JOB.name} was ` +
      `simulated (${REFERENCE_JOB.measuredFeedMin} min feed, ` +
      `${REFERENCE_JOB.measuredTotalMin} min with changes and rapids) and each operation ` +
      `class compared against what this model predicted for it.`,
    `Biggest correction — HELICAL BORING, now x3.8. ${REFERENCE_JOB.longestOp} The model ` +
      `assumed roughly 0.045 in per lap, so it was 4x under on the single longest ` +
      `operation in the program.`,
    `Finishing is halved (x0.43): 62.8 min measured against 124.6 predicted. The model ` +
      `rasters an area with one small cutter; the program contours per Z step with ` +
      `T19 / T13 / T174. The factor patches the symptom, the tool cascade is the real fix.`,
    `ROUGHING IS NOW A CASCADE, not one cutter. The engine sizes a single rougher to ` +
      `the tightest internal corner and charges all the bulk removal to it — on the ` +
      `reference pot that meant hogging with a Ø0.123" endmill the crib does not own. ` +
      `Each roughing op is expanded into the ladder the shop really runs (2" → 1" → ` +
      `0.75" → 0.5"), with volume split as D² and time equal per step. That equal-time ` +
      `prediction is measured, not assumed: 18.2 / 18.5 / 19.3 min on J8441 and ` +
      `17.9 / 18.0 / 19.9 on J8410.`,
    `REMAINING WEAKNESS: finishing. The x0.43 factor above is still a patch. The engine ` +
      `rasters an area with one small cutter where the program contours per Z step, and ` +
      `the cascade fix does not touch that — finishing tool choice is the least ` +
      `trustworthy part of this estimate.`,
  ];
  if (jobBand) {
    notes.push(
      `CONFIDENCE ${jobBand.grade} (${jobBand.score}/100) — ${jobBand.guidance} ` +
        `Band is ${Math.round(lowMin)}–${Math.round(highMin)} min against a point estimate ` +
        `of ${Math.round(totalMin)}. The band opens further on the high side than the low ` +
        `because unplanned work gets discovered at the machine, not removed.`,
      ...jobBand.drivers.map((d) => `  · ${d}`),
    );
  }
  notes.push(...cascadeNotes.map((n) => `Roughing plan: ${n}`));
  if (substitutions.size > 0) {
    notes.push(`Material substitutions: ${[...substitutions].join("; ")}.`);
  }
  if (shortPlates.length > 0) {
    notes.push(
      `BOM stock is smaller than the measured part on: ${shortPlates.join(", ")}. ` +
        `Within the 3/16" per side allowance, so it is quoted — but the row and the ` +
        `CAD disagree, so check the steel sheet before cutting.`,
    );
  }
  if (toolGaps.length > 0) {
    // Dedupe: the same note usually repeats across plates.
    //
    // These are NOT "gaps" any more. Not owning a tool on this machine stopped
    // being a finding the moment the estimator started quoting the right tool
    // wherever it lives. What survives here is worth reading: reach limits that
    // need a peck strategy or a gun drill, and tap-drill diameters that differ
    // from nominal — both true on any machine.
    const uniq = [...new Set(toolGaps)];
    notes.push(
      `TOOLING NOTES (${uniq.length}) — reach limits and size calls worth confirming. ` +
        `Tools that simply live on another machine are quoted normally and are not listed here:`,
      ...uniq.slice(0, 12),
    );
    if (uniq.length > 12) notes.push(`…and ${uniq.length - 12} more.`);
  }
  const rejected = parts.filter((p) => p.skipped);
  if (rejected.length > 0) {
    notes.push(
      `${rejected.length} plate(s) were NOT quoted because the STL and the quote row ` +
        `could not be reconciled — see the "Not estimated" list. A refused plate is ` +
        `deliberate: a wrong number here becomes a wrong price.`,
    );
  }

  return {
    parts,
    summary: {
      part_count: priced.length,
      skipped_count: parts.length - priced.length,
      total_minutes: Math.round(totalMin * 10) / 10,
      total_hours: Math.round(hours * 100) / 100,
      shop_rate_per_hour: rate,
      estimated_cost: Math.round(fullCost * 100) / 100,
      measured_minutes: Math.round(measuredMin * 10) / 10,
      assumed_minutes: Math.round(assumedMin * 10) / 10,
      confidence_pct:
        measuredMin + assumedMin > 0
          ? Math.round((100 * measuredMin) / (measuredMin + assumedMin))
          : 0,
      confidence: jobBand ?? undefined,
      low_minutes: Math.round(lowMin * 10) / 10,
      high_minutes: Math.round(highMin * 10) / 10,
      low_cost: Math.round(fullCost * (jobBand?.lowFactor ?? 1) * 100) / 100,
      high_cost: Math.round(fullCost * (jobBand?.highFactor ?? 1) * 100) / 100,
    },
    rates_used: {
      engine: "Datum (mesh feature recognition)",
      duty_factor: Datum.DUTY,
      tolerance: tol?.label ?? "±0.05",
      shop_rate_per_hour: rate,
      plates_found: priced.length,
    },
    notes,
  };
}
