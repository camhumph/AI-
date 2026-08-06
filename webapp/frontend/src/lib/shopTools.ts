/**
 * The actual tool crib, from the shop's tool list.
 *
 * WHY THIS EXISTS
 *   Datum sizes tools from continuous maths -- rougher diameter is
 *   `cbrt(stockVol) * 0.12` clamped to 8-20 mm, finisher is 65% of that, and a
 *   drill is exactly the hole diameter. Real cribs are discrete. That mismatch
 *   is why the first run quoted a "Ø7 4-flute carbide" and a "Ø14.29 carbide"
 *   drill: neither exists here, and no machinist could audit the number.
 *
 *   Snapping to real tools also catches things the continuous model cannot see:
 *   a hole with no drill to make it, a pocket deeper than any cutter reaches,
 *   and -- the big one on this list -- taps that do not exist.
 *
 * THE TAP FINDING
 *   This crib is UNC only: 1/4-20, 5/16-18, 3/8-16, 1/2-13. There is not one
 *   metric tap in it. Datum's first pass reported "Tap M12 x5", "Tap M10 x5",
 *   "Tap M6 x23" because it matches hole diameters against BOTH metric-coarse
 *   and unified tap-drill tables. Every one of those metric callouts is wrong
 *   for this shop -- and note the drill sizes here ARE the UNC tap drills
 *   (0.213 -> 1/4-20, 0.265 -> 5/16-18, 0.3125 -> 3/8-16, 0.422 -> 1/2-13),
 *   which is strong evidence the imperial reading is the correct one.
 *
 * All sizes in INCHES, as written on the tool list. Convert at the boundary.
 */

export const MM_PER_IN = 25.4;

export interface ShopTool {
  /** Carousel number, so a machinist can find it. 0 when off this machine. */
  num: number;
  name: string;
  /** Overall diameter, inches. */
  dia: number;
  /** Usable cutting depth, inches. Governs reach. */
  depth: number;
  /** Overall length, inches. */
  oal: number;
  flutes?: number;
  /**
   * True when this tool is NOT on the list below.
   *
   * THE LIST IS ONE MACHINE, NOT THE SHOP. The shop owns the tooling it needs;
   * it is spread across machines, and a job runs on whichever machine has the
   * right tool. So a tool that is missing HERE is not a tool the shop lacks --
   * it is a tool on another machine, and the job is quoted with it.
   *
   * Set on tools this file synthesised at the ideal size. They still price
   * normally; the flag exists so the callout can say where the tool lives
   * instead of pretending it has a carousel number on this machine.
   */
  offMachine?: boolean;
}

/**
 * The face mill every finish facing pass runs.
 *
 * Fixed at 2", per the shop: whatever squares the stock, the FINISH is a 2"
 * face mill. Datum used to assume 63 mm (2.48") and the crib lookup then
 * labelled the op with the 3" finisher, so the quote was priced on a cutter
 * 24% wider than the one that actually runs -- which under-counts width passes
 * on every plate that gets faced.
 */
export const FINISH_FACE_MILL_IN = 2.0;

/** The named 2" face mill, or a synthesised one if it ever leaves the list. */
export function finishFaceMill(): ShopTool {
  const exact = FACE_MILLS.filter(
    (t) => Math.abs(t.dia - FINISH_FACE_MILL_IN) <= 0.01,
  ).sort((a, b) => b.depth - a.depth);
  return (
    exact[0] ?? {
      num: 0,
      name: `${FINISH_FACE_MILL_IN}" face mill`,
      dia: FINISH_FACE_MILL_IN,
      depth: 0.375,
      oal: 6,
      offMachine: true,
    }
  );
}

/** Face and shell mills, for squaring stock. */
export const FACE_MILLS: ShopTool[] = [
  { num: 47, name: '3" facemill finish', dia: 3.0, depth: 0.25, oal: 8 },
  { num: 2, name: '2.25" face/shell mill', dia: 2.25, depth: 0.25, oal: 4 },
  { num: 24, name: '2" Zenit HF mill', dia: 2.0, depth: 0.375, oal: 6 },
  { num: 46, name: '1" Dijet', dia: 1.0, depth: 0.375, oal: 6 },
  { num: 23, name: '1" Zenit HF mill', dia: 1.0, depth: 0.225, oal: 4.5 },
  { num: 22, name: '0.75" Zenit HF mill', dia: 0.75, depth: 0.25, oal: 4 },
  { num: 21, name: '0.625" Zenit HF mill', dia: 0.625, depth: 0.25, oal: 2 },
];

/** Square-end mills, roughing and finishing. */
export const END_MILLS: ShopTool[] = [
  { num: 42, name: '2" bull-nose', dia: 1.9978, depth: 0.5, oal: 8 },
  { num: 21, name: '1" flat endmill', dia: 1.0, depth: 3, oal: 4 },
  { num: 173, name: '1" long finish EM', dia: 1.0, depth: 4, oal: 4.5 },
  { num: 190, name: '0.75" long finish EM', dia: 0.7487, depth: 4, oal: 6 },
  { num: 16, name: '0.75" flat Widin', dia: 0.75, depth: 1.25, oal: 4 },
  { num: 17, name: '0.75" flat finish', dia: 0.75, depth: 2, oal: 4 },
  { num: 20, name: '0.75" flat endmill', dia: 0.75, depth: 3, oal: 4 },
  { num: 175, name: '0.5" flat endmill', dia: 0.5, depth: 3, oal: 4 },
  { num: 18, name: '0.5" flat endmill', dia: 0.5, depth: 2, oal: 3 },
  { num: 14, name: '0.5" flat Widin', dia: 0.5, depth: 1, oal: 3 },
  { num: 15, name: '0.5" flat finish', dia: 0.5, depth: 1, oal: 4 },
  { num: 174, name: '0.499" flat endmill', dia: 0.499, depth: 1, oal: 4 },
  { num: 20, name: '0.5" Raptor', dia: 0.5, depth: 1, oal: 3 },
  { num: 162, name: '0.375" long endmill', dia: 0.375, depth: 1, oal: 4 },
  { num: 12, name: '0.375" flat endmill', dia: 0.375, depth: 1, oal: 2 },
  { num: 13, name: '0.375" flat finish', dia: 0.375, depth: 1, oal: 4 },
  { num: 184, name: '0.3125" Widin .02R', dia: 0.3125, depth: 0.5, oal: 1.25 },
  { num: 198, name: '0.3125" 5FL Gorilla', dia: 0.3125, depth: 1, oal: 4 },
  { num: 182, name: '0.25" Gorilla long 5FL', dia: 0.25, depth: 1.5, oal: 2 },
  { num: 11, name: '0.25" Gorilla hardmill', dia: 0.25, depth: 1, oal: 4 },
  { num: 183, name: '0.25" Widin finish', dia: 0.25, depth: 1, oal: 1.625 },
  { num: 10, name: '0.25" flat endmill', dia: 0.25, depth: 0.75, oal: 0.875 },
  { num: 181, name: '0.236" Widin', dia: 0.236, depth: 0.75, oal: 1.3 },
  { num: 19, name: '0.1875" flat endmill', dia: 0.1875, depth: 1, oal: 4 },
  { num: 176, name: '0.123" flat endmill', dia: 0.123, depth: 1, oal: 4 },
];

/** Ball-nose, for radii and blends. */
export const BALL_MILLS: ShopTool[] = [
  { num: 180, name: '0.5" ball mill', dia: 0.5, depth: 1, oal: 4 },
  { num: 25, name: '0.5" Dapra', dia: 0.5, depth: 0.5, oal: 6 },
  { num: 11, name: '0.25" ball nose', dia: 0.25, depth: 1, oal: 4 },
  { num: 176, name: '0.125" ball nose', dia: 0.123, depth: 1, oal: 4 },
  { num: 176, name: '0.06" ball nose', dia: 0.06, depth: 1, oal: 4 },
];

/** Twist, spade and pot drills. depth is the usable flute length. */
export const DRILLS: ShopTool[] = [
  { num: 159, name: '0.9375" drill', dia: 0.9375, depth: 1, oal: 6 },
  { num: 41, name: '0.6875" long spade', dia: 0.6875, depth: 0.625, oal: 7 },
  { num: 35, name: '0.6875" drill', dia: 0.6875, depth: 5, oal: 6 },
  { num: 40, name: '0.6875" Ingersoll pilot', dia: 0.656, depth: 1.75, oal: 2 },
  { num: 200, name: '0.640" drill', dia: 0.64, depth: 6, oal: 6 },
  { num: 38, name: '0.640" Ingersoll LONGPOTS', dia: 0.64, depth: 7, oal: 7.5 },
  { num: 39, name: '0.640" Ingersoll long', dia: 0.64, depth: 7.5, oal: 8 },
  { num: 37, name: '0.640" Ingersoll pilot', dia: 0.64, depth: 2, oal: 3 },
  { num: 159, name: '0.625" drill', dia: 0.625, depth: 1, oal: 6 },
  { num: 35, name: '0.625" spade drill', dia: 0.625, depth: 6, oal: 6 },
  { num: 120, name: '0.5625" drill', dia: 0.5625, depth: 0.75, oal: 5.2 },
  { num: 35, name: '0.5315" spade drill', dia: 0.5315, depth: 5, oal: 6 },
  { num: 152, name: '0.5" drill', dia: 0.5, depth: 5.125, oal: 5.5 },
  { num: 152, name: '0.5" OSG pot drill', dia: 0.5, depth: 4, oal: 4 },
  { num: 151, name: '0.406" drill', dia: 0.406, depth: 5.2, oal: 6 },
  { num: 1, name: '0.422" drill', dia: 0.422, depth: 3, oal: 4.5 },
  { num: 4, name: '0.375" drill', dia: 0.375, depth: 1, oal: 2 },
  { num: 158, name: '0.375" long drill', dia: 0.375, depth: 6, oal: 7 },
  { num: 160, name: '0.343" drill', dia: 0.343, depth: 10, oal: 12 },
  { num: 150, name: '0.343" drill', dia: 0.343, depth: 1, oal: 2 },
  { num: 161, name: '0.3125" drill', dia: 0.3125, depth: 2.5, oal: 4 },
  { num: 156, name: '0.281" drill', dia: 0.281, depth: 1, oal: 2 },
  { num: 164, name: '0.265" drill', dia: 0.265, depth: 1, oal: 2 },
  { num: 163, name: '0.250" drill', dia: 0.25, depth: 1, oal: 2 },
  { num: 3, name: '0.213" drill', dia: 0.213, depth: 1, oal: 2 },
  { num: 153, name: '0.1875" drill', dia: 0.1875, depth: 3, oal: 3.5 },
  { num: 101, name: '0.14" drill', dia: 0.14, depth: 3.5, oal: 4 },
];

export const SPOT_DRILLS: ShopTool[] = [
  { num: 7, name: '0.625" spot drill', dia: 0.625, depth: 1, oal: 3.3 },
  { num: 8, name: '0.375" x 90 spot drill', dia: 0.375, depth: 0.75, oal: 3.5 },
  { num: 8, name: '0.25" spot drill', dia: 0.25, depth: 2, oal: 4 },
  { num: 9, name: '0.25" spot drill', dia: 0.25, depth: 1, oal: 2 },
];

export const CHAMFER_MILLS: ShopTool[] = [
  { num: 6, name: '1" chamfer mill', dia: 1.0, depth: 1, oal: 3 },
  { num: 199, name: '0.5" carbide chamfer', dia: 0.5, depth: 0.5, oal: 3.5 },
  { num: 5, name: '0.45" chamfer mill', dia: 0.45, depth: 1, oal: 2 },
  { num: 5, name: '0.375" Gorilla chamfer', dia: 0.375, depth: 1, oal: 3 },
];

export interface ShopTap {
  num: number;
  /** As called out on the print. */
  spec: string;
  /** Nominal thread diameter, inches. */
  dia: number;
  /** Threads per inch. */
  tpi: number;
  /** Correct tap drill, inches. */
  tapDrill: number;
  /** Usable thread depth, inches. */
  depth: number;
}

/**
 * Every tap in the crib. UNC only -- there is no metric tap on this list, so any
 * metric callout the mesh analysis produces is a misread.
 */
export const TAPS: ShopTap[] = [
  { num: 191, spec: "1/4-20 UNC", dia: 0.25, tpi: 20, tapDrill: 0.201, depth: 1 },
  { num: 32, spec: "5/16-18 UNC", dia: 0.3125, tpi: 18, tapDrill: 0.257, depth: 1.125 },
  { num: 171, spec: "3/8-16 UNC", dia: 0.375, tpi: 16, tapDrill: 0.312, depth: 1.25 },
  { num: 30, spec: "1/2-13 UNC", dia: 0.5, tpi: 13, tapDrill: 0.4219, depth: 2 },
];

/** Thread mills, for sizes with no tap or for blind holes near the bottom. */
export const THREAD_MILLS: ShopTool[] = [
  { num: 45, name: "1/8 NPT thread mill", dia: 0.3125, depth: 0.4, oal: 2.5 },
  { num: 192, name: "7/16-20 thread mill", dia: 0.3125, depth: 1, oal: 2.5 },
  { num: 179, name: '0.36" thread mill', dia: 0.36, depth: 0.59, oal: 2.5 },
];

/** Odds and ends worth naming so they are not silently ignored. */
export const SPECIALS: ShopTool[] = [
  { num: 43, name: '0.705" dovetail cutter', dia: 0.712, depth: 0.27, oal: 0.75 },
  { num: 195, name: '0.5" snap ring cutter', dia: 0.5, depth: 0.056, oal: 2.5 },
  { num: 188, name: '0.062" Widin grease cutter', dia: 0.062, depth: 0.2, oal: 2 },
];

/* ---------------------------------------------------------------- lookups */

/**
 * Largest tool from `list` that is no bigger than `maxDia` and reaches `depth`.
 * Returns null when nothing in the crib can do it -- which is a real answer, not
 * a failure: it means the job needs a tool the shop does not own.
 */
export function pickTool(
  list: ShopTool[],
  maxDiaIn: number,
  needDepthIn = 0,
): ShopTool | null {
  const usable = list
    .filter((t) => t.dia <= maxDiaIn + 1e-6 && t.depth >= needDepthIn - 1e-6)
    .sort((a, b) => b.dia - a.dia);
  return usable[0] ?? null;
}

/**
 * The drill that actually makes this hole.
 *
 * A drill must be the hole size -- not "at least" the hole size. The first
 * version filtered `dia >= hole` and took the smallest, which let a 0.625"
 * spade drill be offered for a 0.500" hole (Ø12.70 mm) simply because no 0.500"
 * drill in the crib reached 5.5" deep. That is not a substitution, it is a
 * different hole.
 *
 * So: match within DRILL_MATCH_IN, and if nothing that size reaches the depth,
 * return null so the caller can flag it. "No tool reaches" is the correct
 * answer and a useful one -- it means the hole needs a step drill, an extension,
 * or gun drilling.
 */
export const DRILL_MATCH_IN = 0.006;

export function pickDrill(holeDiaIn: number, depthIn = 0): ShopTool | null {
  const sameSize = DRILLS.filter((t) => Math.abs(t.dia - holeDiaIn) <= DRILL_MATCH_IN);
  if (sameSize.length === 0) return null;
  // Deepest reach first, so a long-flute variant wins over a stubby one.
  const reaches = sameSize
    .filter((t) => t.depth >= depthIn - 1e-6)
    .sort((a, b) => b.depth - a.depth);
  return reaches[0] ?? null;
}

/** Is there a drill of this size at all, ignoring reach? */
export function drillExists(holeDiaIn: number): ShopTool | null {
  const hits = DRILLS.filter((t) => Math.abs(t.dia - holeDiaIn) <= DRILL_MATCH_IN).sort(
    (a, b) => b.depth - a.depth,
  );
  return hits[0] ?? null;
}

/**
 * The UNC tap whose tap drill matches this hole, so a metric misread can be
 * corrected rather than just rejected.
 *
 * This matters on real parts: Ø5.11 mm is 0.2012", which is EXACTLY the 1/4-20
 * tap drill. Datum reads it as M6 (5.0 mm tap drill, 0.197") because the metric
 * table is close enough. On a UNC-only crib the 1/4-20 reading is right, and
 * saying so is far more useful than "not in crib".
 */
export function impliedTapFromDrill(holeDiaIn: number): ShopTap | null {
  const hits = TAPS.map((t) => ({ t, d: Math.abs(t.tapDrill - holeDiaIn) }))
    .filter((x) => x.d <= 0.008)
    .sort((a, b) => a.d - b.d);
  return hits[0]?.t ?? null;
}

/** Metric tap nominal -> the UNC tap this shop would actually use. */
export const METRIC_TO_UNC: Record<string, string> = {
  M5: "#10-24 (nearest; not in crib)",
  M6: "1/4-20",
  M8: "5/16-18",
  M10: "3/8-16",
  M12: "1/2-13",
  M14: "9/16-12 (not in crib)",
  M16: "5/8-11 (not in crib)",
};

/** Exact-ish drill match, for deciding whether a hole IS a tap drill. */
export function drillNear(diaIn: number, tolIn = 0.004): ShopTool | null {
  return DRILLS.find((t) => Math.abs(t.dia - diaIn) <= tolIn) ?? null;
}

/**
 * The tap this hole is drilled for, or null.
 *
 * UNC only, deliberately -- see the header. Tolerance is tight (0.006") because
 * the tap drills in this crib are distinct enough not to need slop, and loose
 * matching is exactly how a Ø0.25 coolant line becomes a tapped hole.
 */
export function tapForHole(holeDiaIn: number, depthIn = 0): ShopTap | null {
  const hits = TAPS.filter(
    (t) => Math.abs(t.tapDrill - holeDiaIn) <= 0.006 && t.depth >= depthIn - 1e-6,
  );
  return hits[0] ?? null;
}

/** Working tolerance, inches. The shop holds ±.001 on most things. */
export const SHOP_TOL_IN = 0.001;

/**
 * Datum's tolerance bands are metric. ±.001" is ±0.0254 mm, which is the
 * "Precision" band (±0.025), NOT the "Close" band (±0.05) the adapter defaulted
 * to. That default was one band too loose for this shop: it under-counts finish
 * passes, inspection minutes and scrap risk.
 */
export const DATUM_TOL_INDEX_FOR_SHOP = 3;

/* ------------------------------------------- best tool, wherever it lives */
/**
 * THE CRIB IS A PREFERENCE, NOT A GATE.
 *
 * pickTool and pickDrill above answer "is this tool on THIS machine", and
 * returning null was being read as "this shop cannot do this op". That is the
 * wrong conclusion. The shop has the tooling; it is distributed across
 * machines, and a job goes to the machine that has what it needs. Marking the
 * op NO TOOL, dropping it from the time, and docking the quote's confidence
 * was three penalties for a fact about one carousel.
 *
 * These two always return a tool:
 *   - a real one from the list when it fits, so the callout keeps its carousel
 *     number, its true reach and its real diameter, and
 *   - otherwise the IDEAL tool for the job at the size the op actually wants,
 *     flagged offMachine so the callout can say so.
 *
 * Pricing on the ideal size is the point. Snapping a Ø0.640" hole down to the
 * biggest drill that happens to be here does not make the hole; quoting the
 * 0.640" drill does, and that is the tool that will run.
 *
 * WHAT THIS DOES NOT RELAX: geometry. A cutter still cannot be wider than the
 * corner it has to reach into, and roughingCascade still steps down for rest
 * material. Those are facts about the part, and no other machine changes them.
 */
export function bestTool(
  list: ShopTool[],
  idealDiaIn: number,
  needDepthIn = 0,
  kind = "cutter",
): ShopTool {
  const onMachine = pickTool(list, idealDiaIn, needDepthIn);
  // A crib tool counts only when it is genuinely the right size. Anything much
  // smaller is a substitution that would quote a slower cut than the one that
  // will actually run.
  if (onMachine && (!Number.isFinite(idealDiaIn) || onMachine.dia >= idealDiaIn * 0.9)) {
    return onMachine;
  }
  const dia = Number.isFinite(idealDiaIn) ? idealDiaIn : (onMachine?.dia ?? 1);
  return {
    num: 0,
    name: `${dia.toFixed(3)}" ${kind}`,
    dia,
    depth: Math.max(needDepthIn, onMachine?.depth ?? 0, dia * 3),
    oal: Math.max(needDepthIn * 1.5, dia * 4),
    offMachine: true,
  };
}

/** The drill that makes this hole — on this machine, or on the one that has it. */
export function bestDrill(holeDiaIn: number, depthIn = 0): ShopTool {
  const onMachine = pickDrill(holeDiaIn, depthIn);
  if (onMachine) return onMachine;
  return {
    num: 0,
    name: `${holeDiaIn.toFixed(3)}" drill`,
    dia: holeDiaIn,
    depth: Math.max(depthIn, holeDiaIn * 5),
    oal: Math.max(depthIn * 1.5, holeDiaIn * 6),
    offMachine: true,
  };
}

/** How to name a tool in a callout: carousel number here, or where it lives. */
export function toolLabel(t: ShopTool): string {
  return t.offMachine ? `${t.name} (another machine)` : `T${t.num} ${t.name}`;
}

export interface ToolGap {
  op: string;
  need: string;
  note: string;
}

/** Biggest endmill in the crib, for capping Datum's rougher. */
export function largestEndMillIn(): number {
  return END_MILLS.reduce((m, t) => Math.max(m, t.dia), 0);
}

/** Biggest face mill, for capping facing width. */
export function largestFaceMillIn(): number {
  return FACE_MILLS.reduce((m, t) => Math.max(m, t.dia), 0);
}
