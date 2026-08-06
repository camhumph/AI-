// Ported 1:1 from the shop's Match Studio classifier
// (ms_classify_component_filename in the Elgin CMS backend) so a part
// groups the same way here as it does in every other shop tool.

/**
 * PCS / DME mold-base plate vocabulary.
 *
 * Covers the series a shop actually sees:
 *
 *   A / B series  two-plate. A-clamping, A plate, B plate, support, rails,
 *                 ejector stack, bottom clamp.
 *   T series      THREE-plate. Adds X-1 (runner stripper) and X-2 (cavity)
 *                 floating plates, giving two parting lines -- the first opens
 *                 between X-1 and X-2 to break the part off the gate.
 *   X series      stripper-plate, in 5-plate (5X) and 6-plate (6X) forms. The
 *                 stripper (X) plate sits between AX and BX and pushes the part
 *                 off the core; 6X adds a support plate.
 *   hot runner    manifold plate + backing plate above the A plate.
 *
 * Module6121's StandardPlateNameStd can already emit "X" Plate, "Y" Plate,
 * Manifold Plate, Die Plate, Die Backup Plate and Runner Stripper Plate. None of
 * them existed here, so every one silently became OTHER -- invisible in the 3D
 * gallery, absent from the Geometry tab, and skipped by the machining estimate.
 * Any T-series or X-series job was quietly half-processed.
 */
export type ComponentKind =
  | "FULL ASSEMBLY"
  // --- two-plate (A / B series) ---
  | "TOP CLAMP PLATE"
  | "A PLATE"
  | "B PLATE"
  | "SUPPORT PLATE"
  | "RAILS"
  | "EJECTOR PLATE"
  | "BOTTOM EJECTOR PLATE"
  | "PIN PLATE"
  | "BOTTOM CLAMP PLATE"
  // --- stripper-plate (X series, 5X / 6X) ---
  | "STRIPPER PLATE"
  | "AX PLATE"
  | "BX PLATE"
  // --- three-plate (T series) ---
  | "RUNNER STRIPPER PLATE"
  | "X1 PLATE"
  | "X2 PLATE"
  // --- hot runner ---
  | "MANIFOLD PLATE"
  | "BACKING PLATE"
  // --- die / compression tooling ---
  | "DIE PLATE"
  | "DIE BACKUP PLATE"
  // --- slide-and-cam (SC) hardware plates ---
  | "SC RETAINER PLATE"
  | "SC BACKUP PLATE"
  // --- spacers, when not called rails ---
  | "RISER"
  // A steel solid the shop has to cut that no naming rule recognised.
  // Module6121 tags these with a Steel Part prefix (STEEL_PART_STL_TAG) so they
  // stay distinguishable from the merged whole-base STL, which carries no plate
  // label at all and lands in OTHER.
  | "STEEL PART"
  // BMS / pot-block roles.
  | "HOLDERS"
  | "ID POT"
  | "OD POT"
  | "TCP"
  | "BCP"
  | "ID HOLDER"
  | "OD HOLDER"
  | "HOLDER"
  | "ISO"
  | "OTHER";

/**
 * True when `letter` appears as a standalone word immediately before PLATE.
 *
 * Matches:  "J8420  X  PLATE"   "J8420 Y PLATE"   "PLATE _X_ PLATE"
 * Rejects:  "CAVITY PLATE"      "MATRIX PLATE"    "X1 PLATE"   "AX PLATE"
 *
 * The leading `[^A-Z]` guard is the whole point -- it is the difference
 * between a lone "X" plate and the tail of another word. The gap class
 * excludes digits so "X1 PLATE" can never read as an "X" plate.
 */
function isLoneLetterPlate(u: string, letter: "X" | "Y"): boolean {
  return new RegExp(`(^|[^A-Z])${letter}[^A-Z0-9]*PLATE`).test(u);
}

export function classifyComponentFilename(fn: string): ComponentKind {
  const fnu = (fn || "").toUpperCase();
  const u = fnu.replace(/_/g, " ");
  const compact = u.replace(/[^A-Z0-9]/g, "");

  // Whole base + every part assembly (overlay the entire job).
  if (
    compact.includes("FULLASSEMBLY") ||
    compact.includes("WHOLEASSEMBLY") ||
    compact.includes("FULLBASE") ||
    compact.includes("WHOLEBASE")
  ) {
    return "FULL ASSEMBLY";
  }

  // --- Standard mold-base plates ---------------------------------------
  // These run BEFORE the pot-block rules on purpose. A standard base's
  // "Top Clamp Plate" would otherwise be swallowed by the TCP rule's
  // "TOP CLAMP" test and pollute the six-part pot-block signature. The
  // compact forms are distinct: TOPCLAMPPLATE (standard) never matches
  // inside TOPCLAMPINGPLATE (pot block), and vice versa.
  if (compact.includes("TOPCLAMPPLATE")) return "TOP CLAMP PLATE";
  if (compact.includes("BOTTOMCLAMPPLATE") || compact.includes("BOTCLAMPPLATE")) {
    return "BOTTOM CLAMP PLATE";
  }
  if (compact.includes("SCRETAINERPLATE")) return "SC RETAINER PLATE";
  if (compact.includes("SCBACKUPPLATE")) return "SC BACKUP PLATE";

  // --- three-plate (T series) --------------------------------------------
  // RUNNER STRIPPER before plain STRIPPER: on a T-series base both exist and
  // "RUNNERSTRIPPERPLATE" contains "STRIPPERPLATE", so the general test would
  // swallow the specific one and collapse two different plates into one.
  if (compact.includes("RUNNERSTRIPPER") || compact.includes("RUNNERPLATE")) {
    return "RUNNER STRIPPER PLATE";
  }
  if (compact.includes("X1PLATE") || compact.includes("PLATEX1") || u.includes("X-1 PLATE")) {
    return "X1 PLATE";
  }
  if (compact.includes("X2PLATE") || compact.includes("PLATEX2") || u.includes("X-2 PLATE")) {
    return "X2 PLATE";
  }

  // --- hot runner ---------------------------------------------------------
  // "Manifold Backing Plate" is a backing plate, not the manifold plate -- it
  // is the plain steel plate that BACKS the manifold. It has to be tested
  // before the bare MANIFOLD rule or the more specific name loses.
  if (compact.includes("MANIFOLDBACKING") || compact.includes("MANIFOLDBACKUP")) {
    return "BACKING PLATE";
  }
  if (compact.includes("MANIFOLD")) return "MANIFOLD PLATE";
  if (compact.includes("BACKINGPLATE")) return "BACKING PLATE";

  // --- die tooling --------------------------------------------------------
  // Backup before plain DIE, same swallowing problem.
  if (compact.includes("DIEBACKUP") || compact.includes("DIEBACKUPPLATE")) {
    return "DIE BACKUP PLATE";
  }
  if (compact.includes("DIEPLATE")) return "DIE PLATE";

  if (compact.includes("STRIPPERPLATE")) return "STRIPPER PLATE";
  if (compact.includes("SUPPORTPLATE")) return "SUPPORT PLATE";
  // Longest first: BOTTOMEJECTORPLATE contains EJECTORPLATE. The retainer /
  // backup spellings are Module6121's deprecated aliases for the same plate
  // (see AiPlateNameForRole + roles.ROLE_LABELS).
  if (
    compact.includes("BOTTOMEJECTORPLATE") ||
    compact.includes("BOTEJECTORPLATE") ||
    compact.includes("EJECTORRETAINERPLATE") ||
    compact.includes("EJECTORBACKUPPLATE")
  ) {
    return "BOTTOM EJECTOR PLATE";
  }
  if (compact.includes("EJECTORPLATE")) return "EJECTOR PLATE";
  if (compact.includes("PINPLATE")) return "PIN PLATE";

  // Rails go by several names. RISER and SPACER BLOCK are the same part on a
  // PCS base; PARALLEL is the older shop word for it.
  if (u.includes("RAIL")) return "RAILS";
  if (compact.includes("SPACERBLOCK") || compact.includes("PARALLEL")) return "RAILS";
  if (compact.includes("RISER")) return "RISER";

  // --- X-series (5X / 6X stripper) ---------------------------------------
  // AX / BX BEFORE plain A / B, or "AXPLATE" would match the "APLATE" test and
  // a stripper base's AX plate would be quoted as an A plate.
  if (compact.includes("AXPLATE") || compact.includes("PLATEAX")) return "AX PLATE";
  if (compact.includes("BXPLATE") || compact.includes("PLATEBX")) return "BX PLATE";
  // Module6121 spells these "X" Plate / "Y" Plate with literal quotes, and
  // CleanFileName rewrites a quote to an underscore, so what actually lands on
  // disk is J8420__X_ Plate.STL -> "J8420  X  PLATE.STL" here.
  //
  // The letter has to stand ALONE. Requiring a non-letter in front of it is
  // what keeps "cavitY Plate" and "matriX Plate" out of these two rules -- a
  // loose contains() test on the compacted string was live and wrong:
  // "CAVITYPLATE" ends in Y + PLATE, so every cavity plate on a standard base
  // was classified BX before the cavity rule below ever ran. Excluding digits
  // from the gap is what keeps X1/X2 out (already caught above, belt and
  // braces).
  if (isLoneLetterPlate(u, "X")) return "AX PLATE";
  if (isLoneLetterPlate(u, "Y")) return "BX PLATE";

  // Cavity and core are the customer's words for the A and B plates, and
  // stationary / movable are the same two plates named by which half of the
  // press they ride on. Module6121's StandardPlateNameStd already folds all
  // four into A/B; these keep a raw BOM line classifying the same way, since
  // classifyQuoteComponentName is this same function.
  if (
    compact.includes("CAVITYPLATE") ||
    compact.includes("CAVITYRETAINER") ||
    compact.includes("STATIONARYRETAINER")
  ) {
    return "A PLATE";
  }
  if (
    compact.includes("COREPLATE") ||
    compact.includes("CORERETAINER") ||
    compact.includes("MOVABLERETAINER") ||
    compact.includes("MOVEABLERETAINER")
  ) {
    return "B PLATE";
  }

  // Loosest of the plate rules, so they go last in this block.
  if (compact.includes("APLATE")) return "A PLATE";
  if (compact.includes("BPLATE")) return "B PLATE";

  // Combined ID+OD holders.
  if (
    fnu.includes("_HOLDERS_") ||
    (compact.includes("HOLDERS") && !compact.includes("IDHOLDER") && !compact.includes("ODHOLDER"))
  ) {
    return "HOLDERS";
  }

  if (compact.includes("IDPOT") || (u.includes("POT") && u.includes("ID") && !u.includes("OD"))) {
    return "ID POT";
  }

  if (compact.includes("ODPOT") || (u.includes("POT") && (u.includes("OD") || u.includes("BOT")))) {
    return "OD POT";
  }

  if (u.includes("TCP") || compact.includes("TOPSMED") || u.includes("TOP CLAMP")) {
    return "TCP";
  }

  if (u.includes("BCP") || compact.includes("BOTTOMSMED") || u.includes("BOTTOM CLAMP") || u.includes("BOT CLAMP")) {
    return "BCP";
  }

  if (compact.includes("IDHOLDER") || (u.includes("HOLDER") && u.includes("ID") && !u.includes("OD"))) {
    return "ID HOLDER";
  }

  if (compact.includes("ODHOLDER") || (u.includes("HOLDER") && (u.includes("OD") || u.includes("BOT")))) {
    return "OD HOLDER";
  }

  if (u.includes("HOLDER")) return "HOLDER";
  if (u.includes("ISO")) return "ISO";

  // Last real rule. Module6121's STEEL_PART_STL_TAG -- a solid that has to be
  // cut but that nothing above could name. Deliberately below every specific
  // rule, so a tagged part still classifies as an A Plate or a Riser if its CAD
  // name says so.
  if (compact.includes("STEELPART")) return "STEEL PART";

  return "OTHER";
}

// Quote line-item component names ("ID Holder", "OD Pot", "TCP", ...) run
// through the exact same rules, so a priced line and a 3D file land in the
// same bucket whenever they refer to the same physical part.
export const classifyQuoteComponentName = classifyComponentFilename;

// Display order for the Matching tab's part pills and the 3D tab's gallery.
// Standard plates are listed top-of-stack down, the way a mold is built, so
// scanning the pills reads like walking down the mold base.
export const KIND_ORDER: ComponentKind[] = [
  // Top of the stack downwards, the way the mold is built. On a T-series the
  // runner stripper and X plates sit between the clamp and the B plate; on an
  // X-series the stripper sits between AX and BX.
  "TOP CLAMP PLATE",
  "MANIFOLD PLATE",
  "BACKING PLATE",
  "RUNNER STRIPPER PLATE",
  "X1 PLATE",
  "X2 PLATE",
  "SC RETAINER PLATE",
  "SC BACKUP PLATE",
  "AX PLATE",
  "A PLATE",
  "STRIPPER PLATE",
  "B PLATE",
  "BX PLATE",
  "DIE PLATE",
  "DIE BACKUP PLATE",
  "SUPPORT PLATE",
  "RAILS",
  "RISER",
  "EJECTOR PLATE",
  "BOTTOM EJECTOR PLATE",
  "PIN PLATE",
  "BOTTOM CLAMP PLATE",
  "TCP",
  "BCP",
  "ID HOLDER",
  "OD HOLDER",
  "ID POT",
  "OD POT",
  "HOLDERS",
  "HOLDER",
  // Unnamed steel last among the real parts -- present and costed, but it
  // should not push a named plate down the gallery.
  "STEEL PART",
  "FULL ASSEMBLY",
  "ISO",
  "OTHER",
];

/**
 * Kinds that are not one individual machined part.
 *
 * FULL ASSEMBLY is the merged base, ISO is a render, OTHER is unclassified, and
 * HOLDERS is the combined ID+OD model -- counting it would double up with the
 * two single holders.
 */
const NON_PART_KINDS = new Set<ComponentKind>(["FULL ASSEMBLY", "ISO", "OTHER", "HOLDERS"]);

/**
 * Every kind that represents a real part to measure, cost and machine, in
 * stack order.
 *
 * Derived from KIND_ORDER rather than hand-listed, so adding a plate kind to
 * the union and to KIND_ORDER is enough -- it cannot be forgotten here. That
 * omission was live and expensive: the Geometry tab and the machining
 * estimator both filtered on a hardcoded list of the six BMS pot-block roles,
 * so a standard or PCS base (A, B, T, AX, 5X, 6X) had every one of its plates
 * classified correctly and then dropped, reporting no geometry and no
 * machining time at all.
 */
export const MACHINED_PART_KINDS: ComponentKind[] = KIND_ORDER.filter(
  (k) => !NON_PART_KINDS.has(k),
);
