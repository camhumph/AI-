import argparse
import csv
import json
import re
import shutil
import socket
import subprocess
import sys
import time
from pathlib import Path

# Two-pass STL naming. Imported through a try/except because this file is run both
# as a script (python geometry_classifier\qwen_classify_xt_csv.py ...) and as a
# module by the web app, and only the second form has the package on sys.path.
try:
    from . import stl_plate_naming as _stl_naming
except ImportError:  # pragma: no cover - script invocation
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from geometry_classifier import stl_plate_naming as _stl_naming


BASE = Path(__file__).resolve().parent
KNOWLEDGE = BASE / "mold_geometry_knowledge.md"
VENDOR_KNOWLEDGE = BASE / "vendor_knowledge_sources.md"
# PCS is what most non-BMS work sits on, and the series (A/B/T/AX/5X/6X) changes
# the expected stack order -- a T-series has TWO parting lines, so naming keyed on
# one of them inverts. Loaded alongside the generic vendor notes.
PCS_KNOWLEDGE = BASE / "pcs_mold_base_reference.md"
OUT_DIR = BASE / "outputs"
OUT_DIR.mkdir(parents=True, exist_ok=True)


SHORT_RULES = """
Exact shop-standard tokens in imported STEP/CAD component names are STRONG
ANCHOR EVIDENCE, not weak notes. When a component name contains an exact shop
token, trust it over generic bounding-box geometry:
  A-PLATE / A_PLATE            -> a_plate
  B-PLATE / B_PLATE            -> b_plate
  SC-RETAINER-PLATE            -> sc_retainer_plate
  SC-BACKUP-PLATE              -> sc_backup_plate
  CLAMP-PLATE                  -> bottom_clamp_plate
  EJ-RET-PLATE                 -> ejector_plate (thinner ejector-stack plate)
  EJ-BACKUP-PLATE              -> bottom_ejector_plate (thicker/lower ejector-stack plate)
  RAIL / RAIL-TOP / RAIL-BOTTOM -> rail
  LDR-PIN                      -> leader_pin
  LBB                          -> leader_pin_bushing
  PLC75 / LATCH-LOCK / SAFETY-STRAP (any spelling) -> latch_lock
Only fall back to pure bounding-box geometry when names are generic or missing
(e.g. "plate", "block", stale/copied names, McMaster/DME/PCS catalog numbers
with no shop prefix).

Plate naming: always use "A Plate" / "B Plate". Never use cavity_plate/core_plate.

Bottom-up stack anchoring: decide the bottom of the stack from the rails and
ejector-stack plates first. Leader pins and bushings only decide orientation
when rails/ejector plates are missing or ambiguous. Do not let leader-pin
direction flip a stack orientation that rails/ejector plates already establish.
On plate-sequenced/SC bases, leader pins can seat in the B-plate area and run
upward toward the A-side (reversed pins) -- this must never flip a confirmed
A/B assignment.

Ejector stack naming: the thinner ejector-stack plate is always "ejector_plate".
The thicker/lower ejector-stack plate is "bottom_ejector_plate" -- never call
the thinner plate "ejector_retainer_plate".

Latch-lock / sequenced bases: if any component name contains PLC, LATCH-LOCK,
SAFETY-STRAP (or "SAFTEY-STRAP"), or a Progressive Components latch assembly
name, the base is a plate-sequenced/latch-lock standard base. Latch locks mark
secondary parting/opening lines between plates; they do not set guide
direction. Leader pins set guide direction only. A "weird" standard base can
be missing its Top Clamp Plate -- if the first full-footprint plate is
noticeably thicker than the inner stack, call it a_plate and note the top
clamp plate as missing rather than forcing a 5-plate pattern.

Quote-row mapping: ejector-stack and pin-plate rows must never be merged or
mapped into the a_plate row.

PCS series (most non-BMS work sits on a PCS base). Six series exist: A, B, T,
AX, 5X, 6X. COUNT THE FULL-FOOTPRINT PLATES between the top clamp plate and the
B plate before naming anything -- that count identifies the series, and the
series fixes the stack order:
  0 extra full plates            -> A-series (or B-series if no support plate)
  2 extra, above the cavity      -> T-series: X-1 (runner stripper) then X-2
  1 extra, between AX and BX     -> stripper series (AX / 5X / 6X)
T-series has TWO parting lines. The FIRST opens between X-1 and X-2 to break the
part off the gate; the main one opens after. So a single-parting-line assumption
will invert the naming on a T-series -- anchor on rails and the ejector stack
(bottom-up) as above, never on "the" parting line.
X-2 IS the cavity plate: it quotes on the A-plate row, not a row of its own.
X-1 is a separate plate with its own row.
A hot-runner base adds a manifold plate and a manifold BACKING plate above the
A plate. "Manifold Backing Plate" is a backing plate, not the manifold plate.

PCS item numbers are parseable and are strong evidence of plate thickness:
  <nominal size><series>-<A thk code>-<B thk code>     e.g. 1016A-13-37
Thickness codes are whole inches plus a trailing 3 (=3/8") or 7 (=7/8"):
  13 = 1-3/8"   17 = 1-7/8"   23 = 2-3/8"   37 = 3-7/8"   57 = 5-7/8"
Nominal size is NOT actual: 1016 means 9-7/8 x 16", the width rounded UP to the
next whole inch. Do not reject a 9.875" plate as not matching a "10" base.

Analyze the whole mold first:
- Find the stack axis from full-footprint plates.
- Full-footprint plates have nearly the same width/length as the largest base footprint.
- For a 5-full-plate standard stack, sorted top to bottom:
  1 top_clamp_plate
  2 a_plate
  3 b_plate
  4 support_plate
  5 bottom_clamp_plate
- For a plate-sequenced/SC stack with no top clamp, sorted top to bottom:
  1 a_plate
  2 b_plate
  3 sc_retainer_plate
  4 sc_backup_plate
  5 bottom_clamp_plate
- Rails are long narrow side blocks near the ejector side.
- Pin/ejector plates are long narrower plates inside/between rails.
- Leader pins are long round pins.
- Leader/shoulder bushings are short round cylinders near leader-pin locations.
- Support pillars are large round posts and are not leader pins.
- Return/ejector pins are smaller long round pins.
- The parting line is between a_plate and b_plate. Latch-lock/sequenced bases
  can also have secondary parting/opening lines near sc_retainer_plate/sc_backup_plate.
"""


ROLES = [
    "top_clamp_plate",
    # A full-footprint plate in the injection half that is NOT the cavity: on a
    # hot-runner base it carries the manifold and the runner cross-drilling.
    # Without this role the plate had nowhere to go, so counting plates down from
    # the top slid a_plate/b_plate/support_plate one position each. See
    # _detect_ab_by_facing_gap.
    "manifold_plate",
    "a_plate",
    "b_plate",
    "stripper_plate",
    "sc_retainer_plate",
    "sc_backup_plate",
    "support_plate",
    "bottom_clamp_plate",
    "full_footprint_plate",
    "rail",
    "rail_1",
    "rail_2",
    "pin_plate",
    "ejector_plate",
    "bottom_ejector_plate",
    "ejector_retainer_plate",  # deprecated alias, kept for legacy CORRECT_ME.csv corrections
    "ejector_backup_plate",
    "latch_lock",
    "leader_pin",
    "leader_pin_bushing",
    "guided_ejector_bushing",
    "return_pin",
    "ejector_pin",
    "support_pillar",
    "pullcore",
    "insert_or_core_detail",
    "hardware_other",
    "ignore",
]


def safe_float(value):
    try:
        return float(str(value).strip())
    except Exception:
        return 0.0


def read_rows(csv_path, include_names=False):
    with Path(csv_path).open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        rows = list(csv.DictReader(f))
    compact = []
    for row in rows:
        item = {
            "i": row.get("Index", ""),
            "t": safe_float(row.get("Thickness")),
            "w": safe_float(row.get("Width")),
            "l": safe_float(row.get("Length")),
            "v": safe_float(row.get("BBoxVolume_cuin")),
            "x": safe_float(row.get("CenterX")),
            "y": safe_float(row.get("CenterY")),
            "z": safe_float(row.get("CenterZ")),
            # B-rep hole signature from Module6121's MeasureHoleSignaturesForPlates.
            # Absent on jobs exported before that pass existed, so every reader
            # must treat 0 as "not measured" rather than "measured as none".
            "nthru": int(safe_float(row.get("NThruHoles"))),
            "ncbore": int(safe_float(row.get("NCbore"))),
            "ncross": int(safe_float(row.get("NCrossAxis"))),
            "maxbore": safe_float(row.get("MaxBoreDia")),
            "holesig": (row.get("HoleSig") or "").strip(),
            "npockets": int(safe_float(row.get("NPockets"))),
            "pocketarea": safe_float(row.get("PocketAreaIn2")),
            "pocketdepth": safe_float(row.get("MaxPocketDepth")),
            # The same pockets split by which face they were cut from, along the
            # part's thickness axis. "up" = the +axis face, "dn" = the -axis face,
            # and the axis is always a POSITIVE unit vector, so for a plate lying
            # flat in the stack "up" is the face toward the plate above it.
            #
            # This is what identifies the A/B pair: the cavity and the core open
            # toward each other across the parting line, so the A plate has a big
            # recess on its DOWN face and the B plate a big recess on its UP face.
            # No other adjacent pair in a mold base has large openings facing one
            # another -- that gap is where the moulded part sits.
            "pocketarea_up": safe_float(row.get("PocketAreaUpIn2")),
            "pocketarea_dn": safe_float(row.get("PocketAreaDnIn2")),
            "pocketdepth_up": safe_float(row.get("PocketDepthUp")),
            "pocketdepth_dn": safe_float(row.get("PocketDepthDn")),
            # Solid volume as a % of the bounding box. The single strongest cheap
            # plate discriminator measured so far: clamp plates and rails land
            # ~92-95%, an A plate with a cavity ~74%, a deeply cored B plate ~50%.
            # 0 means "not measured" (older export, or a part the pass skipped).
            "fillpct": safe_float(row.get("SolidFillPct")),
        }
        if include_names:
            item["name"] = row.get("Component", "")[:80]
        compact.append(item)
    return compact


def build_prompt(rows, csv_path, long_knowledge=False):
    if long_knowledge:
        knowledge = KNOWLEDGE.read_text(encoding="utf-8")
        parts = []
        if VENDOR_KNOWLEDGE.exists():
            parts.append(VENDOR_KNOWLEDGE.read_text(encoding="utf-8"))
        if PCS_KNOWLEDGE.exists():
            parts.append(PCS_KNOWLEDGE.read_text(encoding="utf-8"))
        vendor_knowledge = "\n\n".join(parts)
    else:
        knowledge = SHORT_RULES
        vendor_knowledge = ""
    role_list = ", ".join(ROLES)
    return f"""
/no_think
You are a CMS mold-base geometry interpreter.

Your job:
1. Analyze the whole mold first.
2. Decide the stack axis.
3. Decide the full-footprint stack order.
4. Use leader pins, bushings, rails, ejector stack, and dimensions to identify the A plate, B plate, parting line, and part names.
5. Exact shop-standard tokens in CAD component names (e.g. A-PLATE, B-PLATE, SC-RETAINER, SC-BACKUP, EJ-RET, EJ-BACKUP, RAIL, LDR-PIN, LBB, PLC75, LATCH-LOCK, SAFETY-STRAP) are STRONG anchor evidence. Trust them over generic geometry.
6. Only fall back to pure geometry when names are generic, stale, or missing.
7. Return JSON only. No explanation outside JSON. No markdown.

Allowed roles:
{role_list}

Knowledge:
{knowledge}

Vendor/DME/PCS knowledge:
{vendor_knowledge}

Return ONLY valid JSON in this exact shape:
{{
  "job_analysis": {{
    "stack_axis": "CenterX|CenterY|CenterZ",
    "parting_line": "short explanation",
    "rules_for_this_job": [
      "rule 1",
      "rule 2"
    ]
  }},
  "classifications": [
    {{
      "index": "1",
      "role": "a_plate",
      "confidence": "HIGH|MEDIUM|LOW",
      "reason": "geometry-only reason",
      "quote": true
    }}
  ]
}}

CSV source:
{csv_path}

Rows:
{json.dumps(rows, separators=(",", ":"))}
"""


def extract_json(text):
    # Try the text exactly as it came first. When the request went through the
    # HTTP API with a schema, the response IS the JSON object, and the salvage
    # rules below would only put it at risk -- the backtick strip in particular
    # would eat the contents of any reason string that happens to contain one.
    stripped = text.strip()
    if stripped.startswith("{"):
        try:
            return json.loads(stripped)
        except ValueError:
            pass  # genuinely malformed; fall through to the salvage path

    # Ollama/Qwen can emit terminal control characters or thinking text. Strip
    # those first, then extract the first JSON object.
    text = re.sub(r"\x1b\[[0-9;?]*[A-Za-z]", "", text)
    text = text.replace("\b", "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S | re.I)
    text = re.sub(r"`[^`]*`", "", text, flags=re.S)
    text = re.sub(r"(?i)done thinking\.\s*", "", text)
    text = re.sub(r"^Thinking\.\.\..*?(?=\{)", "", text, flags=re.S)
    text = text.strip()
    if text.startswith("{"):
        return json.loads(text)
    match = re.search(r"\{.*\}", text, flags=re.S)
    if not match:
        raise ValueError("No JSON object found in model output")
    return json.loads(match.group(0))


def close_enough(a, b, tolerance=0.08):
    if a == 0 or b == 0:
        return False
    return abs(a - b) <= max(abs(a), abs(b)) * tolerance


def round_bar_diameter(row):
    dims = [row["t"], row["w"], row["l"]]
    pairs = [(dims[0], dims[1], dims[2]), (dims[0], dims[2], dims[1]), (dims[1], dims[2], dims[0])]
    for a, b, long_axis in pairs:
        if close_enough(a, b, 0.06):
            return min(a, b), long_axis
    return 0.0, 0.0


def name_key(row):
    return str(row.get("name", "")).upper().replace("\\", "/")


def _shop_token_locked(roles: dict, idx: str) -> bool:
    """True when a strong shop-name token already assigned this index."""
    if idx not in roles:
        return False
    _role, conf, reason, _quote = roles[idx]
    return conf == "HIGH" and "shop token" in str(reason).lower()


def _set_role_if_unlocked(roles: dict, idx: str, role: str, conf: str, reason: str, quote: bool) -> None:
    if not _shop_token_locked(roles, idx):
        roles[idx] = (role, conf, reason, quote)


LATCH_LOCK_TOKENS = (
    "LATCH-LOCK",
    "LATCH_LOCK",
    "SAFETY-STRAP",
    "SAFETY_STRAP",
    "SAFTEY-STRAP",  # observed misspelling in real shop CAD names (e.g. T001015)
    "SAFTEY_STRAP",
    "PLC75",
)

LATCH_LOCK_REGEX = re.compile(r"\bPLC\d")


def is_latch_lock_name(name):
    if any(token in name for token in LATCH_LOCK_TOKENS):
        return True
    return bool(LATCH_LOCK_REGEX.search(name))


def apply_strong_shop_name_hints(rows, roles):
    """Use exact imported shop tokens as primary anchor evidence.

    These are deliberate shop/vendor-standard naming tokens on STEP imports
    (e.g. A-PLATE, B-PLATE, LDR-PIN). They take priority over generic
    bounding-box geometry rules. Generic/stale/copied names still fall through
    to geometry-only classification below.
    """
    for row in rows:
        idx = str(row["i"])
        name = name_key(row)
        if not name:
            continue

        if is_latch_lock_name(name):
            roles[idx] = (
                "latch_lock",
                "HIGH",
                "Strong shop token (PLC/LATCH-LOCK/SAFETY-STRAP); plate-sequenced latch-lock hardware marking a secondary parting/opening line. Does not set guide direction and must not flip A/B plate assignment.",
                False,
            )
        elif "A-PLATE" in name or "A_PLATE" in name:
            roles[idx] = ("a_plate", "HIGH", "Strong shop token A-PLATE, confirmed as a full-footprint mold plate.", True)
        elif "B-PLATE" in name or "B_PLATE" in name:
            roles[idx] = ("b_plate", "HIGH", "Strong shop token B-PLATE, confirmed as a full-footprint mold plate.", True)
        elif "SC-RETAINER-PLATE" in name or "SC_RETAINER_PLATE" in name:
            roles[idx] = ("sc_retainer_plate", "HIGH", "Strong shop token SC-RETAINER-PLATE in a special standard mold-base stack.", True)
        elif "SC-BACKUP-PLATE" in name or "SC_BACKUP_PLATE" in name:
            roles[idx] = ("sc_backup_plate", "HIGH", "Strong shop token SC-BACKUP-PLATE in a special standard mold-base stack.", True)
        elif "CLAMP-PLATE" in name or "CLAMP_PLATE" in name:
            roles[idx] = ("bottom_clamp_plate", "HIGH", "Strong shop token CLAMP-PLATE on the ejector/clamp side.", True)
        elif "EJ-BACKUP-PLATE" in name or "EJ_BACKUP_PLATE" in name:
            roles[idx] = ("bottom_ejector_plate", "HIGH", "Strong shop token EJ-BACKUP-PLATE; thicker/lower backing plate in ejector stack. CMS naming: Bottom Ejector Plate (never Ejector Retainer Plate).", True)
        elif "EJ-RET-PLATE" in name or "EJ_RET_PLATE" in name:
            roles[idx] = ("ejector_plate", "HIGH", "Strong shop token EJ-RET-PLATE; thinner plate in ejector stack. CMS naming: Ejector Plate.", True)
        elif ("EJECTOR PLATE" in name or "EJECTOR-PLATE" in name) and "BACKUP" not in name:
            roles[idx] = ("ejector_plate", "HIGH", "Strong shop token EJECTOR PLATE in component name.", True)
        elif "RAIL-" in name or "_RAIL" in name or "/RAIL" in name:
            roles[idx] = ("rail", "HIGH", "Strong shop token RAIL; long side rail/support block.", True)
        elif "LDR-PIN" in name or "LDR_PIN" in name:
            roles[idx] = ("leader_pin", "HIGH", "Strong shop token LDR-PIN; primary leader pin set.", False)
        elif "/LBB_" in name or "LBB_" in name:
            roles[idx] = ("leader_pin_bushing", "HIGH", "Strong shop token LBB; leader pin bushing.", False)


def looks_like_pot_block_geometry(rows) -> bool:
    """Detect BMS / Tempcraft pot-block stacks that must NOT get A/B/rail roles.

    Signature: ~2 thin full-footprint clamps, >=2 thick non-full holders,
    plus distinguishable pot cubes (thick, chunky, footprint << mold) and/or
    0.25\" insulation with at least one pot. Generic asm_objects names still match.
    Full-size A/B plates are never pots.
    """
    if not rows or len(rows) < 6:
        return False
    max_w = max(r["w"] for r in rows)
    max_l = max(r["l"] for r in rows)
    max_fp = max(r["w"] * r["l"] for r in rows)
    full_thin = [
        r for r in rows
        if r["w"] >= max_w * 0.85 and r["l"] >= max_l * 0.85 and 0.75 <= r["t"] <= 2.5
    ]
    thick_inner = [
        r for r in rows
        if r["t"] >= 3.0 and (r["w"] * r["l"]) < max_fp * 0.85 and (r["w"] * r["l"]) >= max_fp * 0.15
    ]
    thin_sheets = [r for r in rows if abs(r["t"] - 0.25) <= 0.06]

    def _is_pot(r) -> bool:
        t, w, l = r["t"], r["w"], r["l"]
        if t < 3.0 or w <= 0 or l <= 0:
            return False
        if (l / w) > 1.7:
            return False
        fp = w * l
        if fp >= 0.55 * max_fp:
            return False
        dim_max = max(t, w, l)
        dim_min = min(t, w, l)
        return dim_max > 0 and (dim_min / dim_max) >= 0.35

    pot_like = [r for r in rows if _is_pot(r)]
    full_plates = [
        r for r in rows
        if r["w"] >= max_w * 0.85 and r["l"] >= max_l * 0.85 and r["t"] >= 0.5
    ]
    if len(full_plates) >= 5:
        return False
    return (
        len(full_thin) <= 2
        and len(thick_inner) >= 2
        and (len(pot_like) >= 2 or (len(thin_sheets) >= 2 and len(pot_like) >= 1))
    )


# A recess has to be this deep, and this much of the plate footprint, before it
# counts as a mold opening rather than a screw seat or a relief cut.
AB_GAP_MIN_DEPTH_IN = 0.10
AB_GAP_MIN_AREA_FRAC = 0.04


def _detect_ab_by_facing_gap(full_plates, notes=None):
    """Find the A/B pair by the opposed recesses that form the mold cavity.

    `full_plates` must already be sorted top-to-bottom along the stack axis.
    Returns (index_of_upper_plate, evidence_text) where the index is into
    `full_plates` and identifies the A plate, or None when the geometry does not
    show a facing gap -- an unmeasured job, or a base whose cavity lives in
    separate inserts rather than in the plates.

    Deliberately conservative: returning None falls back to positional naming,
    which is what shipped before, so a job this cannot read is no worse off.
    """
    if len(full_plates) < 2:
        return None

    best = None
    best_score = 0.0
    for k in range(len(full_plates) - 1):
        upper = full_plates[k]
        lower = full_plates[k + 1]

        # The A plate's cavity opens DOWN, toward the B plate below it.
        # The B plate's core opens UP, toward the A plate above it.
        up_area = upper.get("pocketarea_dn", 0.0)
        up_deep = upper.get("pocketdepth_dn", 0.0)
        lo_area = lower.get("pocketarea_up", 0.0)
        lo_deep = lower.get("pocketdepth_up", 0.0)

        # Both sides must clear the gates. min() is the whole test: one hollow
        # plate beside a solid one is not a mold gap.
        if min(up_deep, lo_deep) < AB_GAP_MIN_DEPTH_IN:
            continue
        gate_u = max(1.0, upper["w"] * upper["l"] * AB_GAP_MIN_AREA_FRAC)
        gate_l = max(1.0, lower["w"] * lower["l"] * AB_GAP_MIN_AREA_FRAC)
        if up_area < gate_u or lo_area < gate_l:
            continue

        score = min(up_area, lo_area) * min(up_deep, lo_deep)
        if score > best_score:
            best_score = score
            best = k

    if best is None:
        return None

    a_row = full_plates[best]
    b_row = full_plates[best + 1]
    evidence = (
        "A/B pair identified by the opposed pocket faces that form the mold gap: "
        f"idx {a_row['i']} opens downward ({a_row.get('pocketarea_dn', 0.0):.1f} in2, "
        f"{a_row.get('pocketdepth_dn', 0.0):.3f} deep) onto "
        f"idx {b_row['i']} opening upward ({b_row.get('pocketarea_up', 0.0):.1f} in2, "
        f"{b_row.get('pocketdepth_up', 0.0):.3f} deep)."
    )
    return best, evidence


def _looks_like_hot_runner(plate):
    """Is there real evidence this plate carries a hot-runner manifold?

    A manifold plate is drilled ACROSS for the runner channels and bored for the
    nozzle/sprue -- that cross-drilling is the whole point of the plate. Anything
    without it is a plate that happens to sit above the cavity, which is a
    different thing.
    """
    cross = safe_float(plate.get("NCrossAxis", 0))
    bore = safe_float(plate.get("MaxBoreDia", 0))
    thick = safe_float(plate.get("Thickness", 0))
    return cross >= 6 and bore >= 1.0 and thick >= 1.5


def _stack_names_around_ab(full_plates, a_at, top_clamp_present):
    """Name the whole full-footprint stack outward from a known A/B pair.

    Above the A plate: the topmost plate is the top clamp. A plate between it and
    the A plate is a manifold plate ONLY when it looks like one. Below the B
    plate: support plate(s), with the lowest being the bottom clamp.

    The manifold gate matters because manifold_plate was being handed out on
    position alone. On C18027 that produced a manifold this shop does not run,
    and -- because the roles are assigned outward from A/B -- it pushed the real
    A plate down to b_plate and the real B plate down to support_plate. One
    unjustified name silently re-labelled three plates.

    When the evidence is absent the more likely reading is that the A/B pair was
    found one position too low, so A is re-anchored onto that plate instead.
    """
    n = len(full_plates)

    # Re-anchor before naming: an unevidenced manifold means A/B sits too low.
    first_named = 1 if top_clamp_present else 0
    while a_at > first_named and not _looks_like_hot_runner(full_plates[a_at - 1]):
        a_at -= 1

    names = [None] * n
    names[a_at] = "a_plate"
    names[a_at + 1] = "b_plate"

    # --- above the A plate ---
    if a_at > 0:
        start = 0
        if top_clamp_present:
            names[0] = "top_clamp_plate"
            start = 1
        for k in range(start, a_at):
            names[k] = "manifold_plate"

    # --- below the B plate ---
    below = list(range(a_at + 2, n))
    if below:
        names[below[-1]] = "bottom_clamp_plate"
        for k in below[:-1]:
            names[k] = "support_plate"

    return [nm if nm else "full_footprint_plate" for nm in names]


def classify_geometry(rows):
    """Rule-based fallback for when the LLM does not return valid JSON."""
    if not rows:
        return {"job_analysis": {}, "classifications": []}

    max_w = max(r["w"] for r in rows)
    max_l = max(r["l"] for r in rows)
    stack_axis_key = "y"
    stack_axis = "CenterY"
    roles = {}
    apply_strong_shop_name_hints(rows, roles)

    # HARD GUARD: pot-block / BMS geometry must never invent A Plate / B Plate / Rails.
    # Module6121 owns those jobs via BOM (TCP, ID/OD Holder, ID/OD Pot, BCP).
    if looks_like_pot_block_geometry(rows):
        return {
            "job_analysis": {
                "stack_axis": stack_axis,
                "base_type": "bms",
                "rules_for_this_job": [
                    "Pot-block / BMS geometry detected — skipped standard A/B/rail classify. "
                    "Use Module6121 BOM-driven TCP / Holder / Pot / BCP fill."
                ],
            },
            "classifications": [
                {
                    "index": str(r["i"]),
                    "role": "hardware_other",
                    "confidence": "LOW",
                    "reason": "Pot-block job: AI standard-stack roles disabled; macro BOM owns plate naming.",
                    "quote": False,
                }
                for r in rows
            ],
        }

    has_latch_lock = any(is_latch_lock_name(name_key(r)) for r in rows)

    full_plates = [
        r for r in rows
        if r["w"] >= max_w * 0.85 and r["l"] >= max_l * 0.85 and r["t"] >= 0.5
    ]
    if len(full_plates) >= 2:
        ranges = {
            "x": max(r["x"] for r in full_plates) - min(r["x"] for r in full_plates),
            "y": max(r["y"] for r in full_plates) - min(r["y"] for r in full_plates),
            "z": max(r["z"] for r in full_plates) - min(r["z"] for r in full_plates),
        }
        stack_axis_key = max(ranges, key=ranges.get)
        stack_axis = {"x": "CenterX", "y": "CenterY", "z": "CenterZ"}[stack_axis_key]

    full_plates.sort(key=lambda r: r[stack_axis_key], reverse=True)
    n_full = len(full_plates)

    # ── A/B by the gap that faces itself ────────────────────────────────────────
    #
    # Position alone cannot find the A and B plates, and on C17267 it got them
    # wrong. That base is a hot-runner base: TCP / MANIFOLD / A / B, so counting
    # down from the top put a_plate on the manifold plate, b_plate on the A plate,
    # and pushed the real B plate into support_plate. Four names, one shift, all
    # wrong -- and no amount of stack ordering fixes it, because the manifold plate
    # is a full-footprint plate of the same size sitting in the same place.
    #
    # What separates them is what the plates are FOR. The moulded part sits in the
    # gap between the cavity and the core, so the A plate is hollowed out on the
    # face pointing down and the B plate on the face pointing up, and those two
    # recesses face each other across the parting line. Nothing else in a mold base
    # does that: a manifold plate is drilled through but not hollowed, and clamp,
    # support and rail plates carry no opposed pair of pockets at all.
    #
    # So: walk adjacent pairs and score each on the smaller of the two facing
    # openings. Taking the SMALLER is the whole point -- it demands a real recess
    # on BOTH sides, which one deeply pocketed plate next to a flat one cannot fake.
    ab_hit = _detect_ab_by_facing_gap(full_plates)
    ab_pair, ab_evidence = ab_hit if ab_hit else (None, "")

    if n_full >= 4:
        inner = full_plates[1:-1]
        avg_inner_t = sum(r["t"] for r in inner) / len(inner) if inner else 0.0
        top_clamp_present = True
        if avg_inner_t > 0 and full_plates[0]["t"] > avg_inner_t * 1.35:
            top_clamp_present = False

        # FOUR full plates is the most common standard/PCS layout there is:
        # Top Clamp / A / B / Bottom Clamp, with the ejector stack and rails
        # below it. It used to fall through to the generic branch below, which
        # assigns no stack role at all -- so on a plain 4-plate base the AI named
        # nothing, and the bottom clamp plate ended up in "Other Hardware".
        #
        # Deliberately matches Module6121's StdFullPlateName Case 4, so the AI
        # bridge and the macro's own geometry fallback cannot disagree about the
        # same stack.
        if ab_pair is not None and top_clamp_present:
            # Geometry found the mold gap, so build the stack AROUND it instead of
            # counting positions from the top. Everything above the A plate is
            # clamp/manifold, everything below the B plate is support/clamp.
            #
            # Only on the top-clamp-present layout. The `top_clamp_present == False`
            # branch below is this code's stripper-plate signature, and a stripper
            # plate sits BETWEEN the A and B plates -- so A and B are not adjacent
            # there and the facing-gap pair would be A/stripper. That layout has no
            # ground truth behind it yet, so leave it exactly as it was.
            stack_names = _stack_names_around_ab(full_plates, ab_pair, top_clamp_present)
        elif n_full == 4:
            stack_names = (
                ["top_clamp_plate", "a_plate", "b_plate", "bottom_clamp_plate"]
                if top_clamp_present
                else ["a_plate", "b_plate", "support_plate", "bottom_clamp_plate"]
            )
        elif top_clamp_present:
            stack_names = ["top_clamp_plate", "a_plate", "b_plate", "support_plate", "bottom_clamp_plate"]
        else:
            stack_names = ["a_plate", "stripper_plate", "b_plate", "support_plate", "bottom_clamp_plate"]
        if ab_pair is not None and top_clamp_present:
            base_reason = (
                "Full-footprint plate, named outward from the A/B pair that "
                "geometry found. " + ab_evidence
            )
        elif top_clamp_present:
            base_reason = "Full-footprint plate, assigned by top-to-bottom mold stack order."
        else:
            base_reason = (
                "Full-footprint plate, assigned by top-to-bottom mold stack order. "
                "Top clamp missing."
            )
        for name, row in zip(stack_names, full_plates):
            idx = str(row["i"])
            if idx not in roles:
                roles[idx] = (name, "HIGH", base_reason, True)
    else:
        for row in full_plates:
            idx = str(row["i"])
            if idx not in roles:
                roles[idx] = (
                    "full_footprint_plate",
                    "MEDIUM",
                    "Full-footprint plate, but fewer than 4 full plates were found so standard stack role was not forced.",
                    True,
                )

    if len(full_plates) == 2:
        axes = {
            "x": abs(full_plates[0]["x"] - full_plates[1]["x"]),
            "y": abs(full_plates[0]["y"] - full_plates[1]["y"]),
            "z": abs(full_plates[0]["z"] - full_plates[1]["z"]),
        }
        die_axis = max(axes, key=axes.get)
        hi_full, lo_full = sorted(full_plates, key=lambda r: r[die_axis], reverse=True)
        _set_role_if_unlocked(
            roles,
            str(lo_full["i"]),
            "bottom_clamp_plate",
            "HIGH",
            "Two-half mold pattern: lower full-footprint plate along stack axis is BCP.",
            True,
        )
        _set_role_if_unlocked(
            roles,
            str(hi_full["i"]),
            "top_clamp_plate",
            "MEDIUM",
            "Two-half mold pattern: opposite full-footprint clamp plate.",
            True,
        )

        non_full_blocks = [
            r for r in rows
            if str(r["i"]) not in roles
            and not _shop_token_locked(roles, str(r["i"]))
            and r["t"] >= 3.0
            and r["w"] >= max_w * 0.30
            and r["l"] >= max_l * 0.30
        ]
        non_full_blocks.sort(key=lambda r: r["v"], reverse=True)
        if len(non_full_blocks) >= 2:
            high_inner, low_inner = sorted(non_full_blocks[:2], key=lambda r: r[die_axis], reverse=True)
            _set_role_if_unlocked(
                roles,
                str(high_inner["i"]),
                "a_plate",
                "MEDIUM",
                "Two-half mold pattern: larger inner block on high side of stack axis.",
                True,
            )
            _set_role_if_unlocked(
                roles,
                str(low_inner["i"]),
                "b_plate",
                "MEDIUM",
                "Two-half mold pattern: matching inner block on low side of stack axis.",
                True,
            )
            # NOTE: rails/ejector-stack detection above already anchors bottom_pos.
            # a_plate/b_plate assignment here must not be re-derived from leader
            # pin direction; it stays anchored to the rail/ejector-established axis.

        thin_large = [
            r for r in rows
            if str(r["i"]) not in roles
            and r["t"] <= 0.75
            and r["l"] >= max_l * 0.35
            and r["w"] >= max_w * 0.20
        ]
        thin_large.sort(key=lambda r: r["v"], reverse=True)
        if len(thin_large) >= 2:
            rails = thin_large[:2]
            rails.sort(key=lambda r: r[die_axis], reverse=True)
            roles[str(rails[0]["i"])] = ("rail_1", "MEDIUM", "Two-half mold pattern: first largest thin rail/strip plate.", True)
            roles[str(rails[1]["i"])] = ("rail_2", "MEDIUM", "Two-half mold pattern: second largest thin rail/strip plate.", True)
        if len(thin_large) >= 4:
            ejector_candidates = thin_large[2:4]
            ejector_candidates.sort(key=lambda r: r["t"])
            roles[str(ejector_candidates[0]["i"])] = (
                "ejector_plate",
                "MEDIUM",
                "Two-half mold pattern: thinner remaining ejector-stack plate. CMS naming: Ejector Plate.",
                True,
            )
            roles[str(ejector_candidates[1]["i"])] = (
                "bottom_ejector_plate",
                "MEDIUM",
                "Two-half mold pattern: thicker/lower remaining ejector-stack plate. CMS naming: Bottom Ejector Plate (never Ejector Retainer Plate).",
                True,
            )

    top_pos = full_plates[0][stack_axis_key] if full_plates else 0.0
    bottom_pos = full_plates[-1][stack_axis_key] if full_plates else 0.0
    support_pos = next((r[stack_axis_key] for r in full_plates if roles.get(str(r["i"]), ("",))[0] == "support_plate"), 0.0)

    for row in rows:
        idx = str(row["i"])
        if idx in roles:
            continue

        dia, bar_len = round_bar_diameter(row)
        if dia and dia <= 4.0:
            is_long_bar = bar_len >= dia * 1.6
            is_short_cylinder = not is_long_bar and bar_len >= dia * 0.6

            if 2.5 <= dia <= 4.0 and bar_len >= 6.0:
                roles[idx] = ("support_pillar", "HIGH", "Large long round post; support pillar geometry.", False)
            elif is_long_bar and 1.35 <= dia <= 2.2 and bar_len >= 10.0:
                roles[idx] = ("leader_pin", "MEDIUM", "Long smaller round bar; likely leader pin.", False)
            elif is_long_bar and 0.9 <= dia < 1.35 and bar_len >= 8.0:
                roles[idx] = ("return_pin", "MEDIUM", "Long 1-inch-class round pin; likely return/ejector return pin.", False)
            elif is_short_cylinder and 1.0 <= dia <= 2.6 and row[stack_axis_key] >= support_pos:
                roles[idx] = ("leader_pin_bushing", "MEDIUM", "Short round cylinder near guide hardware size.", False)
            elif is_short_cylinder and 1.0 <= dia <= 2.6 and row[stack_axis_key] < support_pos:
                roles[idx] = ("guided_ejector_bushing", "MEDIUM", "Short round cylinder in ejector half; likely guided-ejector bushing.", False)
            else:
                roles[idx] = ("hardware_other", "LOW", "Round hardware but role is not certain from dimensions alone.", False)
            continue

        long_full = row["l"] >= max_l * 0.85
        side_coords = [axis for axis in ("x", "y", "z") if axis != stack_axis_key]
        side_offset = max(abs(row[axis]) for axis in side_coords)
        centered_side = side_offset <= max_w * 0.15
        side_block = side_offset >= max_w * 0.25
        # Rail width.
        #
        # The lower bound used to be a pure fraction of the base width
        # (max_w * 0.18), which silently rejects narrow rails on wide bases:
        # C17267's rails are 4.000" wide against a 23.750" base, and
        # 0.18 * 23.750 = 4.275 -- so they missed by 0.275" and fell through to
        # hardware_other/LOW, which cost the job its Rails quote row AND its
        # Rails STL. A rail is defined by being long, narrow, thick enough to
        # be structural, and offset to one side; its width does not scale with
        # the base the way the lower bound assumed.
        #
        # So: keep a fractional UPPER bound (a rail is never most of the base),
        # but make the lower bound an absolute structural minimum.
        RAIL_MIN_WIDTH_IN = 1.5
        narrow_width = (
            row["w"] >= RAIL_MIN_WIDTH_IN
            and row["w"] <= max_w * 0.72
            and row["l"] >= row["w"] * 2.25   # genuinely slender, not a block
        )
        ejector_width = max_w * 0.58 <= row["w"] <= max_w * 0.86

        axis_pos = row[stack_axis_key]
        if long_full and narrow_width and side_block and bottom_pos <= axis_pos <= support_pos + 1.0:
            roles[idx] = ("rail", "HIGH", "Long narrow full-length side block in ejector/rail zone.", True)
        elif long_full and ejector_width and centered_side and bottom_pos <= axis_pos <= support_pos + 2.0:
            if row["t"] <= 0.8 or axis_pos > bottom_pos + 1.5:
                roles[idx] = ("ejector_plate", "MEDIUM", "Thinner centered ejector-stack plate. CMS naming: Ejector Plate.", True)
            else:
                roles[idx] = ("bottom_ejector_plate", "MEDIUM", "Thicker/lower centered ejector-stack plate. CMS naming: Bottom Ejector Plate (never Ejector Retainer Plate).", True)
        elif axis_pos > support_pos and row["w"] < max_w * 0.75 and row["l"] < max_l * 0.75:
            roles[idx] = ("insert_or_core_detail", "LOW", "Smaller block inside cavity/core area; not a standard full plate.", False)
        elif top_pos >= axis_pos >= bottom_pos:
            roles[idx] = ("hardware_other", "LOW", "Inside mold stack but not enough geometry to name confidently.", False)
        else:
            roles[idx] = ("ignore", "LOW", "Outside main standard-base classification rules.", False)

    # ------------------------------------------------------------------
    # Ejector Retainer vs Back-Up, decided by COUNTERBORE instead of thickness.
    #
    # The rule above splits the pair on "thinner = retainer". That is a shop
    # convention, not a measurement, and on a base where the two plates share a
    # footprint (C18522: both 8.375 x 19.985, differing only 0.518" vs 1.125")
    # thickness is the only thing separating them. The physical difference is that
    # the RETAINER plate is counterbored to seat the ejector pin heads and the
    # BACK-UP plate has plain through-holes. Module6121's B-rep pass measures that
    # directly, so where the holes disagree with thickness, believe the holes.
    #
    # Guards, in order of how badly each would bite:
    #   * needs a real hole signature on both plates -- the NCbore column is absent
    #     on jobs exported before the pass existed, and 0 there means "not
    #     measured", not "measured as none"
    #   * needs a decisive margin, not a 1-hole edge, so facet noise cannot flip it
    #   * never overrides a strong shop token: EJ-RET-PLATE in the CAD name wins
    # ------------------------------------------------------------------
    rules_note_holes = ""
    ej_pair = [
        r for r in rows
        if roles.get(str(r["i"]), ("",))[0] in {"ejector_plate", "bottom_ejector_plate"}
    ]
    if len(ej_pair) == 2 and all(r.get("holesig") for r in ej_pair):
        seated, plain = sorted(ej_pair, key=lambda r: r.get("ncbore", 0), reverse=True)
        n_seated = seated.get("ncbore", 0)
        n_plain = plain.get("ncbore", 0)
        # Decisive: at least 4 seats (a real ejector pattern) and at least double
        # the other plate, so a stray counterbore for a socket screw cannot decide it.
        if n_seated >= 4 and n_seated >= max(2 * n_plain, n_plain + 3):
            by_thickness = {
                str(r["i"]): roles.get(str(r["i"]), ("",))[0] for r in ej_pair
            }
            _set_role_if_unlocked(
                roles,
                str(seated["i"]),
                "ejector_plate",
                "HIGH",
                f"Counterbored ejector-pin seats measured from the solid "
                f"({n_seated} vs {n_plain} on the paired plate): this is the plate the "
                f"pin heads sit in. CMS naming: Ejector Plate.",
                True,
            )
            _set_role_if_unlocked(
                roles,
                str(plain["i"]),
                "bottom_ejector_plate",
                "HIGH",
                f"Paired ejector-stack plate with plain through-holes "
                f"({n_plain} counterbores vs {n_seated}): backing plate behind the "
                f"retainer. CMS naming: Bottom Ejector Plate.",
                True,
            )
            flipped = [
                i for i, old in by_thickness.items()
                if old != roles.get(i, ("",))[0]
            ]
            if flipped:
                rules_note_holes = (
                    "Ejector-stack pair was re-assigned from measured counterbores "
                    "rather than relative thickness."
                )
            else:
                rules_note_holes = (
                    "Ejector-stack pair confirmed by measured counterbores; agrees "
                    "with the thickness rule."
                )

    leader_rows = [r for r in rows if roles.get(str(r["i"]), ("",))[0] == "leader_pin"]
    shoulder_bushings = [r for r in rows if roles.get(str(r["i"]), ("",))[0] == "leader_pin_bushing"]
    ejector_bushings = [r for r in rows if roles.get(str(r["i"]), ("",))[0] == "guided_ejector_bushing"]

    def near_same_axis_plane(a, b, tolerance=0.55):
        xy = abs(a["x"] - b["x"]) <= tolerance and abs(a["y"] - b["y"]) <= tolerance
        xz = abs(a["x"] - b["x"]) <= tolerance and abs(a["z"] - b["z"]) <= tolerance
        yz = abs(a["y"] - b["y"]) <= tolerance and abs(a["z"] - b["z"]) <= tolerance
        return xy or xz or yz

    for row in leader_rows:
        idx = str(row["i"])
        if any(near_same_axis_plane(row, b) for b in shoulder_bushings):
            roles[idx] = (
                "leader_pin",
                "HIGH",
                "Primary leader-pin set: center position matches shoulder bushings that sit on the B plate.",
                False,
            )
        elif any(near_same_axis_plane(row, b) for b in ejector_bushings):
            roles[idx] = (
                "leader_pin",
                "MEDIUM",
                "Secondary guide-pin set: center position matches guided-ejector bushings, so it should not decide A/B plate identity.",
                False,
            )

    long_round_pins = []
    short_round_bushings = []
    for row in rows:
        dia, bar_len = round_bar_diameter(row)
        if not dia or dia > 4.0:
            continue
        if bar_len >= dia * 3.0 and bar_len >= 6.0:
            long_round_pins.append(row)
        elif dia * 0.6 <= bar_len <= dia * 1.35 and 1.0 <= dia <= 2.6:
            short_round_bushings.append(row)

    for pin in long_round_pins:
        matching_bushings = [b for b in short_round_bushings if near_same_axis_plane(pin, b, 0.55)]
        if matching_bushings:
            roles[str(pin["i"])] = (
                "leader_pin",
                "MEDIUM",
                "Long round guide pin matched to short bushing at the same center position plane.",
                False,
            )
            for bushing in matching_bushings:
                roles[str(bushing["i"])] = (
                    "leader_pin_bushing",
                    "MEDIUM",
                    "Short bushing matched to a long guide pin at the same center position plane.",
                    False,
                )

    classifications = []
    for row in rows:
        role, confidence, reason, quote = roles.get(
            str(row["i"]),
            ("ignore", "LOW", "No matching standard-base geometry rule.", False),
        )
        classifications.append(
            {
                "index": str(row["i"]),
                "role": role,
                "confidence": confidence,
                "reason": reason,
                "quote": quote,
            }
        )

    parting = "Between a_plate and b_plate from the full-footprint stack order."
    rules_for_this_job = [
        f"Full-footprint plates were sorted by {stack_axis} from top to bottom.",
        "Rails and the ejector stack anchored the bottom of the stack first; leader-pin direction was not used to flip stack orientation.",
        "Ejector-stack plates were detected as centered long narrower plates near the rails; thinner = ejector_plate, thicker/lower = bottom_ejector_plate.",
        "Round guide hardware was separated by diameter and length.",
        "Exact shop-name tokens (A-PLATE, B-PLATE, SC-RETAINER, SC-BACKUP, EJ-RET, EJ-BACKUP, RAIL, LDR-PIN, LBB) were treated as strong anchors and applied before geometry-only rules.",
    ]
    # Only stated when hole evidence was actually available and decisive, so the
    # rules list never claims a measurement the export did not carry.
    if rules_note_holes:
        rules_for_this_job.append(rules_note_holes)
    if has_latch_lock:
        parting = (
            "Primary parting line between a_plate and b_plate from the full-footprint stack order. "
            "Latch-lock/PLC/safety-strap hardware detected: this is a plate-sequenced/latch-lock standard base with "
            "secondary opening/parting lines at the latch attachment points (not a plain A/B/support stack)."
        )
        rules_for_this_job.append(
            "Latch-lock/PLC/safety-strap tokens were detected (PLC, LATCH-LOCK, SAFETY-STRAP/SAFTEY-STRAP). "
            "This base is plate-sequenced; latch locks mark secondary parting/opening lines and do not set guide "
            "direction. Leader-pin position and any reversed/seated leader pins were not allowed to flip the "
            "A/B plate assignment established by shop-name tokens and the full-footprint stack order."
        )

    return {
        "job_analysis": {
            "stack_axis": stack_axis,
            "parting_line": parting,
            "sequenced_latch_lock_base": has_latch_lock,
            "rules_for_this_job": rules_for_this_job,
        },
        "classifications": classifications,
    }


OLLAMA_URL = "http://127.0.0.1:11434"


def _classification_schema(roles):
    """A JSON schema Ollama enforces during generation.

    Worth doing rather than validating after the fact, because it removes two
    whole failure modes instead of reporting them: the model cannot invent a role
    outside the list (qwen3:8b offered "support_beam" on its first C17880 run),
    and it cannot wrap the answer in prose that then has to be salvaged.
    """
    return {
        "type": "object",
        "required": ["job_analysis", "classifications"],
        "properties": {
            "job_analysis": {
                "type": "object",
                "properties": {
                    "stack_axis": {"type": "string"},
                    "parting_line": {"type": "string"},
                    "rules_for_this_job": {"type": "array", "items": {"type": "string"}},
                },
            },
            "classifications": {
                "type": "array",
                "items": {
                    "type": "object",
                    "required": ["index", "role", "confidence", "reason"],
                    "properties": {
                        "index": {"type": "string"},
                        "role": {"type": "string", "enum": list(roles)},
                        "confidence": {"type": "string", "enum": ["HIGH", "MEDIUM", "LOW"]},
                        "reason": {"type": "string"},
                        "quote": {"type": "boolean"},
                        "measure": {"type": "boolean"},
                        "changed_from_first_pass": {"type": "boolean"},
                    },
                },
            },
            "hardware_corrections": {
                "type": "array",
                "items": {
                    "type": "object",
                    "required": ["index", "role", "reason"],
                    "properties": {
                        "index": {"type": "string"},
                        "role": {"type": "string", "enum": list(roles)},
                        "confidence": {"type": "string", "enum": ["HIGH", "MEDIUM", "LOW"]},
                        "reason": {"type": "string"},
                    },
                },
            },
        },
    }


def run_ollama_api(prompt, model, timeout_minutes, schema=None):
    """Generate through Ollama's HTTP API.

    Preferred over shelling out to `ollama run`, which renders a live streaming
    display and leaves ANSI cursor-movement escapes (\\x1b[1D\\x1b[K) spliced
    through the text -- mid-word, so they corrupt JSON string contents. The API
    returns the completion and nothing else.

    Uses urllib rather than requests: the backend's requirements.txt does not
    carry requests, and this has to run on the shop PC unchanged.
    """
    import urllib.error
    import urllib.request

    body = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        # qwen3 reasons by default and its thinking block is not valid JSON.
        # /no_think in the prompt is advisory; this is not.
        "think": False,
        "options": {"temperature": 0.1, "num_ctx": 32768},
    }
    if schema:
        body["format"] = schema

    req = urllib.request.Request(
        f"{OLLAMA_URL}/api/generate",
        data=json.dumps(body).encode("utf-8"),
        headers={"Content-Type": "application/json"},
    )
    timeout = timeout_minutes * 60 if timeout_minutes > 0 else None
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        payload = json.loads(resp.read().decode("utf-8"))
    return payload.get("response", "")


def ollama_api_available():
    import urllib.request

    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=3) as r:
            return r.status == 200
    except Exception:
        return False


def run_ollama(prompt, model, timeout_minutes):
    if shutil.which("ollama") is None:
        raise SystemExit("Ollama is not installed or not on PATH. Open a new Command Prompt after installing Ollama.")

    print(f"Starting Ollama model: {model}", flush=True)
    print("This can take several minutes on an older CPU. If this is the first run, Ollama may also load/download the model.", flush=True)
    if timeout_minutes <= 0:
        timeout_seconds = None
        print("Timeout disabled. This can run until the model finishes.", flush=True)
    else:
        timeout_seconds = timeout_minutes * 60
        print(f"Timeout set to {timeout_minutes} minutes.", flush=True)
    cmd = ["ollama", "run", model]
    try:
        result = subprocess.run(
            cmd,
            input=prompt,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
        )
    except FileNotFoundError:
        raise SystemExit("Ollama is not installed or not on PATH.")
    except subprocess.TimeoutExpired:
        debug_prompt = OUT_DIR / "last_timeout_prompt.txt"
        debug_prompt.write_text(prompt, encoding="utf-8", errors="replace")
        raise SystemExit(
            f"Ollama timed out after {timeout_minutes} minutes. "
            f"Prompt saved to {debug_prompt}. Try fewer rows or use --timeout-minutes 0."
        )
    if result.returncode != 0:
        raise SystemExit(result.stderr.strip() or "Ollama failed")
    return result.stdout


def original_row_lookup(csv_path):
    with Path(csv_path).open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        return {str(row.get("Index", "")): row for row in csv.DictReader(f)}


def write_outputs(data, csv_path):
    stem = Path(csv_path).stem
    out_json = OUT_DIR / f"{stem}_qwen_classification.json"
    out_csv = OUT_DIR / f"{stem}_qwen_classification.csv"
    out_review = OUT_DIR / f"{stem}_CORRECT_ME.csv"
    originals = original_row_lookup(csv_path)

    out_json.write_text(json.dumps(data, indent=2), encoding="utf-8")

    fields = ["Index", "Component", "Role", "Confidence", "Quote", "Thickness", "Width", "Length", "CenterX", "CenterY", "CenterZ", "Reason"]
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for item in data.get("classifications", []):
            original = originals.get(str(item.get("index", "")), {})
            writer.writerow(
                {
                    "Index": item.get("index", ""),
                    "Component": original.get("Component", ""),
                    "Role": item.get("role", ""),
                    "Confidence": item.get("confidence", ""),
                    "Quote": item.get("quote", ""),
                    "Thickness": original.get("Thickness", ""),
                    "Width": original.get("Width", ""),
                    "Length": original.get("Length", ""),
                    "CenterX": original.get("CenterX", ""),
                    "CenterY": original.get("CenterY", ""),
                    "CenterZ": original.get("CenterZ", ""),
                    "Reason": item.get("reason", ""),
                }
            )

    review_fields = ["Index", "Component", "PredictedRole", "CorrectRole", "Confidence", "Thickness", "Width", "Length", "CenterX", "CenterY", "CenterZ", "Reason", "Notes"]
    with out_review.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=review_fields)
        writer.writeheader()
        for item in data.get("classifications", []):
            original = originals.get(str(item.get("index", "")), {})
            writer.writerow(
                {
                    "Index": item.get("index", ""),
                    "Component": original.get("Component", ""),
                    "PredictedRole": item.get("role", ""),
                    "CorrectRole": "",
                    "Confidence": item.get("confidence", ""),
                    "Thickness": original.get("Thickness", ""),
                    "Width": original.get("Width", ""),
                    "Length": original.get("Length", ""),
                    "CenterX": original.get("CenterX", ""),
                    "CenterY": original.get("CenterY", ""),
                    "CenterZ": original.get("CenterZ", ""),
                    "Reason": item.get("reason", ""),
                    "Notes": "",
                }
            )

    return out_json, out_csv, out_review


def _ask_qwen_json(prompt, model, timeout_minutes, label, fallback=None, schema=None):
    """Run one Qwen pass and parse its JSON, falling back rather than dying.

    A pass that cannot be parsed saves its raw output for inspection and returns
    ``fallback``, so a bad answer from one pass never costs the whole run.
    """
    print(f"[{label}] prompt {len(prompt):,} chars; sending to {model} ...", flush=True)
    started = time.time()
    try:
        if ollama_api_available():
            raw = run_ollama_api(prompt, model, timeout_minutes, schema=schema)
        else:
            print(f"[{label}] Ollama HTTP API not reachable; falling back to the CLI.", flush=True)
            raw = run_ollama(prompt, model, timeout_minutes)
    except Exception as exc:
        # A pass that cannot run must cost only that pass. Timeouts are the
        # common case -- qwen3.5:9b on CPU did not finish this prompt inside 40
        # minutes -- and letting one escape here killed the whole run and threw
        # away the mesh measurements with it. Everything downstream is built to
        # carry on from `fallback`.
        elapsed = time.time() - started
        print(
            f"[{label}] {type(exc).__name__} after {elapsed:.0f}s: {exc}. "
            f"Falling back and carrying on.",
            flush=True,
        )
        if isinstance(exc, (TimeoutError, socket.timeout)):
            print(
                f"[{label}] The model did not answer within {timeout_minutes} minutes. "
                f"Use --timeout-minutes 0 to wait indefinitely, or a smaller model.",
                flush=True,
            )
        return (fallback if fallback is not None else {"classifications": []}), False
    print(f"[{label}] {len(raw):,} chars back in {time.time() - started:.0f}s.", flush=True)
    try:
        return extract_json(raw), True
    except Exception:
        debug = OUT_DIR / f"last_bad_qwen_output_{label}.txt"
        debug.write_text(raw, encoding="utf-8", errors="replace")
        print(f"[{label}] could not parse JSON. Raw output saved to {debug}", flush=True)
        return (fallback if fallback is not None else {"classifications": []}), False


def run_stl_two_pass(
    csv_path,
    model_name,
    timeout_minutes,
    stl_dir=None,
    cell_in=None,
    verbose=True,
):
    """Name every part in three steps: candidates, measure, final call.

    1. Qwen reads the whole CAD dimension export -- sizes, stack positions, steel
       weight, stack order -- and proposes a candidate name per part, flagging
       which parts are worth measuring.
    2. The exported STL triangle meshes for those candidates are read and
       measured: true stack thickness, pockets per face, through-holes,
       counterbores, cross-drilling.
    3. Qwen sees its own candidate next to those measurements and gives the final
       name, expected to overrule itself wherever the mesh disagrees.

    The deterministic geometry rules still run first and are kept as the safety
    net: anything a pass leaves unclassified is patched from them, so the output
    always covers every part in the CSV.
    """
    csv_path = str(csv_path)
    kwargs = {}
    if cell_in:
        kwargs["cell_in"] = cell_in

    rows = _stl_naming.load_cad_rows(csv_path)
    stack = _stl_naming.build_stack_model(rows, job=Path(csv_path).resolve().parent.name)
    if verbose:
        print(
            f"Loaded {len(stack.parts)} parts; {len(stack.in_stack_parts())} look structural. "
            f"Mold footprint {stack.footprint_w:.3f} x {stack.footprint_l:.3f} in."
            + (
                f" Ejector box spans {stack.ejector_box[0]:.3f}..{stack.ejector_box[1]:.3f}."
                if stack.ejector_box
                else ""
            ),
            flush=True,
        )

    # Deterministic safety net, and the source of sequenced_latch_lock_base, which
    # the VBA bridge reads to mark secondary parting lines.
    rules = classify_geometry(read_rows(csv_path, include_names=True))

    # The model is asked about the structural parts only. The rules already
    # separate the 127 fasteners in a job like C17880 reliably, and spending the
    # generation on them crowds out the ten plates that are the actual problem.
    expect = {p.index for p in _stl_naming.structural_parts(stack)}

    # ---- pass 1: candidates from the CSV ---------------------------------
    p1_prompt = _stl_naming.build_candidate_prompt(stack, csv_path, ROLES, rules=rules)
    p1, ok1 = _ask_qwen_json(
        p1_prompt, model_name, timeout_minutes, "pass1", rules,
        schema=_classification_schema(ROLES),
    )
    p1, problems1 = _stl_naming.validate_classifications(p1, stack, ROLES, expect=expect)
    patched1 = _stl_naming.fill_missing(p1, rules, stack, "geometry rules, not named in pass 1")
    _stl_naming.apply_candidates(stack, p1)
    if verbose:
        for prob in problems1:
            print(f"  pass1 issue: {prob}", flush=True)
        if patched1:
            print(f"  pass1: {patched1} parts patched from the geometry rules.", flush=True)

    # ---- measure the candidates ------------------------------------------
    wanted = _stl_naming.candidates_to_measure(stack, p1)
    if verbose:
        print(f"Measuring STL geometry for {len(wanted)} candidate parts ...", flush=True)
    measure = _stl_naming.measure_candidates(
        stack, wanted, stl_dir=stl_dir, csv_path=csv_path, verbose=verbose, **kwargs
    )

    provenance = {
        "naming_method": "stl_two_pass",
        "pass1_parsed": ok1,
        "pass1_problems": problems1,
        "stl_dir": measure.get("stl_dir", ""),
        "stl_files": measure.get("files", 0),
        "stl_matched_to_cad": measure.get("matched", 0),
        "parts_measured_from_mesh": measure.get("measured", 0),
        "candidates_without_mesh": measure.get("wanted_without_mesh", []),
        "measure_notes": measure.get("notes", []),
    }

    if not measure.get("measured"):
        # No mesh to add, so a second pass would see exactly what the first saw.
        if verbose:
            print(
                "No candidate part had an exported STL, so the final pass was skipped "
                "and the first pass stands.",
                flush=True,
            )
        p1.setdefault("job_analysis", {}).update(provenance)
        p1["job_analysis"]["naming_method"] = "csv_only_no_stl"
        p1["job_analysis"].setdefault(
            "sequenced_latch_lock_base",
            rules.get("job_analysis", {}).get("sequenced_latch_lock_base", False),
        )
        return p1

    # ---- pass 2: final call from the mesh --------------------------------
    # The mesh can move a part in or out of the stack (a measured thickness changes
    # its plan area and its gaps), so the expected set is recomputed rather than
    # reused from pass 1.
    expect = {p.index for p in _stl_naming.structural_parts(stack)}
    p2_prompt = _stl_naming.build_final_prompt(
        stack, csv_path, ROLES, first_pass=p1, rules=rules
    )
    p2, ok2 = _ask_qwen_json(
        p2_prompt, model_name, timeout_minutes, "pass2", p1,
        schema=_classification_schema(ROLES),
    )
    p2, problems2 = _stl_naming.validate_classifications(p2, stack, ROLES, expect=expect)
    patched2 = _stl_naming.fill_missing(p2, p1, stack, "pass 1, not revised in pass 2")
    changes = _stl_naming.diff_passes(stack, p2)

    if verbose:
        for prob in problems2:
            print(f"  pass2 issue: {prob}", flush=True)
        if patched2:
            print(f"  pass2: {patched2} parts carried over from pass 1.", flush=True)
        if changes:
            print(f"\nThe mesh changed {len(changes)} names:", flush=True)
            for c in changes:
                flag = "measured" if c["measured"] else "cad only"
                print(
                    f"  index {c['index']:>3} {c['name'][:34]:34s} "
                    f"{c['from']} -> {c['to']}  ({flag})",
                    flush=True,
                )
        else:
            print("\nThe mesh confirmed every name from the first pass.", flush=True)

    provenance.update(
        {
            "pass2_parsed": ok2,
            "pass2_problems": problems2,
            "changed_by_mesh": changes,
        }
    )
    p2.setdefault("job_analysis", {}).update(provenance)
    p2["job_analysis"].setdefault(
        "sequenced_latch_lock_base",
        rules.get("job_analysis", {}).get("sequenced_latch_lock_base", False),
    )
    return p2


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_path", help="Path to XT_Export_CAD_Dimensions.csv")
    parser.add_argument("--model", default="qwen3.5:4b")
    parser.add_argument("--max-rows", type=int, default=140)
    parser.add_argument("--timeout-minutes", type=int, default=240, help="0 disables timeout")
    parser.add_argument("--include-names", action="store_true", help="include shortened component names in the prompt")
    parser.add_argument("--long-knowledge", action="store_true", help="include full knowledge files instead of compact rules")
    parser.add_argument("--rules-only", action="store_true", help="skip Qwen/Ollama and use deterministic geometry rules only")
    parser.add_argument(
        "--stl-two-pass",
        action="store_true",
        help="candidates from the CSV, then measure the STL meshes, then a final "
        "naming pass. On by default when the job has a stl\\ folder.",
    )
    parser.add_argument(
        "--no-stl",
        action="store_true",
        help="force the old single-pass CSV-only prompt even if STL meshes exist",
    )
    parser.add_argument("--stl-dir", default="", help="per-plate STL folder (default: auto-detect)")
    parser.add_argument(
        "--stl-cell-in",
        type=float,
        default=0.0,
        help="mesh rasterisation cell size in inches (default 0.05; smaller finds "
        "smaller holes and runs slower)",
    )
    args = parser.parse_args()

    rows = read_rows(args.csv_path, include_names=args.include_names)
    print(f"Loaded {len(rows)} rows from {args.csv_path}", flush=True)
    if len(rows) > args.max_rows:
        print(f"Limiting to first {args.max_rows} rows for this run.", flush=True)
        rows = rows[: args.max_rows]

    # Two-pass is the default whenever the job actually has meshes to read: it is
    # strictly more evidence than the CSV alone, and it degrades to the first pass
    # on its own when nothing was measurable.
    stl_dir = args.stl_dir or None
    auto_stl = (
        not args.rules_only
        and not args.no_stl
        and _stl_naming.find_stl_dir(args.csv_path, stl_dir) is not None
    )

    if args.rules_only:
        print("Rules-only mode: skipping Qwen/Ollama.", flush=True)
        data = classify_geometry(rows)
    elif args.stl_two_pass or auto_stl:
        data = run_stl_two_pass(
            args.csv_path,
            args.model,
            args.timeout_minutes,
            stl_dir=stl_dir,
            cell_in=args.stl_cell_in or None,
        )
    else:
        prompt = build_prompt(rows, args.csv_path, long_knowledge=args.long_knowledge)
        print(f"Prompt size: {len(prompt):,} characters", flush=True)
        print("Prompt built. Sending to Qwen now...", flush=True)
        raw = run_ollama(prompt, args.model, args.timeout_minutes)
        print("Qwen returned output. Parsing JSON...", flush=True)
        try:
            data = extract_json(raw)
        except Exception:
            debug = OUT_DIR / "last_bad_qwen_output.txt"
            debug.write_text(raw, encoding="utf-8", errors="replace")
            print(f"Could not parse Qwen JSON. Raw output saved to {debug}", flush=True)
            print("Using geometry-rule fallback so you still get output files.", flush=True)
            data = classify_geometry(rows)

    out_json, out_csv, out_review = write_outputs(data, args.csv_path)
    print(f"Wrote {out_json}")
    print(f"Wrote {out_csv}")
    print(f"Wrote correction file {out_review}")


if __name__ == "__main__":
    main()
