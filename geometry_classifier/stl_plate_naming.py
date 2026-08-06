"""Name mold plates in two passes: CSV candidates, then a final call off the mesh.

The flow
--------
    pass 1   XT_Export_CAD_Dimensions.csv (every part)
             sizes, stack positions, steel weight, stack order
                -> Qwen proposes a CANDIDATE name per part
                -> and marks which parts are steel plates worth measuring

    measure  for those candidates only, read the exported STL triangles and
             measure the real solid: true thickness, pockets per face, through
             holes, counterbores, cross-drilling, full-thickness fraction

    pass 2   the candidate name + the measured geometry
                -> Qwen makes the FINAL call, free to overrule pass 1

Splitting it this way matters for two reasons. The candidate pass is cheap and
narrows 137 parts down to the ten or so that are actually plates, so the
expensive triangle work only runs where it changes an answer. And the final pass
gets to see a *proposal* next to *measurements*, which is a much easier judgement
than naming from raw numbers -- it only has to agree or correct.

Why the CSV is not enough on its own
------------------------------------
``XT_Export_CAD_Dimensions.csv`` is the only complete part *list* we have, and its
positions are reliable. As shape evidence it is thin, and on the C178 jobs thin in
ways that produced wrong names:

* C17879 index 8 is a 0.250" sheet sitting **above** the top clamp plate with
  ``SolidFillPct=0`` and no holes or pockets recorded. With nothing to name it
  by, it became "Ejector Plate" on the steel sheet.
* C17879's four **real** ejector plates (indices 6, 7, 11, 12) are only
  15.625 x 16.313 -- about a third of the 23.75 x 31.5 footprint -- so the
  full-footprint test dropped them and they reached no sheet at all.
* C17880 index 5 came out as the generic "Plate 4".
* C17880 index 2 is recorded 1.878" thick, which is not a size steel comes in.
  The mesh measures 1.875 -- 1-7/8.

Also, the CSV sorts Thickness/Width/Length ascending, so a rail's real stack
height is thrown away and the smallest side takes its place. The mesh keeps it.

Why the mesh join is on volume
------------------------------
The per-plate STLs cannot be positioned: SolidWorks re-zeroes each exported body
near its own origin, so C17880's eight plates all report a stack centre between
0.4" and 5.6" regardless of where they really sit. What *is* preserved is
orientation (the stack axis is Y in every file) and shape.

Volume turns out to be an exact join key. SolidWorks' own
``BBoxVolume_cuin x SolidFillPct`` and this module's mesh integration agree to
four significant figures on all eight C17880 plates (549.4/549.68, 435.3/435.50,
266.5/266.74, ...), and no two plates in a base share a volume. Each body is
matched on volume and confirmed on dimensions.

What the model is not told
--------------------------
The macro writes its *current* guess at a plate's role into both the STL file
name and the mesh header ("Plate 4.STL", "Ejector Plate.STL"). Feeding that back
as evidence is what lets a wrong name survive a re-run, so it is withheld from
both prompts -- see ``stl_geometry._solid_name``.

The derived stack facts (order, gaps, ejector-box membership, mirror twins) are
computed here in plain Python rather than left to the model, because they are
geometry bookkeeping with one right answer. The model's job is the part that
needs judgement: turning them into shop names.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np

from . import stl_geometry


# Steel density, lb/in^3, for a real weight figure. The CSV's own Mass_or_Vol
# column is exactly volume / 27.680 on every job measured, so it carries nothing
# volume does not; a weight an estimator recognises is more use to the model.
STEEL_LB_PER_CUIN = 0.283

# A part shares the mold footprint when its plan area reaches this fraction of
# the largest plan area in the job.
FOOTPRINT_FAMILY_FRAC = 0.90

# Minimum plan-area fraction to count as a structural plate rather than hardware.
# Kept low on purpose: C17879's split ejector plates are only ~34% of the
# footprint, and a full-footprint-only test is exactly why they were dropped.
IN_STACK_MIN_AREA_FRAC = 0.25

# Volume agreement required to join an STL body to a CSV row.
VOLUME_MATCH_FRAC = 0.02
VOLUME_MATCH_ABS = 0.05

# Two dimensions this close are the same nominal size.
DIM_TOL_IN = 0.03

# Roles whose parts are worth the expensive mesh pass. Anything the candidate
# pass calls a fastener or a pin is named well enough already.
MEASURE_WORTHY_ROLES = {
    "top_clamp_plate", "bottom_clamp_plate", "a_plate", "b_plate",
    "manifold_plate", "stripper_plate", "sc_retainer_plate", "sc_backup_plate",
    "support_plate", "full_footprint_plate", "rail", "rail_1", "rail_2",
    "pin_plate", "ejector_plate", "bottom_ejector_plate",
    "ejector_retainer_plate", "ejector_backup_plate", "insert_or_core_detail",
}

# Roles a base can only have one of.
SINGULAR_ROLES = {
    "top_clamp_plate", "bottom_clamp_plate", "a_plate", "b_plate",
    "manifold_plate", "stripper_plate", "sc_retainer_plate", "sc_backup_plate",
}


def _f(row: dict, key: str, default: float = 0.0) -> float:
    try:
        return float(str(row.get(key, "")).strip())
    except (TypeError, ValueError):
        return default


def load_cad_rows(csv_path: str | Path) -> list[dict]:
    """Read ``XT_Export_CAD_Dimensions.csv`` into dicts, floats parsed."""
    import csv as _csv

    with Path(csv_path).open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        rows = list(_csv.DictReader(f))
    out = []
    for r in rows:
        out.append(
            {
                "index": str(r.get("Index", "")).strip(),
                "component": (r.get("Component", "") or "").strip(),
                "qty": int(_f(r, "Qty", 1) or 1),
                "thickness": _f(r, "Thickness"),
                "width": _f(r, "Width"),
                "length": _f(r, "Length"),
                "bbox_volume": _f(r, "BBoxVolume_cuin"),
                "solid_fill_pct": _f(r, "SolidFillPct"),
                "center_x": _f(r, "CenterX"),
                "center_y": _f(r, "CenterY"),
                "center_z": _f(r, "CenterZ"),
                "n_thru": int(_f(r, "NThruHoles")),
                "n_cbore": int(_f(r, "NCbore")),
                "n_cross": int(_f(r, "NCrossAxis")),
                "max_bore": _f(r, "MaxBoreDia"),
                "hole_sig": (r.get("HoleSig", "") or "").strip(),
                "n_pockets": int(_f(r, "NPockets")),
                "pocket_area": _f(r, "PocketAreaIn2"),
                "pocket_up": _f(r, "PocketAreaUpIn2"),
                "pocket_dn": _f(r, "PocketAreaDnIn2"),
                "raw": r,
            }
        )
    return out


def csv_volume(row: dict) -> float:
    """The CSV's own solid volume, reconstructed from bbox x fill."""
    return row["bbox_volume"] * row["solid_fill_pct"] / 100.0


# ─────────────────────────────────────────────────────────────────────────────
# The stack model (shared by both passes)
# ─────────────────────────────────────────────────────────────────────────────
def _half_token(component: str) -> str:
    """Which half of the base a component's assembly path puts it in.

    Hewitt's C17880 tree splits into ``bs-quote`` and ``bm-quote`` sub-assemblies
    -- stationary and moving. That single token pins the parting line down harder
    than any dimension does, so it is worth extracting when present.
    """
    # No trailing \b on "quote": the real tokens read "bs-quote_1-1", and an
    # underscore is a word character, so \b would never match there.
    c = component.lower()
    if re.search(r"(?:^|[^a-z0-9])bs[-_ ]?quote|stationary|cavity[-_ ]?half", c):
        return "stationary"
    if re.search(r"(?:^|[^a-z0-9])bm[-_ ]?quote|moving|core[-_ ]?half", c):
        return "moving"
    return ""


def _short_component(component: str) -> str:
    """The leaf of an assembly path, with the instance suffix trimmed."""
    leaf = component.replace("\\", "/").split("/")[-1].strip()
    return re.sub(r"-\d+$", "", leaf) or leaf


@dataclass
class StackPart:
    index: str
    component: str
    short_name: str
    half: str
    qty: int

    # From the CSV. Position is only ever taken from here.
    center_x: float
    center_y: float
    center_z: float
    csv_thickness: float
    csv_width: float
    csv_length: float
    csv_volume: float
    csv_fill_pct: float = 0.0
    csv_n_thru: int = 0
    csv_n_cbore: int = 0
    csv_n_cross: int = 0
    csv_max_bore: float = 0.0
    csv_hole_sig: str = ""
    csv_pocket_up: float = 0.0
    csv_pocket_dn: float = 0.0

    # Measured off the mesh in the measure step, when the part was a candidate.
    stl: stl_geometry.BodyGeometry | None = None
    stl_file: str = ""

    # Stack height worked out from the mold, for parts where the CSV's sorted
    # sides cannot say which one is vertical. Rails are the case that matters:
    # see _infer_rail_thickness.
    inferred_thickness: float = 0.0
    inferred_thickness_why: str = ""

    # Pass-1 proposal.
    candidate_role: str = ""
    candidate_confidence: str = ""
    candidate_reason: str = ""

    # Derived, computed here.
    plan_area: float = 0.0
    plan_area_frac: float = 0.0
    footprint_family: bool = False
    in_stack: bool = False
    order_from_top: int = 0
    gap_above: float | None = None
    gap_below: float | None = None
    inside_ejector_box: bool = False
    is_rail_like: bool = False
    mirror_twins: list[str] = field(default_factory=list)

    # Which CSV centre column is "up the stack" for this job. Set by
    # build_stack_model from detect_stack_axis; never assumed.
    stack_axis_field: str = "center_z"

    @property
    def stack_pos(self) -> float:
        """Position up the stack.

        NOT always CenterZ. C18184 is a 5-plate base stacked along **Y** with
        CenterZ 0.000 on every plate -- reading Z there put all five plates on
        one level and the stack model produced nothing usable.
        """
        return getattr(self, self.stack_axis_field, self.center_z)

    @property
    def plane_pos(self) -> tuple[float, float]:
        """The two in-plane centre coordinates, for mirror-twin detection."""
        return tuple(
            getattr(self, f)
            for f in ("center_x", "center_y", "center_z")
            if f != self.stack_axis_field
        )

    @property
    def thickness(self) -> float:
        """Stack height, best available. The mesh wins outright; failing that an
        inferred height beats the CSV, whose sorted sides lose which is vertical."""
        if self.stl is not None:
            return self.stl.thickness_in
        if self.inferred_thickness > 0:
            return self.inferred_thickness
        return self.csv_thickness

    @property
    def weight_lb(self) -> float | None:
        """Steel weight, or None when nothing has actually measured the solid.

        Returning 0.0 for a part whose feature walk failed would read as "weighs
        nothing" -- which is how C17879's 0.250" sheet looked to the old rules.
        """
        if self.stl is not None:
            return self.stl.volume_cuin * STEEL_LB_PER_CUIN
        if self.csv_volume > 0:
            return self.csv_volume * STEEL_LB_PER_CUIN
        return None

    @property
    def weight_lb_if_solid(self) -> float:
        """Weight as a solid block of its bounding box: an upper bound that is
        always available, so an unmeasured part still has a sense of scale."""
        dims = [self.csv_thickness, self.csv_width, self.csv_length]
        return dims[0] * dims[1] * dims[2] * STEEL_LB_PER_CUIN

    @property
    def stack_top(self) -> float:
        return self.stack_pos + self.thickness / 2.0

    @property
    def stack_bottom(self) -> float:
        return self.stack_pos - self.thickness / 2.0


@dataclass
class StackModel:
    job: str = ""
    parts: list[StackPart] = field(default_factory=list)
    # "center_x" | "center_y" | "center_z"; see detect_stack_axis_field.
    stack_axis_field: str = "center_z"
    max_plan_area: float = 0.0
    footprint_w: float = 0.0
    footprint_l: float = 0.0
    ejector_box: tuple[float, float] | None = None
    stl_measured: int = 0
    notes: list[str] = field(default_factory=list)

    def by_index(self) -> dict[str, StackPart]:
        return {p.index: p for p in self.parts}

    def in_stack_parts(self) -> list[StackPart]:
        return sorted((p for p in self.parts if p.in_stack), key=lambda p: -p.stack_pos)


def build_stack_model(rows: list[dict], job: str = "") -> StackModel:
    """Turn CSV rows into an ordered stack model with derived facts."""
    m = StackModel(job=job)
    for r in rows:
        m.parts.append(
            StackPart(
                index=r["index"],
                component=r["component"],
                short_name=_short_component(r["component"]),
                half=_half_token(r["component"]),
                qty=r["qty"],
                center_x=r["center_x"],
                center_y=r["center_y"],
                center_z=r["center_z"],
                csv_thickness=r["thickness"],
                csv_width=r["width"],
                csv_length=r["length"],
                csv_volume=csv_volume(r),
                csv_fill_pct=r["solid_fill_pct"],
                csv_n_thru=r["n_thru"],
                csv_n_cbore=r["n_cbore"],
                csv_n_cross=r["n_cross"],
                csv_max_bore=r["max_bore"],
                csv_hole_sig=r["hole_sig"],
                csv_pocket_up=r["pocket_up"],
                csv_pocket_dn=r["pocket_dn"],
            )
        )
    refresh_derived(m)
    return m


def detect_stack_axis_field(parts: list[StackPart]) -> str:
    """Which CSV centre column runs up the stack. Measured, never assumed.

    Plates share the mold footprint, so in the two in-plane axes every plate sits
    at essentially the same centre; only along the stack axis do their centres
    spread out. The axis with the widest spread of plate centres is the stack.

    C18184 is why this exists: a five-plate base stacked along **Y**, with
    ``CenterZ`` 0.000 on all five. Assuming Z there collapsed the whole stack onto
    one level. The deterministic rules in qwen_classify_xt_csv already detect this
    (they reported "CenterY" for that job); this brings the STL path in line.

    Preference is measured among the biggest plates when there are enough of
    them, because hardware scattered around the mold spreads on every axis and
    would blur the signal.
    """
    if not parts:
        return "center_z"
    biggest = max((p.plan_area for p in parts), default=0.0)
    plates = [p for p in parts if biggest > 0 and p.plan_area >= 0.9 * biggest]
    sample = plates if len(plates) >= 2 else parts
    spreads = {}
    for f in ("center_x", "center_y", "center_z"):
        vals = [getattr(p, f) for p in sample]
        spreads[f] = max(vals) - min(vals)
    best = max(spreads, key=lambda f: spreads[f])
    # A stack with no spread at all (one plate) tells us nothing; keep the
    # conventional Z rather than picking an axis on floating-point noise.
    return best if spreads[best] > 1e-6 else "center_z"


def _stack_levels(parts: list[StackPart], tol: float = 0.02) -> list[list[StackPart]]:
    """Group parts into stack levels, highest first.

    Parts sitting side by side at one height -- a rail pair, a split ejector set --
    are one level. Treating them as separate stack entries is what produced
    negative "gaps" between C17879's twinned ejector plates.
    """
    levels: list[list[StackPart]] = []
    for p in sorted(parts, key=lambda x: -x.stack_pos):
        if levels and abs(levels[-1][0].stack_pos - p.stack_pos) <= tol:
            levels[-1].append(p)
        else:
            levels.append([p])
    return levels


def _infer_rail_thickness(m: StackModel) -> None:
    """Give rail-shaped parts their real stack height.

    The CSV sorts a part's three sides ascending and calls the smallest
    "Thickness", which for a rail is its *width*, not its height in the mold.
    C17879's rails come through as 1.875 x 3.000 x 31.500 and the shop's own steel
    sheet orders them "3 x 1.875 x 31.5" -- 3" is the stack height.

    Rather than guess which side is vertical, take it from the mold: the ejector
    box is a known opening between two full plates, and a rail is what holds that
    opening apart, so a rail side matching the box height IS the height.
    """
    if not m.ejector_box:
        return
    box_h = m.ejector_box[1] - m.ejector_box[0]
    if box_h <= 0.25:
        return
    for p in m.parts:
        if p.stl is not None or not p.is_rail_like:
            continue
        for side in (p.csv_thickness, p.csv_width, p.csv_length):
            if abs(side - box_h) <= max(DIM_TOL_IN, 0.02 * box_h):
                if abs(side - p.csv_thickness) > 1e-9:
                    p.inferred_thickness = side
                    p.inferred_thickness_why = (
                        f"stack height {side:.3f} taken from the ejector-box opening; "
                        f"the CAD export sorted its sides and listed "
                        f"{p.csv_thickness:.3f} as the thickness"
                    )
                break


def refresh_derived(m: StackModel) -> None:
    """Recompute every derived fact.

    Run again after the mesh pass: a measured thickness moves stack tops, gaps and
    the ejector box with it. Order matters here, because the ejector box is needed
    to infer rail heights and rail heights change the gaps.
    """
    for p in m.parts:
        if p.stl is not None:
            p.plan_area = p.stl.plan_short_in * p.stl.plan_long_in
        else:
            dims = sorted([p.csv_thickness, p.csv_width, p.csv_length])
            p.plan_area = dims[1] * dims[2]
    m.max_plan_area = max((p.plan_area for p in m.parts), default=0.0)

    # Settle the stack axis before anything reads a position off a part. Plan
    # areas are needed first (the detector prefers the big plates), and every
    # ordering, gap and ejector-box test below goes through StackPart.stack_pos.
    m.stack_axis_field = detect_stack_axis_field(m.parts)
    for p in m.parts:
        p.stack_axis_field = m.stack_axis_field

    biggest = max(m.parts, key=lambda p: p.plan_area, default=None)
    if biggest is not None:
        if biggest.stl is not None:
            m.footprint_w, m.footprint_l = biggest.stl.plan_short_in, biggest.stl.plan_long_in
        else:
            dims = sorted([biggest.csv_thickness, biggest.csv_width, biggest.csv_length])
            m.footprint_w, m.footprint_l = dims[1], dims[2]

    for p in m.parts:
        p.plan_area_frac = p.plan_area / m.max_plan_area if m.max_plan_area else 0.0
        p.footprint_family = p.plan_area_frac >= FOOTPRINT_FAMILY_FRAC
        p.in_stack = (
            p.plan_area_frac >= IN_STACK_MIN_AREA_FRAC
            and p.plan_area > 0
            and max(p.csv_width, p.csv_length) > 2.5 * min(
                p.csv_thickness, p.csv_width, p.csv_length
            )
        )

    # Mirror twins first: rail detection leans on them, and they are pure position
    # bookkeeping so nothing else has to be settled yet.
    for p in m.parts:
        p.mirror_twins = []
    for i, a in enumerate(m.parts):
        for b in m.parts[i + 1 :]:
            if abs(a.stack_pos - b.stack_pos) > 0.02:
                continue
            if (
                abs(a.csv_thickness - b.csv_thickness) > DIM_TOL_IN
                or abs(a.csv_width - b.csv_width) > DIM_TOL_IN
                or abs(a.csv_length - b.csv_length) > DIM_TOL_IN
            ):
                continue
            # Mirrored across a centreline in one of the two IN-PLANE axes,
            # which depend on where the stack runs -- not always X and Y.
            ap, bp = a.plane_pos, b.plane_pos
            if any(
                abs(u + v) < 0.05 or abs(u - v) < 0.05 for u, v in zip(ap, bp)
            ):
                a.mirror_twins.append(b.index)
                b.mirror_twins.append(a.index)

    # The ejector box is the biggest clear opening between consecutive
    # full-footprint plates -- the support plate above it and the bottom clamp
    # plate below. Deriving it from the plates rather than from the rails breaks a
    # circular dependency: a rail's own height is exactly what the CSV cannot say,
    # so rails cannot be what locates the box.
    m.ejector_box = None
    for p in m.parts:
        p.inside_ejector_box = False
    full_levels = _stack_levels([p for p in m.parts if p.footprint_family])
    best_gap = 0.0
    for upper, lower in zip(full_levels, full_levels[1:]):
        top_of_lower = max(q.stack_pos + q.thickness / 2.0 for q in lower)
        bot_of_upper = min(q.stack_pos - q.thickness / 2.0 for q in upper)
        gap = bot_of_upper - top_of_lower
        if gap > best_gap and gap > 0.5:
            best_gap = gap
            m.ejector_box = (round(top_of_lower, 4), round(bot_of_upper, 4))

    # A rail is long, narrow, well short of the full footprint, and stands inside
    # the ejector box. Its own thickness is deliberately not part of the test.
    for p in m.parts:
        long_side = max(p.csv_width, p.csv_length)
        short_side = min(p.csv_width, p.csv_length)
        in_box = bool(
            m.ejector_box and m.ejector_box[0] - 0.05 <= p.stack_pos <= m.ejector_box[1] + 0.05
        )
        at_edge = bool(p.mirror_twins)
        p.is_rail_like = (
            long_side > 4.0 * max(short_side, 1e-6)
            and p.plan_area_frac < FOOTPRINT_FAMILY_FRAC
            and p.plan_area_frac > 0.10
            and in_box
            and at_edge
            and long_side > 0.6 * max(m.footprint_w, m.footprint_l)
        )
    _infer_rail_thickness(m)

    # Rails are quoted steel and belong in the stack listing. They can fall below
    # the plan-area threshold on a narrow mold -- C17879's are 12.6% of the
    # footprint -- so they are added on the strength of being rails, not their area.
    for p in m.parts:
        if p.is_rail_like:
            p.in_stack = True

    if m.ejector_box:
        lo, hi = m.ejector_box
        for p in m.parts:
            if not p.is_rail_like and lo - 0.02 <= p.stack_pos <= hi + 0.02:
                p.inside_ejector_box = True

    # Stack order and gaps, by level so side-by-side parts share a position.
    for p in m.parts:
        p.order_from_top = 0
        p.gap_above = p.gap_below = None
    # Rails are left out of the level sequence. They stand *alongside* the ejector
    # plates rather than stacking with them, so a gap measured against a rail is
    # meaningless -- it came out negative, which reads as an interference.
    levels = _stack_levels([p for p in m.parts if p.in_stack and not p.is_rail_like])

    def _clean_gap(gap: float) -> float:
        # Touching plates come out a few thousandths either side of zero from
        # tessellation and rounding. Reporting "-0.004" invites the model to read
        # an interference that is not there.
        return 0.0 if abs(gap) < 0.01 else round(gap, 4)

    for i, level in enumerate(levels):
        top_i = max(q.stack_pos + q.thickness / 2.0 for q in level)
        bot_i = min(q.stack_pos - q.thickness / 2.0 for q in level)
        for p in level:
            p.order_from_top = i + 1
            if i > 0:
                prev_bot = min(q.stack_pos - q.thickness / 2.0 for q in levels[i - 1])
                p.gap_above = _clean_gap(prev_bot - top_i)
            if i + 1 < len(levels):
                next_top = max(q.stack_pos + q.thickness / 2.0 for q in levels[i + 1])
                p.gap_below = _clean_gap(bot_i - next_top)

    # Rails still need a position in the listing, so order them by where their
    # centre falls among the levels. They keep no gaps.
    for p in m.parts:
        if not p.is_rail_like or not p.in_stack:
            continue
        p.order_from_top = 1 + sum(
            1 for lvl in levels if lvl[0].stack_pos > p.stack_pos
        )


# ─────────────────────────────────────────────────────────────────────────────
# PASS 1 -- candidates from the CSV
# ─────────────────────────────────────────────────────────────────────────────
CANDIDATE_RULES = """
WHAT YOU HAVE AND WHAT IT IS WORTH

This is the SolidWorks dimension export for every component in one mold base.
Positions and sizes are reliable. The feature counts are not: the extractor
sometimes gives up and returns fill_pct 0 with no holes and no pockets. A part
like that has NOT been measured. Do not read it as a featureless plate, and do
not name it confidently from nothing -- mark it low confidence and say the
geometry is missing.

thickness_in here is the extent along the STACK axis, not the smallest side of
the box. A rail is thin in plan and TALL in the stack.

  order_from_top       1 = highest part in the stack, counting down
  gap_above/gap_below  clear space to the neighbouring plate, inches. 0 = touching
  plan_area_frac       plan area / biggest plan area in the job. 1.0 = full mold
                       footprint. Ejector plates are often 0.3-0.7 and are still
                       ejector plates -- do NOT require a full footprint
  inside_ejector_box   sits between the rails, inside the ejector housing
  mirror_twins         same-size parts at the same height mirrored across a
                       centreline: a split or paired set
  half                 stationary = cavity side, moving = core/ejector side.
                       When present this fixes the parting line
  weight_lb            steel weight from volume; a plate weighs hundreds of
                       pounds, a fastener ounces

A STANDARD BASE, TOP TO BOTTOM

  top clamp plate      top of the stationary half, bolts to the platen
  A plate              cavity plate, stationary half
  --- parting line --- between the lowest stationary and highest moving part
  B plate              core plate, moving half
  support plate        full plate below the B plate
  rails                two tall narrow blocks enclosing the ejector box
  ejector retainer     inside the box, thinner, many counterbores for pin heads
  ejector back-up      inside the box, thicker, under the retainer
  bottom clamp plate   bottom of the moving half, often LONGER than the rest
                       because it carries clamp slots

Not every base has every plate and some have extras (a second support plate, a
manifold plate on a hot-runner base, a stripper plate).
"""


def _candidate_payload(p: StackPart) -> dict:
    d: dict = {
        "index": p.index,
        "name": p.short_name,
        "qty": p.qty,
        "thickness_in": round(p.csv_thickness, 4),
        "plan_in": [round(p.csv_width, 3), round(p.csv_length, 3)],
        "stack_center": round(p.stack_pos, 4),
        "pos_x": round(p.center_x, 3),
        "pos_y": round(p.center_y, 3),
        "plan_area_frac": round(p.plan_area_frac, 3),
        "in_stack": p.in_stack,
    }
    w = p.weight_lb
    if w is None:
        d["weight_lb"] = None
        d["weight_lb_if_solid"] = round(p.weight_lb_if_solid, 1)
    else:
        d["weight_lb"] = round(w, 1)
    if p.inferred_thickness:
        d["thickness_in"] = round(p.inferred_thickness, 4)
        d["thickness_note"] = p.inferred_thickness_why
    if p.order_from_top:
        d["order_from_top"] = p.order_from_top
    if p.half:
        d["half"] = p.half
    if p.gap_above is not None:
        d["gap_above"] = p.gap_above
    if p.gap_below is not None:
        d["gap_below"] = p.gap_below
    if p.inside_ejector_box:
        d["inside_ejector_box"] = True
    if p.is_rail_like:
        d["rail_shaped"] = True
    if p.mirror_twins:
        d["mirror_twins"] = p.mirror_twins
    # Only report the feature walk's numbers when it actually produced some, so
    # the prompt never presents a failed extraction as a measurement of zero.
    if p.csv_fill_pct > 0:
        d["fill_pct"] = round(p.csv_fill_pct, 1)
        d["thru_holes"] = p.csv_n_thru
        d["cbores"] = p.csv_n_cbore
        d["cross_holes"] = p.csv_n_cross
        d["max_bore"] = round(p.csv_max_bore, 3)
        if p.csv_hole_sig:
            d["hole_sig"] = p.csv_hole_sig
        if p.csv_pocket_up or p.csv_pocket_dn:
            d["pocket_up_in2"] = round(p.csv_pocket_up, 2)
            d["pocket_dn_in2"] = round(p.csv_pocket_dn, 2)
    else:
        d["geometry_missing"] = True
    return d


def structural_parts(model: StackModel) -> list[StackPart]:
    """The parts the model is asked to name, highest in the stack first."""
    return sorted(
        (p for p in model.parts if p.in_stack or p.is_rail_like),
        key=lambda p: -p.stack_pos,
    )


def hardware_parts(model: StackModel) -> list[StackPart]:
    return [p for p in model.parts if not (p.in_stack or p.is_rail_like)]


def _hardware_payload(p: StackPart, rules_role: str) -> dict:
    return {
        "index": p.index,
        "name": p.short_name,
        "qty": p.qty,
        "dims_in": [
            round(p.csv_thickness, 3),
            round(p.csv_width, 3),
            round(p.csv_length, 3),
        ],
        "stack_center": round(p.stack_pos, 3),
        "pos_x": round(p.center_x, 3),
        "pos_y": round(p.center_y, 3),
        "rules_role": rules_role or "hardware_other",
        "twins": len(p.mirror_twins) or None,
    }


def _hardware_block(model: StackModel, rules: dict | None) -> tuple[str, dict[str, str]]:
    """The hardware list, already named by the geometry rules, plus that mapping.

    Hardware is not re-classified by the model. There are 127 fasteners in C17880
    against 10 plates, so asking for all 137 spends nearly all of the generation
    on the parts that were never the problem -- and the rules already separate
    pins, bushings and pillars reliably by diameter and length. The list is still
    shown, so a plate hiding among the hardware can be pulled back out via
    "hardware_corrections".
    """
    rules_by_index = {
        str(c.get("index", "")): str(c.get("role", ""))
        for c in (rules or {}).get("classifications", [])
    }
    hw = hardware_parts(model)
    payload = [_hardware_payload(p, rules_by_index.get(p.index, "")) for p in hw]
    return json.dumps(payload, separators=(",", ":")), rules_by_index


HARDWARE_INSTRUCTION = """
The hardware list below has ALREADY been named by the geometry rules, which are
reliable for fasteners, pins, bushings and pillars. Do NOT re-classify it and do
NOT copy it into "classifications".

Only use "hardware_corrections" for a part in that list you believe is actually a
structural plate or is plainly misnamed. Leave the array empty if none are.
"""


def build_candidate_prompt(
    model: StackModel,
    csv_path: str | Path,
    roles: list[str],
    rules: dict | None = None,
) -> str:
    """Pass 1: propose a candidate name per structural part from the CSV alone."""
    struct = structural_parts(model)
    hardware_json, _ = _hardware_block(model, rules)
    job_facts = {
        "job": model.job,
        "mold_footprint_in": [round(model.footprint_w, 3), round(model.footprint_l, 3)],
        "parts_total": len(model.parts),
        "structural_parts": len(struct),
        "hardware_parts": len(model.parts) - len(struct),
        "ejector_box_stack_span": list(model.ejector_box) if model.ejector_box else None,
        "halves_present": sorted({p.half for p in model.parts if p.half}) or None,
    }
    return f"""/no_think
You are a CMS mold-base geometry interpreter, doing a FIRST PASS.

Propose a candidate name for each of the {len(struct)} structural parts below from
its size, position, stack order and weight. A later pass will measure the real
solid geometry and can correct you, so give your best reading now and be honest
about confidence.

Also set "measure": true for any part whose name would change if we measured its
pockets, holes and true thickness.

Return JSON only. No explanation outside JSON. No markdown.

Allowed roles:
{", ".join(roles)}

{CANDIDATE_RULES}
{HARDWARE_INSTRUCTION}

Mold-level facts:
{json.dumps(job_facts, indent=2)}

Structural parts, highest in the stack first -- name every one of these:
{json.dumps([_candidate_payload(p) for p in struct], separators=(",", ":"))}

Hardware, already named by the geometry rules -- do not re-classify:
{hardware_json}

Return ONLY valid JSON in this exact shape:
{{
  "job_analysis": {{
    "stack_axis": "CenterZ",
    "parting_line": "which two indices the parting line runs between, and why",
    "rules_for_this_job": ["rule 1", "rule 2"]
  }},
  "classifications": [
    {{
      "index": "1",
      "role": "a_plate",
      "confidence": "HIGH|MEDIUM|LOW",
      "reason": "what decided it",
      "quote": true,
      "measure": true
    }}
  ],
  "hardware_corrections": []
}}

Every one of the {len(struct)} structural indices must appear exactly once in
"classifications".
CSV source: {csv_path}
"""


def apply_candidates(model: StackModel, data: dict) -> int:
    """Record pass 1's proposal on the stack model. Returns how many landed."""
    by_index = model.by_index()
    n = 0
    for item in data.get("classifications", []):
        p = by_index.get(str(item.get("index", "")).strip())
        if p is None:
            continue
        p.candidate_role = str(item.get("role", "")).strip().lower()
        p.candidate_confidence = str(item.get("confidence", "MEDIUM")).strip().upper()
        p.candidate_reason = str(item.get("reason", "")).strip()
        n += 1
    return n


def candidates_to_measure(model: StackModel, data: dict | None = None) -> set[str]:
    """Which part indices earn the expensive mesh pass.

    The model's own ``measure`` flag is honoured, then widened: any part whose
    candidate role is structural, anything in the stack, anything the CSV failed
    to measure, and anything the model was unsure about. Widening is deliberate --
    measuring a part we did not need costs seconds, while skipping one that
    mattered puts a wrong name on a steel order.
    """
    by_index = model.by_index()
    want: set[str] = set()
    if data:
        for item in data.get("classifications", []):
            if item.get("measure"):
                want.add(str(item.get("index", "")).strip())

    for p in model.parts:
        # Structural on the face of it: always measure.
        if p.in_stack or p.is_rail_like or p.candidate_role in MEASURE_WORTHY_ROLES:
            want.add(p.index)
            continue
        # Otherwise widen only to parts big enough in plan to be a plate at all.
        # Widening on a failed CAD extraction alone would pull in every fastener in
        # the job -- 129 of C17880's 137 parts -- and each would then be reported
        # as "a candidate with no mesh", which is noise, not a finding.
        if p.plan_area_frac < 0.15:
            continue
        if p.candidate_confidence in {"LOW", "MEDIUM"} or p.csv_fill_pct <= 0:
            want.add(p.index)

    # A "measure" flag on something with no plausible mesh is dropped for the same
    # reason -- the model does not know which parts were exported.
    return {
        i
        for i in want
        if i
        and i in by_index
        and (by_index[i].in_stack or by_index[i].is_rail_like or by_index[i].plan_area_frac >= 0.15)
    }


# ─────────────────────────────────────────────────────────────────────────────
# MEASURE -- read the triangles for the candidates
# ─────────────────────────────────────────────────────────────────────────────
PLATE_MANIFEST = ".plate_stls.json"


def per_plate_stl_files(folder: Path) -> list[Path]:
    """The one-part-per-file STLs in a folder, excluding whole-assembly exports.

    A job's ``models\\`` folder holds the per-plate meshes *and* two multi-body
    exports -- the merged ``<job>.stl`` and ``<job> component.stl``. Reading those
    as if they were single parts wastes a 10 MB integration on a body whose volume
    can never match a CAD row, and it skews the shared stack-axis detection.

    The web app already writes a manifest naming exactly the per-plate files, so
    use it when present and fall back to filtering by name.
    """
    manifest = folder / PLATE_MANIFEST
    if manifest.is_file():
        try:
            names = json.loads(manifest.read_text(encoding="utf-8"))
            picked = [folder / n for n in names if (folder / n).is_file()]
            if picked:
                return sorted(picked)
        except (OSError, ValueError, TypeError):
            pass  # fall through to the name filter

    out = []
    for p in sorted(folder.glob("*.[sS][tT][lL]")):
        stem = p.stem.lower()
        # The per-plate writer always joins the job name to the plate name with an
        # underscore; the two assembly exports have no underscore at all.
        if stem.endswith(" component") or "_" not in p.stem:
            continue
        out.append(p)
    return out


def find_stl_dir(csv_path: str | Path, explicit: str | Path | None = None) -> Path | None:
    """Locate the per-plate STL folder for a job.

    Looks where the macro actually writes them: a ``stl\\`` folder beside the CSV,
    in the job's staging folder, or the web app's ``models\\`` copy.
    """
    if explicit:
        p = Path(explicit)
        return p if p.is_dir() else None
    base = Path(csv_path).resolve().parent
    for cand in (base / "stl", base / "models", base.parent / "stl"):
        if cand.is_dir() and any(cand.glob("*.[sS][tT][lL]")):
            return cand
    return None


def measure_candidates(
    model: StackModel,
    wanted: set[str],
    stl_dir: str | Path | None = None,
    csv_path: str | Path | None = None,
    cell_in: float = stl_geometry.DEFAULT_CELL_IN,
    verbose: bool = True,
) -> dict:
    """Read the STL triangles for the candidate parts and measure them.

    Two steps, because the join key is a measurement. Every file is first read and
    cheaply integrated for volume and bounding box, which is what matches it to a
    CAD row; only bodies matching a wanted index then get the full feature pass.

    Returns a report dict for the prompt and the log.
    """
    report = {
        "stl_dir": "",
        "files": 0,
        "matched": 0,
        "measured": 0,
        "unmatched_files": [],
        "wanted_without_mesh": [],
        "notes": [],
    }
    folder = find_stl_dir(csv_path or ".", stl_dir)
    if folder is None:
        report["notes"].append(
            "No per-plate STL folder for this job; the final pass sees CAD numbers only."
        )
        if verbose:
            print(report["notes"][-1], flush=True)
        return report
    report["stl_dir"] = str(folder)

    files = per_plate_stl_files(folder)
    report["files"] = len(files)
    if not files:
        report["notes"].append(f"{folder} holds no per-plate STL files.")
        return report

    if verbose:
        print(f"Reading {len(files)} STL meshes from {folder} ...", flush=True)

    raw: list[tuple[Path, np.ndarray]] = []
    for p in files:
        try:
            tris = stl_geometry.read_stl(p)
        except Exception as exc:
            report["notes"].append(f"{p.name}: unreadable ({exc})")
            continue
        if len(tris):
            raw.append((p, tris))
    if not raw:
        report["notes"].append("Every STL in the folder was empty or unreadable.")
        return report

    # Units are settled by the join itself rather than by a magnitude guess.
    #
    # An STL records no units and SolidWorks writes whatever the document was set
    # to -- millimetres, on this shop's machine. The web app's stlGeometry.ts hit
    # the same thing (it once read a clamping plate as 1.7 million lb) and solved
    # it the right way: scale against a dimension the CAD export already knows in
    # inches. Here that hint is free, because matching a mesh to its CAD row is
    # something we have to do anyway -- so both readings are tried and the one
    # that actually matches a row is the correct one. No heuristic involved.
    scales = [("in", 1.0), ("mm", 1.0 / stl_geometry.MM_PER_IN)]

    stack_axis = stl_geometry.detect_stack_axis(
        np.array([np.ptp(t.reshape(-1, 3), axis=0) for _, t in raw])
    )

    # Cheap pass: volume + box for every file under both readings.
    quick: list[tuple[Path, np.ndarray, str, dict]] = []
    variants: list[list[tuple[str, float, dict]]] = []
    for p, tris in raw:
        vs = []
        for unit, sc in scales:
            vs.append((unit, sc, stl_geometry.quick_measure(tris * sc, stack_axis)))
        variants.append(vs)
        # Provisional; replaced by whichever reading the join confirms.
        unit_guess, scale_guess = stl_geometry.detect_units(tris)
        quick.append((p, tris, unit_guess, {"scale": scale_guess}))

    # Join on volume, greedy best-first so the tightest match claims its row.
    candidates: list[tuple[float, int, str, str, float]] = []
    for fi, vs in enumerate(variants):
        for unit, sc, q in vs:
            q_dims = sorted([q["thickness_in"], q["plan_short_in"], q["plan_long_in"]])
            for part in model.parts:
                rv = part.csv_volume
                if rv <= 0 or q["volume_cuin"] <= 0:
                    continue
                err = abs(q["volume_cuin"] - rv) / max(rv, 1e-9)
                if err > VOLUME_MATCH_FRAC and abs(q["volume_cuin"] - rv) > VOLUME_MATCH_ABS:
                    continue
                r_dims = sorted([part.csv_thickness, part.csv_width, part.csv_length])
                # The CSV sorts its three sides, so compare sorted against sorted.
                # The mesh's unsorted axes are used later, where they carry the
                # real thickness the CSV threw away.
                if any(abs(a - c) > DIM_TOL_IN for a, c in zip(q_dims, r_dims)):
                    continue
                candidates.append((err, fi, part.index, unit, sc))

    candidates.sort(key=lambda c: c[0])
    by_index = model.by_index()
    used_files: set[int] = set()
    used_rows: set[str] = set()
    pairs: list[tuple[int, str]] = []
    chosen: dict[int, tuple[str, float]] = {}
    for _err, fi, idx, unit, sc in candidates:
        if fi in used_files or idx in used_rows:
            continue
        used_files.add(fi)
        used_rows.add(idx)
        pairs.append((fi, idx))
        chosen[fi] = (unit, sc)
    if chosen:
        units_seen = {u for u, _ in chosen.values()}
        if len(units_seen) > 1:
            report["notes"].append(
                f"meshes in this folder are not all in the same units ({', '.join(sorted(units_seen))}); "
                f"each was scaled to whatever matched its own CAD row"
            )

    # Second join, on dimensions alone, for rows the volume join could not reach.
    #
    # A row whose feature walk failed has SolidFillPct 0, so its CSV volume is 0
    # and no volume can ever match it -- yet those are exactly the parts that most
    # need measuring. C17879 index 8 is the case: the 0.250" sheet that became
    # "Ejector Plate" precisely because nothing had measured it. Its mesh was
    # sitting right there in the folder, unmatched.
    #
    # Dimensions alone are a weaker key, so this only runs for rows with no usable
    # volume, and only when exactly one unmatched mesh fits the size.
    for part in model.parts:
        if part.index in used_rows or part.csv_volume > 0:
            continue
        r_dims = sorted([part.csv_thickness, part.csv_width, part.csv_length])
        fits = []
        for fi, vs in enumerate(variants):
            if fi in used_files:
                continue
            for unit, sc, q in vs:
                q_dims = sorted([q["thickness_in"], q["plan_short_in"], q["plan_long_in"]])
                if all(abs(a - c) <= DIM_TOL_IN for a, c in zip(q_dims, r_dims)):
                    fits.append((fi, unit, sc))
                    break
        if len(fits) == 1:
            fi, unit, sc = fits[0]
            used_files.add(fi)
            used_rows.add(part.index)
            pairs.append((fi, part.index))
            chosen[fi] = (unit, sc)
            report["notes"].append(
                f"index {part.index} matched {quick[fi][0].name} on dimensions; "
                f"its CAD row has no usable volume (the feature walk returned fill 0)"
            )
        elif len(fits) > 1:
            report["notes"].append(
                f"index {part.index} has no usable CAD volume and {len(fits)} meshes "
                f"share its size, so it was left unmatched rather than guessed"
            )

    report["matched"] = len(pairs)

    # Expensive pass: only the wanted parts.
    for fi, idx in pairs:
        if wanted and idx not in wanted:
            continue
        path, tris, _guess, _q = quick[fi]
        unit, scale = chosen.get(fi, (_guess, _q.get("scale", 1.0)))
        part = by_index[idx]
        try:
            part.stl = stl_geometry.analyze_body(
                tris * scale if scale != 1.0 else tris,
                stack_axis=stack_axis,
                cell_in=cell_in,
                source=path.name,
                body_index=fi,
                label_hint="",  # withheld on purpose: it is the macro's old guess
                units=unit,
            )
        except Exception as exc:
            report["notes"].append(f"{path.name}: measurement failed ({exc})")
            continue
        part.stl_file = path.name
        report["measured"] += 1
        if verbose:
            print(
                f"  measured index {idx}: {part.stl.thickness_in:.3f} x "
                f"{part.stl.plan_short_in:.3f} x {part.stl.plan_long_in:.3f} in, "
                f"{part.stl.n_triangles:,} triangles",
                flush=True,
            )

    report["unmatched_files"] = [
        quick[i][0].name for i in range(len(quick)) if i not in used_files
    ]
    report["wanted_without_mesh"] = sorted(
        i for i in wanted if by_index.get(i) and by_index[i].stl is None
    )
    model.stl_measured = report["measured"]

    if report["unmatched_files"]:
        report["notes"].append(
            f"{len(report['unmatched_files'])} STL files matched no CAD row: "
            + ", ".join(report["unmatched_files"][:6])
        )
    if report["wanted_without_mesh"]:
        report["notes"].append(
            f"{len(report['wanted_without_mesh'])} candidate parts have no exported STL "
            f"(indices {', '.join(report['wanted_without_mesh'][:12])}); these are named "
            f"from CAD numbers only. The macro only exports an STL for a plate it "
            f"already named, so an unnamed plate has no mesh to read."
        )
    refresh_derived(model)
    if verbose:
        for n in report["notes"]:
            print(f"  note: {n}", flush=True)
    return report


# ─────────────────────────────────────────────────────────────────────────────
# PASS 2 -- final call from the measured geometry
# ─────────────────────────────────────────────────────────────────────────────
FINAL_RULES = """
WHAT CHANGED SINCE THE FIRST PASS

Parts marked "measured": true below were read straight off the exported STL
solid -- the triangle mesh itself, integrated for volume and rasterised for
pockets and holes. Those numbers are measurements, not estimates, and they
override anything the first pass assumed. In particular:

  thickness_in       the real stack height. Where this disagrees with the first
                     pass, the mesh is right. A CAD export listing 1.878" for a
                     plate the mesh measures at 1.875" is rounding noise in the
                     export; steel comes in 1-7/8
  pocket_top         [area in2, max depth] on the UP-facing face
  pocket_bottom      [area in2, max depth] on the DOWN-facing face
                     Which side a pocket opens on is the strongest single cue
                     you have: a cavity plate is pocketed toward the parting
                     line, a core plate away from it, an ejector retainer
                     carries many counterbores that capture pin heads while its
                     back-up plate below is comparatively plain
  full_thickness_pct how much of the plate is still full thickness. Low means
                     heavily pocketed
  outline_fill_pct   100 = a plain rectangle; lower means a notched or stepped
                     outline
  thru_holes/cbores  counts measured off the mesh, with a size signature.
                     Suffix c = cross-axis (through the side, e.g. a waterline),
                     b = blind
  pocket_regions     [count on top, count on bottom]

Parts marked "measured": false have no exported mesh. Their numbers are the CAD
export's, with the same caveat as before: fill_pct 0 and no holes means the
extractor failed, not that the part is featureless.

YOUR JOB NOW

For every part, give the FINAL role. You are expected to overrule the first pass
wherever the measured geometry disagrees with it -- that is the entire point of
this pass. When you change a name, say in the reason which measurement changed
it. When you keep a name, say which measurement confirms it.

Watch for these specifically:
  * A part the first pass named from missing geometry that now has real numbers.
  * A plate the first pass skipped because it was not full footprint. Ejector
    plates are routinely a third of the mold footprint, and a split ejector set
    appears as two mirror twins at the same height.
  * A part sitting ABOVE the top clamp plate or BELOW the bottom clamp plate. It
    is outside the mold stack, so it is not an ejector plate. An unfeatured thin
    sheet on the outside face is usually insulation.
  * An ejector retainer and its back-up plate being the wrong way round. The
    retainer holds the pin heads and carries far more counterbores.
"""


def _final_payload(p: StackPart) -> dict:
    d: dict = {
        "index": p.index,
        "name": p.short_name,
        "qty": p.qty,
        "stack_center": round(p.stack_pos, 4),
        "pos_x": round(p.center_x, 3),
        "pos_y": round(p.center_y, 3),
        "order_from_top": p.order_from_top or None,
        "in_stack": p.in_stack,
        "plan_area_frac": round(p.plan_area_frac, 3),
        "weight_lb": None if p.weight_lb is None else round(p.weight_lb, 1),
        "first_pass": {
            "role": p.candidate_role or None,
            "confidence": p.candidate_confidence or None,
            "reason": p.candidate_reason or None,
        },
    }
    if p.half:
        d["half"] = p.half
    if p.gap_above is not None:
        d["gap_above"] = p.gap_above
    if p.gap_below is not None:
        d["gap_below"] = p.gap_below
    if p.inside_ejector_box:
        d["inside_ejector_box"] = True
    if p.is_rail_like:
        d["rail_shaped"] = True
    if p.mirror_twins:
        d["mirror_twins"] = p.mirror_twins

    g = p.stl
    if g is not None:
        d["measured"] = True
        d["triangles"] = g.n_triangles
        d["thickness_in"] = round(g.thickness_in, 4)
        d["plan_in"] = [round(g.plan_short_in, 3), round(g.plan_long_in, 3)]
        d["volume_cuin"] = round(g.volume_cuin, 2)
        d["fill_pct"] = round(g.solid_fill_pct, 1)
        d["outline_fill_pct"] = round(g.outline_fill_pct, 1)
        d["full_thickness_pct"] = round(g.full_thickness_pct, 1)
        d["pocket_top"] = [round(g.pocket_top_in2, 2), round(g.pocket_top_max_depth, 3)]
        d["pocket_bottom"] = [round(g.pocket_bottom_in2, 2), round(g.pocket_bottom_max_depth, 3)]
        d["pocket_regions"] = [g.n_pocket_regions_top, g.n_pocket_regions_bottom]
        d["thru_holes"] = g.n_thru_holes
        d["cbores"] = g.n_counterbores
        d["cross_holes"] = g.n_cross_holes
        d["max_bore"] = round(g.max_bore_dia, 3)
        d["hole_sig"] = g.hole_signature
        if abs(g.thickness_in - p.csv_thickness) > 0.002:
            d["cad_thickness_was"] = round(p.csv_thickness, 4)
        if g.notes:
            d["mesh_notes"] = g.notes
    else:
        d["measured"] = False
        d["thickness_in"] = round(p.csv_thickness, 4)
        d["plan_in"] = [round(p.csv_width, 3), round(p.csv_length, 3)]
        d["volume_cuin"] = round(p.csv_volume, 2)
        if p.csv_fill_pct > 0:
            d["fill_pct"] = round(p.csv_fill_pct, 1)
            d["thru_holes"] = p.csv_n_thru
            d["cbores"] = p.csv_n_cbore
            d["cross_holes"] = p.csv_n_cross
            if p.csv_hole_sig:
                d["hole_sig"] = p.csv_hole_sig
            if p.csv_pocket_up or p.csv_pocket_dn:
                d["pocket_up_in2"] = round(p.csv_pocket_up, 2)
                d["pocket_dn_in2"] = round(p.csv_pocket_dn, 2)
        else:
            d["geometry_missing"] = True
    return d


def build_final_prompt(
    model: StackModel,
    csv_path: str | Path,
    roles: list[str],
    first_pass: dict | None = None,
    rules: dict | None = None,
) -> str:
    """Pass 2: final naming, with the mesh measurements in hand."""
    struct = structural_parts(model)
    hardware_json, _ = _hardware_block(model, rules)

    job_facts = {
        "job": model.job,
        "mold_footprint_in": [round(model.footprint_w, 3), round(model.footprint_l, 3)],
        "parts_total": len(model.parts),
        "structural_parts": len(struct),
        "parts_measured_from_mesh": model.stl_measured,
        "ejector_box_stack_span": list(model.ejector_box) if model.ejector_box else None,
        "halves_present": sorted({p.half for p in model.parts if p.half}) or None,
    }
    if first_pass:
        ja = first_pass.get("job_analysis", {}) or {}
        job_facts["first_pass_parting_line"] = ja.get("parting_line", "")

    return f"""/no_think
You are a CMS mold-base geometry interpreter, doing the FINAL PASS.

A first pass proposed a name for each structural part from the CAD dimension
export. Since then those parts have been measured off their exported STL triangle
meshes. Give the FINAL name for each of the {len(struct)} structural parts below.

Return JSON only. No explanation outside JSON. No markdown.

Allowed roles:
{", ".join(roles)}

{FINAL_RULES}
{HARDWARE_INSTRUCTION}

Mold-level facts:
{json.dumps(job_facts, indent=2)}

Structural parts, highest in the stack first -- name every one of these:
{json.dumps([_final_payload(p) for p in struct], separators=(",", ":"))}

Hardware, already named by the geometry rules -- do not re-classify:
{hardware_json}

Return ONLY valid JSON in this exact shape:
{{
  "job_analysis": {{
    "stack_axis": "CenterZ",
    "parting_line": "which two indices the parting line runs between, and why",
    "rules_for_this_job": ["rule 1", "rule 2"]
  }},
  "classifications": [
    {{
      "index": "1",
      "role": "a_plate",
      "confidence": "HIGH|MEDIUM|LOW",
      "reason": "which measurement decided it, and whether it changed the first pass",
      "quote": true,
      "changed_from_first_pass": false
    }}
  ],
  "hardware_corrections": []
}}

Every one of the {len(struct)} structural indices must appear exactly once in
"classifications".
CSV source: {csv_path}
"""


# ─────────────────────────────────────────────────────────────────────────────
# Validation and gap-filling
# ─────────────────────────────────────────────────────────────────────────────
def validate_classifications(
    data: dict, model: StackModel, roles: list[str], expect: set[str] | None = None
) -> tuple[dict, list[str]]:
    """Check an answer against the part list. Returns ``(data, problems)``.

    ``expect`` is the set of indices the prompt actually asked for -- the
    structural parts. Only those are reported missing; the hardware is filled from
    the geometry rules on purpose and its absence here is not a fault.

    Problems are reported rather than repaired blind: the caller decides whether
    to keep the answer, patch its gaps, or fall back wholesale.
    """
    problems: list[str] = []
    role_set = set(roles)
    valid_idx = {p.index for p in model.parts}
    seen: dict[str, str] = {}
    singular_used: dict[str, str] = {}
    cleaned: list[dict] = []

    # A correction is the model saying a part in the hardware list is really a
    # plate. Fold those in as ordinary classifications before validating, so they
    # go through the same role and duplicate checks.
    items = list(data.get("classifications", []))
    for fix in data.get("hardware_corrections", []) or []:
        if isinstance(fix, dict) and str(fix.get("index", "")).strip():
            fix = dict(fix)
            fix["reason"] = f"[pulled out of the hardware list] {fix.get('reason', '')}".strip()
            items.append(fix)

    for item in items:
        idx = str(item.get("index", "")).strip()
        role = str(item.get("role", "")).strip().lower()
        if idx not in valid_idx:
            problems.append(f"index {idx!r} is not a part in this job")
            continue
        if idx in seen:
            problems.append(f"index {idx} classified more than once")
            continue
        conf = str(item.get("confidence", "MEDIUM")).strip().upper()
        if role not in role_set:
            problems.append(f"index {idx}: role {role!r} is not allowed")
            role, conf = "hardware_other", "LOW"
        if role in SINGULAR_ROLES:
            if role in singular_used:
                problems.append(
                    f"index {idx}: second {role} (already index {singular_used[role]}), "
                    f"downgraded to full_footprint_plate"
                )
                role, conf = "full_footprint_plate", "LOW"
            else:
                singular_used[role] = idx
        seen[idx] = role
        cleaned.append(
            {
                "index": idx,
                "role": role,
                "confidence": conf if conf in {"HIGH", "MEDIUM", "LOW"} else "MEDIUM",
                "reason": str(item.get("reason", "")).strip(),
                "quote": bool(item.get("quote", True)),
                "changed_from_first_pass": bool(item.get("changed_from_first_pass", False)),
                "measure": bool(item.get("measure", False)),
            }
        )

    wanted = expect if expect is not None else {p.index for p in model.parts}
    missing = [i for i in wanted if i not in seen]
    if missing:
        problems.append(
            f"{len(missing)} structural parts not classified: {', '.join(missing[:12])}"
            f"{'...' if len(missing) > 12 else ''}"
        )
    data["classifications"] = cleaned
    return data, problems


def fill_missing(data: dict, fallback: dict, model: StackModel, label: str) -> int:
    """Patch parts an answer skipped, taking their role from ``fallback``.

    Qwen returns a short list on jobs with a hundred-plus fasteners. Patched rows
    are tagged in their reason so a reviewer can see where each came from.
    """
    have = {c["index"] for c in data.get("classifications", [])}
    by_index = {str(c.get("index", "")): c for c in fallback.get("classifications", [])}
    added = 0
    for p in model.parts:
        if p.index in have:
            continue
        src = by_index.get(p.index)
        if not src:
            continue
        data["classifications"].append(
            {
                "index": p.index,
                "role": src.get("role", "hardware_other"),
                "confidence": str(src.get("confidence", "LOW")).upper(),
                "reason": f"[{label}] {src.get('reason', '')}".strip(),
                "quote": bool(src.get("quote", False)),
            }
        )
        added += 1
    data["classifications"].sort(
        key=lambda c: int(c["index"]) if str(c["index"]).isdigit() else 0
    )
    return added


def diff_passes(model: StackModel, final: dict) -> list[dict]:
    """Every name the final pass changed, for the log and the run summary."""
    out = []
    by_index = model.by_index()
    for c in final.get("classifications", []):
        p = by_index.get(c["index"])
        if p is None or not p.candidate_role:
            continue
        if p.candidate_role != c["role"]:
            out.append(
                {
                    "index": c["index"],
                    "name": p.short_name,
                    "from": p.candidate_role,
                    "to": c["role"],
                    "measured": p.stl is not None,
                    "reason": c.get("reason", ""),
                }
            )
    return out
