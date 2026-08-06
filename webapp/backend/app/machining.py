"""Machining time estimates for quoted steel.

DESIGN PRINCIPLE
    A foreman has to be able to argue with this. So every number is derived from
    geometry the macro actually measured, every rate is named and editable, and
    the output shows the arithmetic instead of a single opaque hour figure. An
    estimate you cannot audit is worse than no estimate, because someone will
    quote off it.

WHAT IS MEASURED vs WHAT IS ASSUMED
    Measured (trustworthy -- comes from the CAD scan):
      * stock volume        = T x W x L from the quote row
      * finished volume     = mass / density, or the STL mesh volume
      * material removed    = stock - finished   <- dominates roughing time
      * footprint / faces   = from T, W, L
      * material            = from the BOM row

    Assumed (tunable -- flagged in the output):
      * removal + feed rates per material
      * hole and tap counts, seeded per role
      * setup minutes per part

    The geometry terms carry the estimate. The hole counts are the weakest input,
    which is why they are reported separately and can be overridden rather than
    buried in a total.

RATES
    Defaults are deliberately conservative mid-range shop numbers for a 40-taper
    machining centre with carbide tooling. Override any of them by adding a
    "machining" block to pricing_config.json -- see DEFAULT_RATES for the shape.
    Nothing here is secret or hard-coded past the point of editing.
"""
from __future__ import annotations

import json
import math

from . import config

# Density lb/in^3, used to turn measured mass back into finished volume.
DENSITY_LB_PER_CUIN = {
    "4140": 0.284,
    "4130": 0.284,
    "P20": 0.284,
    "H13": 0.281,
    "A-2": 0.284,
    "D-2": 0.278,
    "S7": 0.283,
    "A36": 0.284,
    "1030": 0.284,
    "STAINLESS": 0.29,
    "420SS": 0.28,
    "ALUMINUM": 0.098,
    "6061": 0.098,
}
DEFAULT_DENSITY = 0.284

# Per-material cutting rates.
#
#   rough_mrr        in^3/min of material removed while roughing
#   finish_area      in^2/min of finish-milled surface
#   grind_area       in^2/min of ground face
#   drill_ipm        in/min of drilled depth (includes pecking)
#   tap_ipm          in/min of tapped depth (includes reversal)
#   hardness_note    why the numbers look the way they do
DEFAULT_RATES = {
    "materials": {
        "A36":  {"rough_mrr": 6.5, "finish_area": 11.0, "grind_area": 26.0, "drill_ipm": 4.0, "tap_ipm": 1.6,
                 "hardness_note": "Soft hot-rolled: fastest removal, poorest finish."},
        "1030": {"rough_mrr": 6.0, "finish_area": 10.5, "grind_area": 25.0, "drill_ipm": 3.8, "tap_ipm": 1.5},
        "4140": {"rough_mrr": 4.2, "finish_area": 8.5,  "grind_area": 20.0, "drill_ipm": 2.6, "tap_ipm": 1.0,
                 "hardness_note": "Pre-hardened 28-34 Rc. The shop's bread-and-butter plate steel."},
        "4130": {"rough_mrr": 4.4, "finish_area": 8.8,  "grind_area": 21.0, "drill_ipm": 2.7, "tap_ipm": 1.1},
        "P20":  {"rough_mrr": 4.0, "finish_area": 8.2,  "grind_area": 19.0, "drill_ipm": 2.5, "tap_ipm": 1.0},
        "A-2":  {"rough_mrr": 2.6, "finish_area": 6.0,  "grind_area": 14.0, "drill_ipm": 1.6, "tap_ipm": 0.6,
                 "hardness_note": "Air-hardening tool steel: slow, and hard on tooling."},
        "D-2":  {"rough_mrr": 2.2, "finish_area": 5.4,  "grind_area": 12.0, "drill_ipm": 1.3, "tap_ipm": 0.5},
        "H13":  {"rough_mrr": 2.4, "finish_area": 5.8,  "grind_area": 13.0, "drill_ipm": 1.4, "tap_ipm": 0.55},
        "S7":   {"rough_mrr": 2.5, "finish_area": 5.9,  "grind_area": 13.5, "drill_ipm": 1.5, "tap_ipm": 0.58},
        "STAINLESS": {"rough_mrr": 2.0, "finish_area": 5.0, "grind_area": 11.0, "drill_ipm": 1.2, "tap_ipm": 0.45,
                      "hardness_note": "Work-hardens: keep the tool moving or it glazes."},
        "6061": {"rough_mrr": 18.0, "finish_area": 28.0, "grind_area": 40.0, "drill_ipm": 12.0, "tap_ipm": 4.0},
    },
    "default_material": "4140",

    # Fixed minutes per part, independent of size.
    "setup": {
        "per_part_min": 25.0,
        "per_extra_face_min": 8.0,
        "program_min_per_part": 15.0,
    },

    # Finishing is applied to a fraction of total surface area -- you do not
    # finish-mill the faces that get ground, or the ones nobody sees.
    "finish_area_fraction": 0.55,

    # Grinding: mold plates are ground on the two big faces at minimum.
    "ground_faces": 2,

    # Hole/tap counts per role. THE WEAKEST INPUT IN THE MODEL -- the CAD export
    # carries no hole data, so these are pattern-based estimates from typical
    # mold-base practice. Shown separately in the output for exactly that reason.
    "holes_by_role": {
        "TCP":            {"holes": 26, "avg_depth": 1.6, "taps": 12, "avg_tap_depth": 1.1},
        "BCP":            {"holes": 26, "avg_depth": 1.6, "taps": 12, "avg_tap_depth": 1.1},
        "TOP CLAMP PLATE":    {"holes": 24, "avg_depth": 1.6, "taps": 10, "avg_tap_depth": 1.1},
        "BOTTOM CLAMP PLATE": {"holes": 24, "avg_depth": 1.6, "taps": 10, "avg_tap_depth": 1.1},
        "A PLATE":        {"holes": 34, "avg_depth": 3.0, "taps": 16, "avg_tap_depth": 1.4},
        "B PLATE":        {"holes": 34, "avg_depth": 3.0, "taps": 16, "avg_tap_depth": 1.4},
        "SUPPORT PLATE":  {"holes": 22, "avg_depth": 2.2, "taps": 8,  "avg_tap_depth": 1.2},
        "ID HOLDER":      {"holes": 30, "avg_depth": 3.2, "taps": 14, "avg_tap_depth": 1.3},
        "OD HOLDER":      {"holes": 30, "avg_depth": 3.2, "taps": 14, "avg_tap_depth": 1.3},
        "ID POT BLOCK":   {"holes": 16, "avg_depth": 2.6, "taps": 8,  "avg_tap_depth": 1.2},
        "OD POT BLOCK":   {"holes": 16, "avg_depth": 2.6, "taps": 8,  "avg_tap_depth": 1.2},
        "RAILS":          {"holes": 10, "avg_depth": 1.5, "taps": 6,  "avg_tap_depth": 1.0},
        "EJECTOR PLATE":  {"holes": 40, "avg_depth": 0.9, "taps": 6,  "avg_tap_depth": 0.8},
        "BOTTOM EJECTOR PLATE": {"holes": 40, "avg_depth": 0.9, "taps": 6, "avg_tap_depth": 0.8},
        "_DEFAULT":       {"holes": 12, "avg_depth": 1.5, "taps": 4,  "avg_tap_depth": 1.0},
    },

    # Shop efficiency. Cut time is not wall-clock time: tool changes, probing,
    # chip clearing, inspection and the operator all live in here.
    "efficiency_factor": 0.72,

    "shop_rate_per_hour": 85.0,
}


def _load_rates() -> dict:
    """Defaults, deep-merged with any "machining" block in pricing_config.json."""
    rates = json.loads(json.dumps(DEFAULT_RATES))  # deep copy
    try:
        path = config.PRICING_CONFIG_PATH
        if path.exists():
            cfg = json.loads(path.read_text(encoding="utf-8"))
            override = cfg.get("machining")
            if isinstance(override, dict):
                _deep_merge(rates, override)
    except Exception:
        pass
    return rates


def _deep_merge(base: dict, patch: dict) -> None:
    for k, v in patch.items():
        if isinstance(v, dict) and isinstance(base.get(k), dict):
            _deep_merge(base[k], v)
        else:
            base[k] = v


def _norm_material(raw: str) -> str:
    s = (raw or "").upper().strip()
    if not s:
        return ""
    for key in ("4140", "4130", "P20", "H13", "D-2", "D2", "A-2", "A2", "S7",
                "A36", "1030", "6061", "420", "STAINLESS"):
        if key in s:
            if key in ("D2",):
                return "D-2"
            if key in ("A2",):
                return "A-2"
            if key == "420":
                return "STAINLESS"
            return key
    return ""


def _f(v, default=0.0) -> float:
    try:
        f = float(v)
        return f if math.isfinite(f) else default
    except (TypeError, ValueError):
        return default


def estimate_part(row: dict, rates: dict, mass_lb: float | None = None) -> dict:
    """Minutes for one quoted steel row, broken out by operation."""
    t = _f(row.get("thickness"))
    w = _f(row.get("width"))
    l = _f(row.get("length"))
    qty = max(1, int(_f(row.get("qty"), 1)))

    name = str(row.get("component") or row.get("role_label") or "").strip()
    role_key = str(row.get("role_label") or row.get("component") or "").upper().strip()

    mat_key = _norm_material(str(row.get("material") or ""))
    mat_tbl = rates["materials"].get(mat_key or rates["default_material"]) \
        or rates["materials"][rates["default_material"]]
    used_material = mat_key or rates["default_material"]

    if t <= 0 or w <= 0 or l <= 0:
        return {
            "component": name,
            "role_label": row.get("role_label", ""),
            "material": used_material,
            "qty": qty,
            "skipped": True,
            "reason": "No stock dimensions on this row -- nothing to estimate from.",
            "total_min": 0.0,
            "total_hours": 0.0,
        }

    stock_vol = t * w * l
    density = DENSITY_LB_PER_CUIN.get(used_material, DEFAULT_DENSITY)

    # Finished volume: prefer measured mass (SolidWorks), else assume a pocketed
    # plate keeps ~82% of its stock. Stated either way so nobody has to guess
    # which happened.
    if mass_lb and mass_lb > 0:
        finished_vol = mass_lb / density
        finished_basis = "measured mass"
    else:
        finished_vol = stock_vol * 0.82
        finished_basis = "assumed 82% of stock"

    removed = max(0.0, stock_vol - finished_vol)

    # --- operations ------------------------------------------------------
    rough_min = removed / max(0.1, _f(mat_tbl.get("rough_mrr"), 4.0))

    face_area = w * l
    side_area = 2 * (t * w) + 2 * (t * l)
    total_area = 2 * face_area + side_area

    finish_area = total_area * _f(rates.get("finish_area_fraction"), 0.55)
    finish_min = finish_area / max(0.1, _f(mat_tbl.get("finish_area"), 8.0))

    grind_area = face_area * max(0, int(_f(rates.get("ground_faces"), 2)))
    grind_min = grind_area / max(0.1, _f(mat_tbl.get("grind_area"), 20.0))

    holes_cfg = rates["holes_by_role"].get(role_key) or rates["holes_by_role"]["_DEFAULT"]
    n_holes = int(_f(holes_cfg.get("holes"), 0))
    hole_depth = _f(holes_cfg.get("avg_depth"), 1.0)
    n_taps = int(_f(holes_cfg.get("taps"), 0))
    tap_depth = _f(holes_cfg.get("avg_tap_depth"), 1.0)

    drill_min = (n_holes * hole_depth) / max(0.1, _f(mat_tbl.get("drill_ipm"), 2.5))
    tap_min = (n_taps * tap_depth) / max(0.1, _f(mat_tbl.get("tap_ipm"), 1.0))

    setup = rates["setup"]
    setup_min = _f(setup.get("per_part_min"), 25.0) + _f(setup.get("program_min_per_part"), 15.0)

    cut_min = rough_min + finish_min + grind_min + drill_min + tap_min

    eff = _f(rates.get("efficiency_factor"), 0.72)
    eff = min(1.0, max(0.2, eff))
    # Cut time is not wall-clock. Dividing by efficiency adds the tool changes,
    # probing, chip clearing and inspection that never show up in a feed rate.
    per_part_min = (cut_min / eff) + setup_min
    total_min = per_part_min * qty

    return {
        "component": name,
        "role_label": row.get("role_label", ""),
        "material": used_material,
        "qty": qty,
        "skipped": False,
        "stock": {"thickness": t, "width": w, "length": l},
        "stock_volume_cuin": round(stock_vol, 2),
        "finished_volume_cuin": round(finished_vol, 2),
        "finished_basis": finished_basis,
        "removed_cuin": round(removed, 2),
        "removed_pct": round(100.0 * removed / stock_vol, 1) if stock_vol else 0.0,
        "operations": [
            {"op": "Rough mill", "min": round(rough_min, 1), "measured": True,
             "detail": f"{removed:,.1f} in3 removed at {mat_tbl.get('rough_mrr')} in3/min"},
            {"op": "Finish mill", "min": round(finish_min, 1), "measured": True,
             "detail": f"{finish_area:,.0f} in2 at {mat_tbl.get('finish_area')} in2/min"},
            {"op": "Grind", "min": round(grind_min, 1), "measured": True,
             "detail": f"{grind_area:,.0f} in2 over {rates.get('ground_faces')} face(s)"},
            {"op": "Drill", "min": round(drill_min, 1), "measured": False,
             "detail": f"{n_holes} holes x {hole_depth}\" (estimated pattern)"},
            {"op": "Tap", "min": round(tap_min, 1), "measured": False,
             "detail": f"{n_taps} taps x {tap_depth}\" (estimated pattern)"},
            {"op": "Setup + program", "min": round(setup_min, 1), "measured": False,
             "detail": "fixed per part"},
        ],
        "cut_min": round(cut_min, 1),
        "efficiency_factor": eff,
        "per_part_min": round(per_part_min, 1),
        "total_min": round(total_min, 1),
        "total_hours": round(total_min / 60.0, 2),
    }


def estimate_job(quote_sheet: dict, parts: list | None = None) -> dict:
    """Machining estimate for every steel row in a quote sheet."""
    rates = _load_rates()

    # Measured mass per component name, so finished volume is real where we have it.
    mass_by_component: dict[str, float] = {}
    for p in parts or []:
        comp = str(p.get("Component") or "").strip().lower()
        lb = _f(p.get("MassLb"))
        if comp and lb > 0:
            mass_by_component[comp] = lb

    # Pick ONE source of plates, never the union.
    #
    # On a standard base, pricing.build_quote_sheet emits a "classified" row for
    # every classified CAD part AND appends the steel-workbook rows for the same
    # physical plates -- both land in line_items. Iterating line_items therefore
    # counted each plate twice and doubled the hours. Prefer the workbook (it is
    # the priced source of truth), fall back to classified, fall back to the
    # flat list only if the sheet has no sections at all.
    secs = quote_sheet.get("sections") or {}
    steel_rows = secs.get("steel") or []
    classified_rows = secs.get("classified") or []

    if steel_rows:
        items = steel_rows
        source = "steel workbook"
    elif classified_rows:
        items = classified_rows
        source = "AI classification"
    else:
        items = [
            r for r in (quote_sheet.get("line_items") or [])
            if str(r.get("section") or "").lower() in ("steel", "classified")
        ]
        source = "line items"

    results = []
    for row in items:
        if not _is_machined_row(row):
            continue
        comp = str(row.get("component") or "").strip().lower()
        results.append(estimate_part(row, rates, mass_by_component.get(comp)))

    priced = [r for r in results if not r["skipped"]]
    total_min = sum(r["total_min"] for r in priced)
    measured_min = sum(
        o["min"] for r in priced for o in r["operations"] if o["measured"]
    )
    estimated_min = sum(
        o["min"] for r in priced for o in r["operations"] if not o["measured"]
    )

    rate = _f(rates.get("shop_rate_per_hour"), 85.0)
    hours = total_min / 60.0

    return {
        "parts": results,
        "summary": {
            "part_count": len(priced),
            "skipped_count": len(results) - len(priced),
            "total_minutes": round(total_min, 1),
            "total_hours": round(hours, 2),
            "shop_rate_per_hour": rate,
            "estimated_cost": round(hours * rate, 2),
            # How much of the number rests on measured geometry vs assumed
            # patterns. Below ~60% measured, treat the total as a rough order.
            "measured_minutes": round(measured_min, 1),
            "assumed_minutes": round(estimated_min, 1),
            "confidence_pct": round(
                100.0 * measured_min / (measured_min + estimated_min), 0
            ) if (measured_min + estimated_min) > 0 else 0.0,
        },
        "rates_used": {
            "efficiency_factor": rates.get("efficiency_factor"),
            "finish_area_fraction": rates.get("finish_area_fraction"),
            "ground_faces": rates.get("ground_faces"),
            "setup": rates.get("setup"),
            "shop_rate_per_hour": rate,
        },
        "notes": [
            "Rough / finish / grind come from measured CAD volume and area.",
            "Drill and tap counts are pattern estimates -- the CAD export carries "
            "no hole data. They are listed separately so you can see their weight.",
            "Cut time is divided by the efficiency factor to cover tool changes, "
            "probing, chip clearing and inspection.",
            "Override any rate with a \"machining\" block in pricing_config.json.",
        ],
    }
