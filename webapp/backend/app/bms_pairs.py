"""Stop ID and OD getting swapped on pot-block jobs.

THE PROBLEM
    A BMS base has two pairs of plates that are nearly the same size as each
    other and are told apart only by which one has been hollowed out:

        ID Holder / OD Holder
        ID Pot    / OD Pot

    Whatever names those four end up with is what goes on the steel order, the
    quote sheet, the STL files and the machining plan. Get the pair backwards and
    every one of them is wrong together, consistently, which is the worst kind of
    wrong: nothing looks broken.

    They do get backwards. On C18328 BOTH pairs are, measured against the CAD:

        currently called      bbox in3     solid in3   fill%   max bore
        OD Holder                572.3         452.8   79.1%      1.050
        ID Holder                497.6         351.3   70.6%      1.125
        OD Pot                   208.0         135.5   65.2%      3.030   <-- bored
        ID Pot                   180.6         166.9   92.4%      0.638

THE SHOP'S RULE, WHICH IS THE ONE THIS ENCODES
    Of each pair, ONE member is hollowed out by a big bore down its centre. That
    member has less steel in it, so at comparable size it is the lighter of the
    two. Which name that member gets is not symmetric between the pairs:

        holders   the bored, lighter one is the OD HOLDER
        pots      the bored, lighter one is the ID POT  (a ~2.75" bore)

    So on C18328 the plate with the 3.030" bore and 65% fill is being called OD
    Pot and is really the ID Pot, and the lighter holder is being called ID
    Holder and is really the OD Holder. Both pairs are backwards.

WHY NOT THICKNESS
    Because thickness is what the current code leans on, and thickness is not the
    discriminator. On C18328 the bored pot is the THICKER of its pair (6.875 vs
    5.970) and still has a third less steel in it, because a 3" hole through 6.875"
    removes more than the extra thickness adds. Ranking the pair by size gets it
    exactly backwards on that job. Hollowness is the property that actually names
    these plates, so hollowness is what this measures.

WHAT IT MEASURES, BEST EVIDENCE FIRST
    1. MaxBoreDia   -- a big central bore is the defining feature. Direct.
    2. SolidFillPct -- solid volume over bounding box. Already size-normalised,
                       so it compares two plates of different sizes honestly.
    3. Mass_or_Vol  -- proportional to solid volume (verified: the C18328 pot pair
                       gives 1.2314 by mass against 1.2312 by fill). The fallback
                       for older exports that predate the feature columns, and
                       sound there because a pair's footprints are near-identical.

    Anything less than a clear separation returns "unsure" and changes nothing. A
    coin-flip that silently renames two plates is worse than saying so.
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Which member of each pair is the hollow one. This is the whole shop rule.
BORED_MEMBER = {
    "holder": "od_holder",
    "pot": "id_pot",
}
SOLID_MEMBER = {
    "holder": "id_holder",
    "pot": "od_pot",
}

PAIR_OF_ROLE = {
    "id_holder": "holder",
    "od_holder": "holder",
    "id_pot": "pot",
    "od_pot": "pot",
}

# A bore only counts as "the big bore down the centre" at this size or over, and
# only when it clearly beats the other member's. The pot bore the shop quotes is
# 2.75"; 1.5" is well clear of the leader-pin and waterline holes that riddle
# both members without meaning anything.
BIG_BORE_MIN_IN = 1.5
BORE_RATIO_MIN = 1.5

# Separation thresholds. Below these the two plates are not meaningfully
# different and the answer is "unsure", not a guess.
FILL_SEPARATION_PTS = 5.0
MASS_SEPARATION_FRAC = 0.05

# Dimensions match a CAD row to a sheet row. Generous, because the sheet carries
# stock sizes rounded to the quarter while the CSV carries finished geometry.
DIM_TOL_IN = 0.35


def _f(v: Any, default: float = 0.0) -> float:
    try:
        s = str(v).strip()
        return float(s) if s else default
    except (TypeError, ValueError):
        return default


def load_cad_rows(job_dir: Path) -> List[Dict[str, Any]]:
    """Every solid from the macro's XT export, with the columns that matter."""
    for name in ("XT_Export_CAD_Dimensions.csv", "pdf/XT_Export_CAD_Dimensions.csv",
                 "documents/XT_Export_CAD_Dimensions.csv"):
        p = Path(job_dir) / name
        if not p.exists():
            continue
        try:
            with p.open("r", encoding="utf-8-sig", newline="") as fh:
                rows = []
                for r in csv.DictReader(fh):
                    rows.append(
                        {
                            "index": str(r.get("Index") or "").strip(),
                            "thickness": _f(r.get("Thickness")),
                            "width": _f(r.get("Width")),
                            "length": _f(r.get("Length")),
                            "bbox": _f(r.get("BBoxVolume_cuin")),
                            "mass": _f(r.get("Mass_or_Vol")),
                            "max_bore": _f(r.get("MaxBoreDia")),
                            # Absent on exports older than the feature columns.
                            "fill_pct": _f(r.get("SolidFillPct"), -1.0),
                        }
                    )
                return rows
        except OSError:
            return []
    return []


def _dims_match(a: Dict[str, Any], t: float, w: float, l: float) -> bool:
    """Same plate, allowing for stock-vs-finished rounding and swapped W/L."""
    if abs(a["thickness"] - t) > DIM_TOL_IN:
        return False
    got = sorted([a["width"], a["length"]])
    want = sorted([w, l])
    return abs(got[0] - want[0]) <= DIM_TOL_IN and abs(got[1] - want[1]) <= DIM_TOL_IN


def find_cad_row(cad: List[Dict[str, Any]], t: float, w: float, l: float) -> Optional[Dict[str, Any]]:
    hits = [r for r in cad if _dims_match(r, t, w, l)]
    if len(hits) == 1:
        return hits[0]
    if not hits:
        return None
    # Several candidates: take the closest on thickness, which is the dimension
    # that separates the two members of a pair.
    return sorted(hits, key=lambda r: abs(r["thickness"] - t))[0]


def which_is_bored(
    a: Dict[str, Any], b: Dict[str, Any]
) -> Tuple[Optional[str], str, str]:
    """Which of two plates is the hollow one.

    Returns (index_of_bored, evidence_code, human_explanation). The index is None
    when the two cannot be told apart, and then nothing should be renamed.
    """
    # 1. A big central bore, clearly bigger than the other's.
    ba, bb = a["max_bore"], b["max_bore"]
    if max(ba, bb) >= BIG_BORE_MIN_IN:
        big, small = (a, b) if ba > bb else (b, a)
        hi, lo = max(ba, bb), min(ba, bb)
        if lo <= 0 or hi / max(lo, 1e-6) >= BORE_RATIO_MIN:
            return (
                big["index"],
                "bore",
                f'a {hi:.3f}" bore through it against {lo:.3f}" on the other',
            )

    # 2. Solid fill: how much of the bounding box is actually steel.
    fa, fb = a["fill_pct"], b["fill_pct"]
    if fa >= 0 and fb >= 0 and abs(fa - fb) >= FILL_SEPARATION_PTS:
        low = a if fa < fb else b
        return (
            low["index"],
            "fill",
            f"{min(fa, fb):.1f}% of its bounding box is steel against "
            f"{max(fa, fb):.1f}% on the other",
        )

    # 3. Mass. Proportional to solid volume, and a pair's footprints are close
    #    enough that comparing raw mass is fair.
    ma, mb = a["mass"], b["mass"]
    if ma > 0 and mb > 0 and abs(ma - mb) / max(ma, mb) >= MASS_SEPARATION_FRAC:
        light = a if ma < mb else b
        return (
            light["index"],
            "mass",
            f"{min(ma, mb):.3f} against {max(ma, mb):.3f} — "
            f"{100 * (1 - min(ma, mb) / max(ma, mb)):.0f}% less material in it",
        )

    return (
        None,
        "unsure",
        "the two are within a few percent on bore, fill and mass, so there is "
        "nothing here that says which one is hollow",
    )


def check_pair(
    pair: str,
    current: Dict[str, Dict[str, Any]],
) -> Dict[str, Any]:
    """Is this pair the right way round?

    `current` maps role -> the CAD row currently carrying that name, e.g.
    {"id_pot": <row>, "od_pot": <row>}.
    """
    bored_role = BORED_MEMBER[pair]
    solid_role = SOLID_MEMBER[pair]
    a = current.get(bored_role)
    b = current.get(solid_role)
    out: Dict[str, Any] = {
        "pair": pair,
        "swap": False,
        "confidence": "none",
        "why": "",
        "roles": {},
    }
    if not a or not b:
        out["why"] = "only one of the two plates was found in the CAD export."
        return out

    bored_index, code, detail = which_is_bored(a, b)
    out["evidence"] = code
    if bored_index is None:
        out["confidence"] = "unsure"
        out["why"] = f"Left alone: {detail}."
        return out

    out["confidence"] = {"bore": "high", "fill": "medium", "mass": "medium"}[code]
    if bored_index == a["index"]:
        out["why"] = (
            f"Correct as named. The plate called {_pretty(bored_role)} is the hollow one — "
            f"{detail}."
        )
        return out

    out["swap"] = True
    out["roles"] = {bored_role: b["index"], solid_role: a["index"]}
    out["why"] = (
        f"SWAPPED. The plate now called {_pretty(solid_role)} is the hollow one — {detail} — "
        f"and on a pot base the hollow one of that pair is the {_pretty(bored_role)}. "
        f"The two names belong the other way round."
    )
    return out


def _pretty(role: str) -> str:
    return {
        "id_holder": "ID Holder",
        "od_holder": "OD Holder",
        "id_pot": "ID Pot",
        "od_pot": "OD Pot",
    }.get(role, role)


def audit(job_dir: Path, sheet_rows: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Check both pairs of a pot job against the CAD.

    `sheet_rows` is what the workbook currently says: a list of
    {"role", "name", "thickness", "width", "length"} — i.e. the steel rows the
    quote builder already produces.

    Returns a report and never raises. A rename this drives is a real edit to the
    shop's paperwork, so it reports rather than acts.
    """
    report: Dict[str, Any] = {"ok": False, "pairs": [], "renames": [], "reason": ""}
    cad = load_cad_rows(job_dir)
    if not cad:
        report["reason"] = "No XT_Export_CAD_Dimensions.csv for this job, so there is nothing to check against."
        return report

    by_role: Dict[str, Dict[str, Any]] = {}
    names: Dict[str, str] = {}
    for r in sheet_rows:
        role = (r.get("role") or "").strip()
        if role not in PAIR_OF_ROLE:
            continue
        hit = find_cad_row(cad, _f(r.get("thickness")), _f(r.get("width")), _f(r.get("length")))
        if hit:
            by_role[role] = hit
            names[role] = str(r.get("name") or _pretty(role))

    report["ok"] = True
    for pair in ("holder", "pot"):
        present = {k: v for k, v in by_role.items() if PAIR_OF_ROLE[k] == pair}
        if len(present) < 2:
            continue
        res = check_pair(pair, present)
        res["names"] = {k: names.get(k, _pretty(k)) for k in present}
        report["pairs"].append(res)
        if res["swap"]:
            # The two sheet rows trade names. Emitted in the shape
            # sheet_rename.rewrite_names takes, so the fix is one call away.
            bored_role = BORED_MEMBER[pair]
            solid_role = SOLID_MEMBER[pair]
            report["renames"].append(
                {"from": names.get(bored_role, _pretty(bored_role)), "to": _pretty(solid_role)}
            )
            report["renames"].append(
                {"from": names.get(solid_role, _pretty(solid_role)), "to": _pretty(bored_role)}
            )

    if not report["pairs"]:
        report["reason"] = "No ID/OD holder or pot pair on this job."
    return report
