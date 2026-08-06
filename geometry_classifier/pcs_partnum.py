#!/usr/bin/env python3
"""Parse PCS Company mold-base item numbers.

WHY
    A PCS item number states the base size, the series and the A/B plate
    thicknesses exactly. When one appears on a BOM line or in a CAD component
    name it is far stronger evidence than a bounding box, and it tells the
    classifier which SERIES it is looking at -- which fixes the expected stack
    order and stops a T-series being named as if it had one parting line.

GRAMMAR
    A / B series:   <nominal><series>-<A thk code>-<B thk code>
                    1016A-13-37  ->  9-7/8 x 16", A-series, A=1-3/8, B=3-7/8

    X series:       <nominal>X<5|6>-<AX thk code>
                    1016X5-23    ->  9-7/8 x 16", 5-plate stripper, AX=2-3/8

    T series:       <nominal>T-<...>
                    Same shape as A, with the X-1/X-2 floating plates implied.

THICKNESS CODES
    Whole inches followed by 3 (= 3/8") or 7 (= 7/8"). So 13 = 1.375,
    17 = 1.875, 23 = 2.375, 37 = 3.875, 57 = 5.875. Single-digit 7 = 0.875.

NOMINAL SIZE IS NOT ACTUAL SIZE
    1016 = 9-7/8 x 16". The WIDTH is the nominal rounded UP to the next whole
    inch; the LENGTH is actual. So nominal 10 means 9.875 wide. Matching a BOM
    row on 10.000 will miss every time -- which is the whole reason this module
    returns both numbers.

    PCS publishes 32 standard sizes from 7-7/8 x 7-7/8" to 23-3/4 x 35-1/2", and
    the width offset is not a single constant across all of them (the largest is
    23-3/4 nominal 24). Where a size is not in the table below the width falls
    back to nominal - 1/8", which is the common case, and `width_exact` says
    False so a caller can decide how much to trust it.
"""
from __future__ import annotations

import re
from typing import Dict, List, Optional

# Known nominal-width -> actual-width, inches. From the published size list.
# Anything not here falls back to nominal - 0.125 with width_exact False.
NOMINAL_WIDTH_ACTUAL: Dict[int, float] = {
    8: 7.875,
    9: 8.875,
    10: 9.875,
    12: 11.875,
    13: 12.875,
    15: 14.875,
    16: 15.875,
    18: 17.875,
    20: 19.875,
    21: 20.875,
    24: 23.75,
}

SERIES_INFO = {
    "A": ("A", "2-plate, with support plate"),
    "B": ("B", "2-plate, no support plate"),
    "T": ("T", "3-plate, adds X-1 runner stripper and X-2 cavity, TWO parting lines"),
    "AX": ("AX", "stripper plate between AX and BX"),
    "X5": ("5X", "5-plate stripper, no support plate"),
    "X6": ("6X", "6-plate stripper, with support plate"),
}

# 1016A-13-37 / 1016B-17-23 / 1016T-13-37 / 1016AX-23 / 1016X5-23 / 1016X6-27
ITEM_RE = re.compile(
    r"""(?<![0-9A-Z])
    (?P<nom>\d{3,5})                 # nominal size, e.g. 1016 or 71678
    (?P<series>AX|X5|X6|A|B|T)       # series -- AX/X5/X6 before A/B
    (?:[-\s]?(?P<t1>\d{1,2}))?       # first thickness code
    (?:[-\s]?(?P<t2>\d{1,2}))?       # second thickness code
    (?![0-9])
    """,
    re.X | re.I,
)


def thickness_from_code(code: str) -> Optional[float]:
    """Decode a PCS thickness code. 13 -> 1.375, 7 -> 0.875, 57 -> 5.875."""
    if not code:
        return None
    c = code.strip()
    if not c.isdigit():
        return None
    last = c[-1]
    if last == "3":
        frac = 0.375
    elif last == "7":
        frac = 0.875
    else:
        # Not a valid PCS code -- codes always end 3 or 7.
        return None
    whole = c[:-1]
    inches = int(whole) if whole else 0
    return round(inches + frac, 4)


def split_nominal(nom: str) -> Optional[tuple]:
    """Split a nominal size token into (width, length) whole inches.

    PCS writes the two dimensions concatenated, which is ambiguous for odd digit
    counts. 4 digits is the common case and splits 2+2. A 5-digit token is
    2+3 (e.g. 21355 would be 21 x 35... but PCS writes 2135 for that), so 5
    digits is treated as 2+3 and flagged by returning None on anything that does
    not yield plausible plate sizes.
    """
    if len(nom) == 4:
        w, l = int(nom[:2]), int(nom[2:])
    elif len(nom) == 3:
        # e.g. 798 -> 7 x 98? No: PCS uses 0708 style. Treat as 1+2.
        w, l = int(nom[0]), int(nom[1:])
    elif len(nom) == 5:
        w, l = int(nom[:2]), int(nom[2:])
    else:
        return None
    # PCS bases run 7-7/8 to 23-3/4 wide and up to 35-1/2 long.
    if not (7 <= w <= 26 and 7 <= l <= 40):
        return None
    return w, l


def parse_item_number(text: str) -> Optional[dict]:
    """First PCS item number found in `text`, or None.

    Returns a plain dict so it drops straight into JSON output.
    """
    if not text:
        return None
    for m in ITEM_RE.finditer(text.upper()):
        nom = m.group("nom")
        dims = split_nominal(nom)
        if not dims:
            continue
        w_nom, l_nom = dims
        series_key = m.group("series").upper()
        series, note = SERIES_INFO.get(series_key, (series_key, ""))

        actual_w = NOMINAL_WIDTH_ACTUAL.get(w_nom)
        width_exact = actual_w is not None
        if actual_w is None:
            actual_w = round(w_nom - 0.125, 4)

        t1 = thickness_from_code(m.group("t1") or "")
        t2 = thickness_from_code(m.group("t2") or "")

        # On the X series there is only one code and it is the AX plate.
        if series in ("5X", "6X", "AX"):
            a_thk, b_thk = t1, None
        else:
            a_thk, b_thk = t1, t2

        return {
            "raw": m.group(0),
            "series": series,
            "series_note": note,
            "nominal_width_in": w_nom,
            "nominal_length_in": l_nom,
            "actual_width_in": actual_w,
            "actual_length_in": float(l_nom),
            "width_exact": width_exact,
            "a_plate_thk_in": a_thk,
            "b_plate_thk_in": b_thk,
        }
    return None


def expected_full_plates(series: str) -> List[str]:
    """Full-footprint plates for a series, top of stack downwards.

    Used to check a classification for completeness: if the series says a plate
    should be there and no CAD part got that role, something was missed.
    """
    if series == "T":
        return [
            "top_clamp_plate", "x1_plate", "x2_plate", "b_plate",
            "support_plate", "bottom_clamp_plate",
        ]
    if series in ("AX", "5X", "6X"):
        base = ["top_clamp_plate", "ax_plate", "stripper_plate", "bx_plate"]
        if series == "6X":
            base.append("support_plate")
        return base + ["bottom_clamp_plate"]
    if series == "B":
        return ["top_clamp_plate", "a_plate", "b_plate", "bottom_clamp_plate"]
    return [
        "top_clamp_plate", "a_plate", "b_plate", "support_plate",
        "bottom_clamp_plate",
    ]


if __name__ == "__main__":
    import json
    import sys

    tests = [
        "1016A-13-37",
        "1016B-17-23",
        "1016T-13-37",
        "1016X5-23",
        "1016X6-27",
        "1218AX-23",
        "PCS 1016A-13-37 MOLD BASE",
        "1016A 13 37",
        "0810A-7-13",
        "2436A-23-47",
        "no part number here",
    ]
    args = sys.argv[1:] or tests
    for t in args:
        r = parse_item_number(t)
        if r:
            print(
                f"{t:<28} -> {r['series']:<3} "
                f"{r['actual_width_in']}x{r['actual_length_in']}"
                f"{'' if r['width_exact'] else ' (width approx)'}  "
                f"A={r['a_plate_thk_in']} B={r['b_plate_thk_in']}"
            )
            print(f"{'':28}    expects: {', '.join(expected_full_plates(r['series']))}")
        else:
            print(f"{t:<28} -> no match")
