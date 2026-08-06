"""Tell apart plates that belong to DIFFERENT mold bases in one assembly.

THE PROBLEM
    C18616 (Dynacast 2223605) looks at first glance like the classifier inventing
    plates: two A plates, two B plates, two clamp plates, two ejector sets, eight
    leader pins, four rails. The shop says there are seven plates.

    They are both right. The assembly holds TWO COMPLETE MOLD BASES, and the CAD
    says so plainly:

        2223605_A-PLATE-1        T=3.375     <- base 1
        2223605_A-PLATE_2-1      T=4.375     <- base 2
        2223605_B-PLATE-1        T=3.879
        2223605_B-PLATE_2-1      T=5.875
        2223605_CLAMP-PLATE-1    T=1.062
        2223605_CLAMP-PLATE_2-1  T=1.375

    Seven plates per base, fourteen in the job. The thicknesses differ, so these
    are not duplicate instances of one plate — they are two different bases quoted
    together.

    Flattening that into one list gives two rows both labelled "B Plate", which is
    worse than the raw CAD name it replaced: identical names, different steel.

THE DISCRIMINATOR
    `_2` immediately before the SolidWorks instance suffix (`-1`, `-2`, ...) is
    the shop's own marker for the second base. It is in the data, not inferred
    from geometry, which is why this is a naming rule and not a guess.

WHAT THIS DELIBERATELY DOES NOT DO
    It does not decide how many molds a job has, and it does not change roles or
    pricing. It only answers "do these two same-role plates come from different
    bases, and if so which". Everything else keeps working exactly as before, and
    a job with one base is untouched.
"""
from __future__ import annotations

import re
from typing import Dict, List, Optional

# `..._2-1` / `..._2` / `..._3-2`.  The group is the base number.
#
# Anchored on the END of the leaf name so it cannot fire on a part number that
# merely contains an underscore-digit — "LDR-PIN_2OD-X-7-3-4_PCS" must NOT read
# as base 2, and it does not, because its `_2` is followed by "OD".
_GROUP_RE = re.compile(r"_(\d+)(?:-\d+)?$")

# SolidWorks instance suffix, e.g. the "-1" in "2223605_A-PLATE-1".
_INSTANCE_RE = re.compile(r"-\d+$")


def leaf(component: str) -> str:
    """Last path segment of a CAD component name."""
    if not component:
        return ""
    return component.replace("\\", "/").split("/")[-1].strip()


def group_of(component: str) -> Optional[int]:
    """Which mold base this component belongs to, or None if unmarked.

    Returns 2 for `2223605_A-PLATE_2-1`, None for `2223605_A-PLATE-1`.
    Unmarked means base 1 by convention, but None is returned rather than 1 so a
    caller can tell "explicitly base 1" from "not marked at all" — a job with no
    markers anywhere must not sprout "(Mold 1)" labels.
    """
    name = leaf(component)
    if not name:
        return None
    m = _GROUP_RE.search(name)
    if m:
        try:
            n = int(m.group(1))
        except ValueError:
            return None
        # A plausible base number. `_12` is a part number, not a 12th mold.
        return n if 2 <= n <= 9 else None
    return None


def strip_group(component: str) -> str:
    """Leaf name with the instance suffix and any base marker removed."""
    name = _INSTANCE_RE.sub("", leaf(component))
    return re.sub(r"_\d+$", "", name)


def _is_plate_role(role: str) -> bool:
    """A structural mold-base plate, i.e. something there is exactly ONE of per base.

    Rails are deliberately EXCLUDED even though they are plate-like steel. Every
    mold base has two of them, so a repeated `rail` role is the normal case and
    carries no information about how many bases are in the assembly. The module
    docstring already said as much ("a repeated role with no marker is more likely
    a real pair (two rails, four pins)") while the test below happily counted rails
    as evidence — see the C18626 note in detect_groups.
    """
    role = (role or "").strip()
    if not role.endswith("_plate"):
        return False
    return role not in {"full_footprint_plate"}


def detect_groups(rows: List[dict]) -> Dict[str, int]:
    """Assign a base number to every row, keyed by CAD index.

    Only returns anything when the assembly genuinely holds more than one base.
    Three things must all hold, and each one is here because dropping it produced a
    false positive on a real job:

      1. A `_N` marker appears, AND
      2. it appears on a PLATE -- not on an insert, a screw or a bushing, AND
      3. some single-per-base plate role actually repeats.

    C18626 (Hewitt 25-424) is why (2) and the rail exclusion in (1)/(3) exist. It
    is ONE mold base, and it was labelled "(Mold 1)" throughout. Every `_2`/`_3`
    marker in that job sits on a cavity/core insert three levels deep:

        25-424--bs-quote-1/25-424--8810-a00-1/25-424--8810-a00_2-1
        25-424--bm-quote-1/25-424--8820-a00-3/25-424--8820-a00_3-1

    Those are insert VARIANTS inside insert sub-assemblies. None of the nine steel
    plates carries a marker at all. The old test passed anyway: markers existed
    somewhere (condition 1), and `rail` appeared twice (condition 3) because every
    mold has two rails. So a single-base job sprouted a mold qualifier on every
    row, which is worse than no label -- it tells the estimator to go looking for a
    second base that does not exist.
    """
    marked = {}
    marked_on_plate = False
    for r in rows:
        comp = r.get("Component") or r.get("component") or ""
        g = group_of(comp)
        if g:
            marked[str(r.get("index"))] = g
            if _is_plate_role(r.get("role") or ""):
                marked_on_plate = True
    if not marked:
        return {}

    # A marker on an insert or a fastener is a part-number variant, not a base.
    if not marked_on_plate:
        return {}

    # Does a single-per-base plate role actually repeat, WITH BOTH COPIES REAL?
    #
    # THE LABEL ONLY EARNS ITS PLACE WHEN THERE ARE TWO ROWS TO TELL APART.
    #
    # C18619 (Dynacast 2223602) is one mold base, and the estimator's hand quote
    # lists six plates. The CAD carries `_2` markers on four plate families, so the
    # marker test and the repeat test both passed and every row was labelled
    # "(Mold 1)" / "(Mold 2)". Measuring the pairs shows what they really are:
    #
    #     122761_B-PLATE          dCtr=0.374  (T1+T2)/2=4.502  live=1/0
    #     122761_A-PLATE          dCtr=0.011  (T1+T2)/2=4.386  live=1/0
    #     122761_EJ-BACKUP-PLATE  dCtr=0.000  (T1+T2)/2=1.125  live=1/0
    #     122761_EJ-RET-PLATE     dCtr=0.000  (T1+T2)/2=0.750  live=1/0
    #
    # Each pair INTERPENETRATES -- centres a few thousandths apart on plates 4.4"
    # thick -- so they occupy one physical slot: two revisions of one plate, not two
    # plates. And exactly ONE copy of each survived the STEP import.
    #
    # Liveness is the gate rather than the overlap distance, because it is exact and
    # needs no tolerance: a plate with no solid cannot be measured, so it never
    # reaches the quote. If only one copy of a role survives there is a single row
    # for it, and a mold qualifier on a single row tells the estimator to go looking
    # for a second base that is not there. That holds even if the job genuinely has
    # two bases -- with one copy dead there is still only one row to name.
    #
    # MassLb is 0 for a surfaces-only import (see PartHasNoSolid in Module6121.bas,
    # same signature: real bounding box, no mass).
    def _is_live(r: dict) -> bool:
        try:
            return float(r.get("MassLb") or r.get("Mass_or_Vol") or 0) > 0
        except (TypeError, ValueError):
            return False

    plate_roles: Dict[str, int] = {}
    for r in rows:
        role = (r.get("role") or "").strip()
        if _is_plate_role(role) and _is_live(r):
            plate_roles[role] = plate_roles.get(role, 0) + 1
    if not any(v > 1 for v in plate_roles.values()):
        return {}

    out: Dict[str, int] = {}
    for r in rows:
        idx = str(r.get("index"))
        out[idx] = marked.get(idx, 1)
    return out


def group_count(groups: Dict[str, int]) -> int:
    return len(set(groups.values())) if groups else 1


def label_with_group(label: str, group: Optional[int], total_groups: int) -> str:
    """Append a mold-base qualifier when there is more than one base."""
    if not group or total_groups < 2:
        return label
    return f"{label} (Mold {group})"
