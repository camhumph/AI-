"""Per-job steel-grade overrides, so the estimator can switch a plate's material
from the Parts & Pricing tab.

WHY THIS MIRRORS plate_names.py RATHER THAN REUSING IT
    A rename changes the LABEL and never the ROLE, because the role decides the
    quote row and therefore the money. A grade override is the opposite: changing
    the grade is EXACTLY a change of money, because the grade picks which block of
    the quote workbook the plate is priced in (#1 A-36, #2 4140, #3 P20, ...).

    So the two are stored side by side but kept separate: same override-file
    pattern, same "empty string clears it" rule, same role-key vs cad-index keying
    -- and a completely separate file, so clearing all renames cannot silently
    reprice a job.

WHAT THE SHOP ACTUALLY DOES
    Both hand quotes checked against the macro -- C18597 and C18619 -- put every
    plate of a standard base in the #2 block, while the macro split them per plate
    role (P20 for A/B, A-36 for the rest). The macro default is now #2 to match
    (STD_DEFAULT_GRADE_ALL in Module6121.bas). This module is the escape hatch for
    the jobs that are not #2 throughout, without editing the macro or the workbook.

KEYS
    Same two kinds plate_names uses, so one UI control can drive both:
      "<cad_index>"          e.g. "4"        -- this one CAD part
      "role:<role_key>"      e.g. "role:a_plate" -- every plate with that role
    A cad-index entry wins over a role entry for the same plate.
"""
from __future__ import annotations

import json
import re
from pathlib import Path

OVERRIDE_FILE = "plate_grades.json"

# Prefix marking a key as a role rather than a CAD index. Kept identical to
# plate_names.ROLE_PREFIX so the frontend can build one key for both stores.
ROLE_PREFIX = "role:"
NAME_PREFIX = "name:"

# Roles that identify nothing specific. A role key built from one of these would
# tie every unmapped row together, so setting the grade on one steel row would
# silently reprice all of them. Kept identical to plate_names.GENERIC_ROLES.
GENERIC_ROLES = {"", "steel_plate", "purchased_component", "other", "ignore"}

# The grade blocks the quote workbook actually has. An override outside this set
# is refused rather than written, because StdQuoteRowFor has no rows for it and the
# plate would silently drop off the sheet.
VALID_GRADES = ("A36", "4140", "P20", "420SS", "6061", "A2", "O1")

# What the shop calls each one, for the picker.
GRADE_LABELS = {
    "A36": "#1 A-36",
    "4140": "#2 4140",
    "P20": "#3 P20",
    "420SS": "420 SS",
    "6061": "6061 ALM",
    "A2": "A-2",
    "O1": "O-1",
}


def role_key(role: str) -> str:
    return f"{ROLE_PREFIX}{(role or '').strip()}"


def _norm_name(s: str) -> str:
    """Identical to plate_names._norm_name and normName() in PartsTable.tsx."""
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def name_key(name: str) -> str:
    """Last-resort key for a row with no CAD index and no specific role.

    A BMS steel row carries no CAD index, and any row the sheet parser could not
    map carries role "steel_plate". Without this third key those rows could only be
    keyed by their generic role, and one grade change repriced all of them.
    """
    return f"{NAME_PREFIX}{_norm_name(name)}"


def clean_grade(value: str) -> str:
    """Normalise a submitted grade to one of VALID_GRADES, or "" to clear it."""
    if value is None:
        return ""
    # "#" stripped along with "-" and space, because the shop writes the blocks as
    # "#1 A-36" / "#2" / "#3 P20" and those are exactly the strings the picker and
    # the steel sheet use. Without it "#2" normalised to "#2", matched no alias, and
    # was rejected as unrecognised -- i.e. the most common input silently failed.
    v = str(value).strip().upper().replace("#", "").replace("-", "").replace(" ", "")
    if not v:
        return ""
    aliases = {
        "1": "A36", "A36": "A36", "1A36": "A36",
        "2": "4140", "4140": "4140", "24140": "4140",
        "3": "P20", "P20": "P20", "3P20": "P20",
        "420SS": "420SS", "420": "420SS",
        "6061": "6061", "6061ALM": "6061", "ALM": "6061",
        "A2": "A2", "O1": "O1",
    }
    return aliases.get(v, "")


def _path(job_dir: Path) -> Path:
    return job_dir / OVERRIDE_FILE


def get_overrides(job_dir: Path) -> dict:
    p = _path(job_dir)
    if not p.exists():
        return {}
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return {str(k): str(v) for k, v in data.items() if str(v).strip()}
    except Exception:
        # A corrupt override file must not take the quote down with it.
        return {}


def set_overrides(job_dir: Path, names: dict) -> dict:
    """Merge in new overrides. An empty/None value CLEARS that key."""
    current = get_overrides(job_dir)
    rejected = {}
    for k, v in (names or {}).items():
        key = str(k).strip()
        if not key:
            continue
        if v is None or str(v).strip() == "":
            current.pop(key, None)
            continue
        g = clean_grade(v)
        if not g:
            rejected[key] = str(v)
            continue
        current[key] = g
    job_dir.mkdir(parents=True, exist_ok=True)
    _path(job_dir).write_text(json.dumps(current, indent=2), encoding="utf-8")
    if rejected:
        # Surfaced to the caller rather than swallowed: a grade that did not stick
        # is a priced-wrong plate, and silence there is how a bad quote goes out.
        return {"grades": current, "rejected": rejected}
    return {"grades": current}


def grade_for(overrides: dict, cad_index, role: str, name: str = "") -> str:
    """The override for this plate, or "" when there is none.

    Most specific key first: CAD index, then a SPECIFIC role, then the displayed
    name. Same three tiers plate_names uses and the same order renameKeyFor() in
    PartsTable.tsx builds them in, so one UI control can drive both stores.

    A generic role is skipped rather than matched: "steel_plate" is shared by every
    row the sheet parser could not identify, and honouring it here would let one
    grade change reprice all of them.
    """
    if not overrides:
        return ""
    hit = overrides.get(str(cad_index))
    if hit:
        return hit
    r = (role or "").strip()
    if r and r not in GENERIC_ROLES:
        hit = overrides.get(role_key(r))
        if hit:
            return hit
    if name:
        return overrides.get(name_key(name), "")
    return ""
