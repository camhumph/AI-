"""Per-job plate name overrides.

WHAT THIS IS
    A rename typed in the web app, remembered for that job, and applied
    everywhere a plate name is shown or written -- the parts table, the 3D
    gallery, the Geometry and Machining tabs, the bridge file the macro reads
    back, and (via sheet_rename.py) the quote sheet and steel sheet themselves.

THE ONE RULE: A RENAME CHANGES THE LABEL, NEVER THE ROLE
    `role` (a_plate, support_plate, ...) is what decides the quote row and the
    price. `StdQuoteRowFor()` in the macro maps slot -> row number, and
    `price_for_part()` looks up stock dimensions by role. So renaming "A Plate"
    to "Cavity Plate" must leave `role` as `a_plate` and move nothing.

    This is deliberate and it is the whole safety story. If a rename re-keyed
    the role, typing a new label into a text box would silently move money
    between rows of the grade block -- and nothing in the UI would show that it
    had happened. Relabelling for a customer's benefit is a display concern;
    re-slotting a plate is a pricing decision and needs to be an explicit,
    separate action.

    A consequence worth knowing: after renaming "A Plate" to "Cavity Plate" the
    part still prices on the A row. That is correct, and `role` is returned
    alongside the override so the UI can show it.

TWO KINDS OF KEY
    `"7"`             a CAD index — one specific solid in the assembly.
    `"role:a_plate"`  a role — every row for that role, wherever it came from.

    Both are needed because the parts table draws rows from two places. CAD rows
    carry a real CAD index. STEEL rows come out of the .xls workbook and are
    numbered `S1`, `S2`, ... — those are row positions in a table, not part
    identities, and they change whenever the workbook does.

    Storing a steel row's rename under `S1` is what made renaming look broken:
    the key matched no CAD part, so `apply_to_rows` never applied it,
    `override_by_role` never saw it, the sheet rewrite had no old name to search
    for, and the user got a text box that accepted input and then did nothing.
    Steel rows are therefore keyed by ROLE, which is the thing they actually
    identify.

WHERE IT LIVES
    `meta.json` in the job folder, under `plate_name_overrides`.

    Keyed on CAD index rather than on role or on the old name because the index
    is the only stable identity: `classify_job()` overwrites classification.json
    wholesale (jobs.py), so a re-classify would drop overrides keyed on anything
    the classifier produces. meta.json survives re-classify untouched.

    Chosen over a new file because `update_meta()` already exists, is already
    per-job, and already survives the one operation that destroys everything
    else.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, Optional

META_KEY = "plate_name_overrides"

# Long enough for "Manifold Backing Plate" plus a job-specific qualifier, short
# enough that it still fits column A of the quote sheet and column B of the
# steel sheet without spilling.
MAX_NAME_LEN = 48

# Excel formula injection is the real risk here: a name starting = + - @ is
# interpreted as a formula when the workbook is opened, and these names get
# written straight into cells. Also strip the double quote, because the macro
# already strips it when writing the steel sheet
# (`Replace(stdName(i), Chr(34), "")`) and a name that changes on the way to the
# sheet cannot be matched back.
_BAD_LEAD = ("=", "+", "-", "@", "\t", "\r")


def clean_name(raw: str) -> str:
    """Normalise a user-typed name, or return "" if it is unusable."""
    if raw is None:
        return ""
    s = str(raw).replace('"', "").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"[\x00-\x1f]", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    while s and s[0] in _BAD_LEAD:
        s = s[1:].strip()
    return s[:MAX_NAME_LEN]


def _meta_path(job_dir: Path) -> Path:
    return job_dir / "meta.json"


def _read_meta(job_dir: Path) -> dict:
    p = _meta_path(job_dir)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}


def get_overrides(job_dir: Path) -> Dict[str, str]:
    """{cad_index: name}. Always strings, always cleaned."""
    meta = _read_meta(job_dir)
    raw = meta.get(META_KEY) or {}
    if not isinstance(raw, dict):
        return {}
    out: Dict[str, str] = {}
    for k, v in raw.items():
        name = clean_name(v)
        if name:
            out[str(k)] = name
    return out


def _canonical_key(key: str) -> str:
    """Normalise a store key so a write lands where `lookup` will read it.

    Index and role keys pass through; a `name:` key is re-normalised through
    `_norm_name`. Without this, a caller that sends the raw displayed text --
    `name:B Plate` instead of `name:b plate` -- writes a key nothing ever reads:
    the save succeeds, the response echoes the new name, and the row still shows
    the old one. Normalising on the way IN means the stored key always matches
    what `name_key` produces, and re-saving repairs an old bad key.
    """
    k = str(key or "")
    if k.startswith(NAME_PREFIX):
        return name_key(k[len(NAME_PREFIX):])
    return k


def set_overrides(job_dir: Path, updates: Dict[str, Optional[str]]) -> Dict[str, str]:
    """Merge `updates` into the store. A blank or None value CLEARS an override.

    Returns the full override map after the merge.
    """
    if not job_dir.exists():
        return {}
    meta = _read_meta(job_dir)
    current = dict(meta.get(META_KEY) or {})

    for k, v in (updates or {}).items():
        key = _canonical_key(str(k))
        name = clean_name(v) if v is not None else ""
        if name:
            current[key] = name
        else:
            current.pop(key, None)
            # A caller that sent an un-normalised name key would otherwise leave
            # the real entry behind and appear to have cleared nothing.
            current.pop(str(k), None)

    meta[META_KEY] = current
    import time

    meta["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    _meta_path(job_dir).write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return get_overrides(job_dir)


ROLE_PREFIX = "role:"
NAME_PREFIX = "name:"

# Roles that identify nothing in particular. `steel_plate` is the catch-all
# sheet_pricing assigns to any workbook row whose name it cannot map, so keying a
# rename on it would rename every unmapped steel row at once.
GENERIC_ROLES = {"", "steel_plate", "purchased_component", "other", "ignore"}


def role_key(role: str) -> str:
    """Store key for a role-wide rename."""
    return f"{ROLE_PREFIX}{(role or '').strip()}"


def name_key(name: str) -> str:
    """Store key for a rename identified by the name currently displayed.

    The last resort, and the reason renaming can no longer fail silently: every
    row has a name even when it has no CAD index and no meaningful role.
    """
    return f"{NAME_PREFIX}{_norm_name(name)}"


def _norm_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def lookup(index, role: str, overrides: Dict[str, str], name: str = "") -> Optional[str]:
    """Override for a row, most specific first.

    index  -> one CAD solid.          A job with two B plates renames them apart.
    role   -> every row of that role. A steel row can rename "the B plate".
    name   -> whatever is displayed.  Covers rows with neither of the above.
    """
    if not overrides:
        return None
    hit = overrides.get(str(index))
    if hit:
        return hit
    if role and role not in GENERIC_ROLES:
        hit = overrides.get(role_key(role))
        if hit:
            return hit
    if name:
        return overrides.get(name_key(name))
    return None


def apply_to_rows(rows: list, overrides: Dict[str, str]) -> list:
    """Stamp overrides onto part rows from get_job().

    Sets `role_label` (what every panel displays) and adds:
      * `name_override`      the user's text, or absent
      * `role_label_original` what the role would have been called

    `role` is untouched. See the module docstring.
    """
    if not overrides:
        return rows
    for r in rows:
        name = lookup(
            r.get("index", ""),
            r.get("role", ""),
            overrides,
            r.get("role_label") or r.get("Component") or "",
        )
        if not name:
            continue
        r["role_label_original"] = r.get("role_label", "")
        r["role_label"] = name
        r["name_override"] = name
    return rows


def display_for(index, role_label_value: str, overrides: Dict[str, str]) -> str:
    """Override for this CAD index, else the label passed in."""
    if not overrides:
        return role_label_value
    return overrides.get(str(index), role_label_value)


def sheet_name_aliases(job_dir: Path, role_of_index: Dict[str, str]) -> Dict[str, list]:
    """Extra name keywords per role, so a renamed plate is still found in the .xls.

    `read_sheet_dimensions()` locates a plate's stock T/W/L by searching the
    workbook for a known keyword ("a plate", "support plate", ...). Once
    `sheet_rename.py` has written a custom name into the sheet, that search no
    longer matches and the dimensions silently disappear -- which changes the
    price, from a rename that was supposed to be cosmetic.

    So the override map is fed back in as additional aliases for the role it
    belongs to, closing the loop.

    `role_of_index` maps CAD index -> role, which the caller has and this module
    does not.
    """
    extra: Dict[str, list] = {}
    for key, name in get_overrides(job_dir).items():
        if key.startswith(ROLE_PREFIX):
            role = key[len(ROLE_PREFIX):]
        else:
            role = role_of_index.get(str(key), "")
        if not role:
            continue
        extra.setdefault(role, []).append(name.strip().lower())
    return extra
