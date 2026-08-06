"""BOM-driven role hints for the geometry classifier.

WHY THIS EXISTS
---------------
The classifier guesses roles from geometry, and it never sees the one document
that already states them. XT_Export_CAD_Dimensions.csv -- the only input
qwen_classify_xt_csv.py reads -- has these columns:

    Index, Component, Qty, Thickness, Width, Length, BBoxVolume_cuin,
    Mass_or_Vol, CenterX/Y/Z, NThruHoles, NCbore, NCrossAxis, MaxBoreDia,
    HoleSig, NPockets, PocketAreaIn2, MaxPocketDepth, SolidFillPct,
    PocketAreaUp/DnIn2, PocketDepthUp/Dn

Rich geometry, no BOM. Meanwhile the customer's BOM for 25-424 says, in plain
English, one row per plate:

    1000-a00   Top clamp plate, 1.375 x 9-7/8 x 11-7/8          #7 steel
    1200-a00   Stationary retainer plate, 1.375 x 9-7/8 x 11-7/8 #7 steel
    1680-a00   Ejector retainer plate, .50 x 6-7/8 x 11-7/8      #7 steel

That is the answer, written down, and the classifier was inferring it from
bounding boxes instead. Every plate in a mold base is a rectangle; a 1.375
plate and another 1.375 plate at the same footprint are geometrically identical
and only their position and their NAME tell them apart. Geometry has to reason
about stack order to separate a stationary retainer from a movable one. The BOM
just says which is which.

HOW THE JOIN WORKS
------------------
The STEP/CAD component carries the BOM detail number: the BOM row "1200-a00" is
CAD component "25-424--1200-a00". Strip the job prefix and they collide on one
key -- the same match email_ai.reconcile() already does, which paired 29 of 30
BOM lines to CAD parts on C18640 and 9 of 9 plates.

So this is a lookup, not a guess. When a detail number matches, the BOM name is
near-certain evidence and is emitted as HIGH confidence. When it does not match,
this module emits nothing rather than guessing -- geometry is better than a bad
name, and a wrong high-confidence hint is worse than no hint.

WHAT THIS DOES NOT DO
---------------------
It does not re-implement Module6121's naming. The macro's StandardPlateName /
StdSlotForName already map BOM text to quote slots, and on C18640 they got all
nine plates right. This is for the WEBAPP classifier and the AI bridge, which
run outside the macro and never saw the BOM.

It also does not read STL triangles. See STL_NOTE below.

STL_NOTE
--------
An STL is a triangle soup: no units, no material, no names, no topology. Every
feature worth having from one is already in the CSV above, measured by
SolidWorks off real B-rep geometry instead of re-derived from a tessellation:

    volume / bbox fill  ->  BBoxVolume_cuin, Mass_or_Vol, SolidFillPct
    holes               ->  NThruHoles, NCbore, NCrossAxis, MaxBoreDia, HoleSig
    pockets             ->  NPockets, PocketAreaIn2, MaxPocketDepth,
                            PocketAreaUpIn2/DnIn2 (which face they open onto)
    stack position      ->  CenterX/Y/Z

Counting holes from triangle loops would be a worse NThruHoles. So triangles are
not read here: the STLs stay what they are for, which is looking at the part.
"""
from __future__ import annotations

import csv
import re
from pathlib import Path

from . import email_ai

HINTS_FILENAME = "BOM_Role_Hints.csv"

# BOM description -> canonical role key (see roles.py ROLE_LABELS).
#
# Ordered longest-phrase-first and tested in order, because the short tokens are
# substrings of the long ones: "ejector retainer plate" contains "ejector" and
# "retainer", and "stationary retainer plate" contains "retainer". Testing
# "retainer" first would collapse three different plates onto one role. This
# mirrors the ordering discipline in Module6121's StdSlotForName for the same
# reason.
_ROLE_PHRASES = (
    # --- ejector housing, most specific first ---
    ("ejector retainer plate", "ejector_plate"),
    ("ejector back-up plate", "bottom_ejector_plate"),
    ("ejector backup plate", "bottom_ejector_plate"),
    ("ejector back up plate", "bottom_ejector_plate"),
    ("ejector rail", "rail"),
    ("ejector housing", "rail"),
    ("ejector plate", "bottom_ejector_plate"),
    # --- clamps ---
    ("top clamp plate", "top_clamp_plate"),
    ("bottom clamp plate", "bottom_clamp_plate"),
    ("top clamp", "top_clamp_plate"),
    ("bottom clamp", "bottom_clamp_plate"),
    # --- cavity / core ---
    # "Sta." / "Mov." are how C17880's BOM writes it ("Sta. retainer plate.",
    # "Mov. Support plate."), so the abbreviations need their own entries -- the
    # full-word phrases below never matched that job.
    ("sta retainer plate", "a_plate"),
    ("mov retainer plate", "b_plate"),
    ("sta support plate", "support_plate"),
    ("mov support plate", "support_plate"),
    ("stationary retainer plate", "a_plate"),
    ("movable retainer plate", "b_plate"),
    ("moveable retainer plate", "b_plate"),
    ("stationary retainer", "a_plate"),
    ("movable retainer", "b_plate"),
    ("moveable retainer", "b_plate"),
    ("cavity plate", "a_plate"),
    ("core plate", "b_plate"),
    ("a plate", "a_plate"),
    ("b plate", "b_plate"),
    # --- other plates ---
    ("support plate", "support_plate"),
    ("stripper plate", "stripper_plate"),
    ("manifold plate", "manifold_plate"),
    ("sc retainer", "sc_retainer_plate"),
    ("sc backup", "sc_backup_plate"),
    ("pin plate", "pin_plate"),
    ("riser", "rail"),
    ("rail", "rail"),
    # --- purchased hardware ---
    ("ejector guide bushing", "guided_ejector_bushing"),
    ("guided ejector bushing", "guided_ejector_bushing"),
    ("ejector guide pin", "ejector_pin"),
    ("shoulder bushing", "leader_pin_bushing"),
    ("leader pin bushing", "leader_pin_bushing"),
    ("guide bushing", "leader_pin_bushing"),
    ("leader pin", "leader_pin"),
    ("return pin", "return_pin"),
    ("ejector pin", "ejector_pin"),
    ("support pillar", "support_pillar"),
    ("pillar", "support_pillar"),
    ("safety strap", "latch_lock"),
    ("latch lock", "latch_lock"),
    ("tie strap", "latch_lock"),
    ("side lock", "latch_lock"),
    ("stop disc", "purchased_component"),
    ("tubular dowel", "purchased_component"),
    ("dowel pin", "purchased_component"),
    ("sprue bushing", "purchased_component"),
    ("locating ring", "purchased_component"),
)

# A "retainer plate" with no stationary/movable qualifier is ambiguous between
# the A/B side and the ejector stack, so it is deliberately absent above. Say so
# instead of picking one.
_AMBIGUOUS = (
    ("retainer plate", "retainer plate with no stationary/movable/ejector qualifier"),
)


def role_for_description(text: str) -> tuple[str, str]:
    """(role_key, matched_phrase). ("", reason) when nothing is certain."""
    blob = " " + re.sub(r"[^a-z0-9]+", " ", (text or "").lower()).strip() + " "
    for phrase, role in _ROLE_PHRASES:
        if f" {phrase} " in blob:
            return role, phrase
    for phrase, why in _AMBIGUOUS:
        if f" {phrase} " in blob:
            return "", f"ambiguous: {why}"
    return "", "no role phrase matched"


# Spreadsheets the pipeline itself writes into a job folder. None of these is a
# customer BOM, and interpret_bom() is permissive enough to read plate-looking
# rows out of the CAD export -- which it did on the first test, silently
# accepting XT_Export_CAD_Dimensions.csv as the BOM and producing zero hints.
_GENERATED_PATTERNS = (
    "xt_export", "cad_all_components", "standard_quote_rows", "stack_leaderpin",
    "job_file_inventory", "pdf_knowledge", "pcs_naming", "classification",
    "purchased components quote", "bom_role_hints", "pullcore prices",
    "purchased components prices", "steel sheet", "quote steel grinding",
)


def _is_generated(name: str) -> bool:
    low = (name or "").lower()
    return any(pat in low for pat in _GENERATED_PATTERNS)


def _read_cad_index(csv_path: Path) -> list[dict]:
    """Index + Component from the macro's CAD export."""
    out: list[dict] = []
    try:
        with csv_path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
            for row in csv.DictReader(f):
                idx = (row.get("Index") or "").strip()
                comp = (row.get("Component") or "").strip()
                if idx:
                    out.append({"index": idx, "component": comp})
    except Exception:
        return []
    return out


def _cad_leaf(component: str) -> str:
    """Last component of a SolidWorks instance path, instance suffix removed.

    "25-424--bm-quote-1/25-424--1580-a00-1" -> "25-424--1580-a00"
    """
    leaf = (component or "").replace("\\", "/").split("/")[-1].strip()
    leaf = re.sub(r"-\d+$", "", leaf)      # trailing instance number
    return leaf


def build_hints(job_dir, bom_path=None) -> dict:
    """Match BOM rows to CAD indices and derive a role hint for each.

    Returns {"hints": [...], "unmatched_bom": [...], "bom": <name>, "error": ...}.
    Each hint: cad_index, component, detail, bom_name, role, confidence,
    evidence, grade, qty, kind. No dimensions -- see HINT_FIELDS.
    """
    job_dir = Path(job_dir)
    result: dict = {"bom": "", "hints": [], "unmatched_bom": [], "unmatched_cad": [],
                    "error": ""}

    from . import job_housekeeping

    cad_csv = job_housekeeping.find_job_file(job_dir, "XT_Export_CAD_Dimensions.csv")
    if cad_csv is None:
        result["error"] = "No XT_Export_CAD_Dimensions.csv for this job."
        return result
    cad_rows = _read_cad_index(cad_csv)
    if not cad_rows:
        result["error"] = f"{cad_csv.name} has no readable rows."
        return result

    # Find a BOM: the caller's path, else any spreadsheet in the job folder or
    # the shop folder it came from.
    #
    # The shop folder matters. During a macro run the webapp's job dir holds only
    # the copied CAD export, while the customer's BOM is still sitting in
    # C:\CMS_Local_Workspace\<C-number> next to it. meta.json records that path
    # as source_folder, so follow it rather than reporting "no BOM" for the one
    # case -- mid-run classification -- where the hints are most useful.
    candidates = [Path(bom_path)] if bom_path else []
    if not candidates:
        roots = [job_dir]
        try:
            import json as _json

            meta = _json.loads((job_dir / "meta.json").read_text(encoding="utf-8"))
            src = (meta.get("source_folder") or "").strip()
            if src and Path(src).is_dir() and Path(src).resolve() != job_dir.resolve():
                roots.append(Path(src))
        except Exception:
            pass
        for root in roots:
            for sub in ("", "documents", "pdf"):
                base = root / sub if sub else root
                if base.is_dir():
                    candidates += sorted(
                        (p for p in base.glob("*")
                         if p.suffix.lower() in email_ai.BOM_EXTS
                         and p.is_file()
                         and not _is_generated(p.name)),
                        # Real BOMs are workbooks; a .csv in a job folder is far
                        # more likely to be something the macro wrote.
                        key=lambda p: (p.suffix.lower() == ".csv", p.name.lower()),
                    )
    bom = None
    for cand in candidates:
        if not cand.is_file():
            continue
        parsed = email_ai.interpret_bom(email_ai.read_spreadsheet(cand), cand.name)
        if parsed.get("error"):
            continue
        if parsed["plates"] or parsed["purchased"]:
            bom, result["bom"] = parsed, cand.name
            break
    if bom is None:
        result["error"] = "No readable BOM found for this job."
        return result

    # CAD leaf key -> rows (a detail number can be placed more than once).
    cad_by_key: dict[str, list[dict]] = {}
    for row in cad_rows:
        key = email_ai._detail_key(_cad_leaf(row["component"]))
        if key:
            cad_by_key.setdefault(key, []).append(row)

    matched_keys: set[str] = set()
    for kind, item in (
        [("plate", i) for i in bom["plates"]]
        + [("purchased", i) for i in bom["purchased"]]
    ):
        detail = (item.get("detail") or "").strip()
        key = email_ai._detail_key(detail)
        hits = cad_by_key.get(key) or []
        role, phrase = role_for_description(item.get("description", ""))
        if not hits:
            result["unmatched_bom"].append({
                "detail": detail, "description": item.get("description", ""),
                "role": role, "why": "no CAD component carries this detail number",
            })
            continue
        matched_keys.add(key)
        dims = item.get("dims") or {}
        for hit in hits:
            if not role:
                # Matched the part but cannot name its role: record it so the
                # classifier can still use the grade/dims, with no role claim.
                result["hints"].append({
                    "cad_index": hit["index"], "component": hit["component"],
                    "detail": detail, "bom_name": item.get("description", ""),
                    "role": "", "confidence": "NONE", "evidence": phrase,
                    "grade": item.get("grade", ""), "qty": item.get("qty", 0),
                    "kind": kind,
                })
                continue
            result["hints"].append({
                "cad_index": hit["index"], "component": hit["component"],
                "detail": detail, "bom_name": item.get("description", ""),
                "role": role, "confidence": "HIGH",
                "evidence": f"BOM detail {detail} matched CAD; name says '{phrase}'",
                "grade": item.get("grade", ""), "qty": item.get("qty", 0),
                "kind": kind,
            })

    for key, rows in cad_by_key.items():
        if key not in matched_keys:
            for row in rows:
                result["unmatched_cad"].append(
                    {"cad_index": row["index"], "component": row["component"]}
                )
    return result


# NO DIMENSIONS HERE, DELIBERATELY.
#
# The BOM is reference material and its sizes are not always right, so nothing
# downstream should be able to pick a dimension up from this file and treat it as
# authoritative. Sizes come from XT_Export_CAD_Dimensions.csv, measured off the
# model. This file answers "what is this part called and what is it made of";
# "how big is it" has exactly one source, and it is not the BOM.
HINT_FIELDS = ("cad_index", "component", "detail", "bom_name", "role", "confidence",
               "grade", "qty", "kind", "evidence")


def write_hints(job_dir, hints: dict) -> Path | None:
    """Write BOM_Role_Hints.csv beside the CAD export. None when there is nothing."""
    rows = hints.get("hints") or []
    if not rows:
        return None
    job_dir = Path(job_dir)
    path = job_dir / HINTS_FILENAME
    with path.open("w", encoding="utf-8", newline="") as f:
        wr = csv.DictWriter(f, fieldnames=HINT_FIELDS, extrasaction="ignore")
        wr.writeheader()
        for r in rows:
            wr.writerow(r)
    return path


def load_hints(job_dir) -> dict:
    """cad_index -> hint, from a previously written BOM_Role_Hints.csv."""
    from . import job_housekeeping

    path = job_housekeeping.find_job_file(job_dir, HINTS_FILENAME)
    if path is None:
        return {}
    out = {}
    try:
        with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
            for row in csv.DictReader(f):
                idx = (row.get("cad_index") or "").strip()
                if idx:
                    out[idx] = row
    except Exception:
        return {}
    return out


def apply_to_classification(job_dir, classification: dict) -> dict:
    """Overlay BOM role hints onto a geometry classification.

    A matched BOM detail number outranks a geometry guess, so a HIGH hint
    replaces the role. Every replacement is recorded in the row's `reason` and
    counted in job_analysis, because a role that silently changed source is
    exactly the kind of thing that is impossible to debug later.

    Rows the BOM does not cover are left untouched: this adds evidence, it never
    removes any.
    """
    rows = (classification or {}).get("classifications") or []
    hints = load_hints(job_dir)
    if not rows or not hints:
        return classification

    changed = 0
    confirmed = 0
    for row in rows:
        hint = hints.get(str(row.get("index", "")).strip())
        if not hint:
            continue
        role = (hint.get("role") or "").strip()
        if not role:
            continue
        was = (row.get("role") or "").strip()
        detail = hint.get("detail", "")
        name = hint.get("bom_name", "")
        if was == role:
            confirmed += 1
            row["reason"] = f"{row.get('reason', '')} Confirmed by BOM {detail} ('{name}')."
        else:
            changed += 1
            row["role"] = role
            row["confidence"] = "HIGH"
            row["reason"] = (
                f"BOM {detail} names this '{name}', so the role is {role}"
                + (f" (geometry had said {was})" if was else "")
                + f". Source: {hint.get('evidence', 'BOM detail match')}."
            )
        row["bom_detail"] = detail
        row["bom_name"] = name
        if hint.get("grade"):
            row["bom_grade"] = hint["grade"]

    analysis = classification.setdefault("job_analysis", {})
    analysis["bom_role_hints"] = {
        "bom": "",  # filled by the caller that knows the filename
        "matched": len(hints),
        "roles_set_from_bom": changed,
        "roles_confirmed_by_bom": confirmed,
    }
    return classification


def refresh(job_dir, bom_path=None) -> dict:
    """Build and persist the hints for a job. Safe to call repeatedly."""
    hints = build_hints(job_dir, bom_path=bom_path)
    path = write_hints(job_dir, hints)
    hints["written"] = str(path) if path else ""
    named = sum(1 for h in hints.get("hints", []) if h.get("role"))
    hints["summary"] = (
        f"{named} of {len(hints.get('hints', []))} matched CAD part(s) named from "
        f"{hints.get('bom') or 'no BOM'}; "
        f"{len(hints.get('unmatched_bom', []))} BOM line(s) unmatched"
    )
    return hints
