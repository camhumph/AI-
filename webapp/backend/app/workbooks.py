"""Which .xls in a job folder is THIS job's quote sheet, and which is its steel sheet.

WHY THIS EXISTS AT ALL
    Three modules needed the answer and each grew its own version:
    sheet_pricing scored candidates, sheet_rename and sheet_grades took the
    alphabetically-first filename containing "STEEL SHEET". All three then picked
    the wrong file on C18500, in two different ways, for the same reason.

    C18500's folder contains both:

        J000-STEEL SHEET-std.xls               <- the blank template, or another
                                                  job's leftover copy of it
        Majestic-5475-C18500 STEEL SHEET.xls   <- what the macro wrote for THIS job

    sheet_pricing scored the template 120 against the real sheet's 80, because it
    gave +40 for "J000" and nothing at all for matching the job. sheet_rename and
    sheet_grades took the first name in sort order, and "J000" sorts before
    "Majestic". So:

      * the Parts & Pricing tab listed the template's stale rows -- TCP at 11.875 x
        9.875 in #7, a plate literally named "0", and a row of 23s from the
        template's Machining Sheet -- next to the five real plates, and
      * renaming a plate or changing its steel wrote into the template. Nothing
        appeared to happen, because nothing happened to the job's own workbook.

THE RULE
    A workbook whose filename carries this job's identity is this job's workbook.
    Everything else -- "J000", "-std", "template" -- is a weak hint that only
    matters when no job-named file exists, which is the case mid-run, before the
    macro has written its copy.
"""
from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Dict, Iterable, List, Optional

# Names that mark a file as a blank template or another job's copy rather than
# this job's output. Only ever used to BREAK A TIE, never to beat a job match.
TEMPLATE_HINTS = ("J000", "-STD", "_STD", "TEMPLATE", "BLANK")


def job_identities(job_dir: Path) -> List[str]:
    """Upper-case strings that identify this job in a filename.

    The folder name is the C-number ("C18500"). The macro also writes its exports
    as "<Customer>-<PartNo>-<CNumber>.<ext>", so the stem of any .easm/.dxf/.stl
    beside the workbooks gives the fuller prefix ("MAJESTIC-5475-C18500").
    """
    job_dir = Path(job_dir)
    out: List[str] = []

    folder = job_dir.name.strip().upper()
    if folder:
        out.append(folder)
        # A folder named "CMS_ACTIVE_QUOTE_<name>_<timestamp>" carries the job name
        # in the middle; and a bare C-number may appear with or without a hyphen.
        m = re.search(r"C-?\d{4,6}", folder)
        if m:
            out.append(m.group(0).replace("-", ""))

    for suffix in (".easm", ".dxf", ".stl", ".x_t", ".igs"):
        for p in job_dir.glob(f"*{suffix}"):
            stem = p.stem.strip().upper()
            # Drop the "_2" the macro adds to a second export of the same job.
            stem = re.sub(r"_\d+$", "", stem)
            if stem and stem not in out:
                out.append(stem)

    return out


def _identity_bonus(name_upper: str, identities: Iterable[str]) -> int:
    """How strongly this filename claims to belong to this job."""
    best = 0
    for ident in identities:
        if len(ident) < 4:
            continue
        if ident in name_upper:
            # A longer match is a more specific claim: "MAJESTIC-5475-C18500" beats
            # a bare "C18500", which in turn beats nothing.
            best = max(best, 200 + len(ident))
    return best


def _template_penalty(name_upper: str) -> int:
    for hint in TEMPLATE_HINTS:
        if hint in name_upper:
            return 60
    return 0


def _candidates_in(base: Path) -> List[Path]:
    """Every workbook in ONE folder, in sort order."""
    out: List[Path] = []
    try:
        if not base.is_dir():
            return out
        entries = sorted(base.iterdir())
    except OSError:
        # An offline network share must not take a rename down with it.
        return out
    for p in entries:
        try:
            if not p.is_file():
                continue
        except OSError:
            continue
        if p.suffix.lower() not in (".xls", ".xlsx", ".xlsm"):
            continue
        if p.name.lower().startswith("~$"):
            continue
        out.append(p)
    return out


def _candidates(job_dir: Path) -> List[Path]:
    """Workbooks to CHOOSE BETWEEN when one answer is wanted.

    Deduplicated by filename, so `documents/X.xls` is dropped when `X.xls` sits
    at the job root: for a READ they hold the same thing and picking either is
    the same answer. A WRITE is the opposite -- see `search_bases`.
    """
    job_dir = Path(job_dir)
    seen = set()
    out: List[Path] = []
    for sub in ("", "documents"):
        for p in _candidates_in(job_dir / sub if sub else job_dir):
            key = p.name.lower()
            if key in seen:
                continue
            seen.add(key)
            out.append(p)
    return out


def score_steel(path: Path, identities: Iterable[str]) -> int:
    nm = path.name.upper()
    if "PURCHASED" in nm:
        return -1000
    score = 0
    if "STEEL" in nm and "SHEET" in nm:
        score += 80
    if "MACHINING" in nm:
        score += 20
    # Quote_Steel_Grinding also contains "STEEL"; it is the other workbook.
    if "QUOTE" in nm and "GRIND" in nm:
        score -= 130
    if score <= 0:
        return score
    score += _identity_bonus(nm, identities)
    score -= _template_penalty(nm)
    # "J000" still helps, but only against another non-job file.
    if "J000" in nm:
        score += 10
    return score


def score_quote(path: Path, identities: Iterable[str]) -> int:
    nm = path.name.upper()
    if "PURCHASED" in nm:
        return -1000
    score = 0
    if "QUOTE" in nm and "STEEL" in nm:
        score += 80
    if "GRIND" in nm:
        score += 30
    if "QUOTE" in nm:
        score += 20
    if score <= 0:
        return score
    score += _identity_bonus(nm, identities)
    score -= _template_penalty(nm)
    return score


def find(job_dir: Path) -> Dict[str, Optional[Path]]:
    """This job's quote and steel workbooks, or None where there is no candidate."""
    job_dir = Path(job_dir)
    idents = job_identities(job_dir)
    cands = _candidates(job_dir)

    best_q: Optional[Path] = None
    best_qs = 0
    best_s: Optional[Path] = None
    best_ss = 0

    for p in cands:
        q = score_quote(p, idents)
        if q > best_qs:
            best_qs, best_q = q, p
        s = score_steel(p, idents)
        if s > best_ss:
            best_ss, best_s = s, p

    # The same file cannot be both. Quote wins the tie because its scorer is the
    # more specific of the two ("QUOTE" + "GRIND" together).
    if best_q is not None and best_s is not None and best_q == best_s:
        best_s, best_ss = None, 0
        for p in cands:
            if p == best_q:
                continue
            s = score_steel(p, idents)
            if s > best_ss:
                best_ss, best_s = s, p

    return {"quote": best_q, "steel": best_s}


# --------------------------------------------------------------------------
# EVERY COPY, for writing.
#
# THIS IS THE BUG THAT MADE RENAMING LOOK LIKE IT DID NOTHING.
#
# `find()` returns ONE quote sheet and ONE steel sheet, which is right for
# reading a price and wrong for writing a name. A finished job has the same
# workbook sitting in up to four places:
#
#     cms_data\jobs\C18328\BMS-...STEEL SHEET.xls              the registry copy
#     cms_data\jobs\C18328\documents\BMS-...STEEL SHEET.xls    what the Docs tab serves
#     C:\CMS_Local_Workspace\C18328\BMS-...STEEL SHEET.xls     what the macro wrote
#     C:\CMS_Local_Workspace\C18328\documents\...              if the macro made one
#
# jobs.import_from_folder copies the macro's output into the registry twice, at
# the root and under documents/, and records where it came from in meta.json as
# `source_folder`. sheet_rename wrote to whichever ONE `find()` returned -- the
# registry root -- and reported, truthfully, "updated 3 cell(s)". The estimator
# then opened the macro's folder, or downloaded the file from the Docs tab, and
# saw the old name. Same for a steel-grade change.
#
# So a write has to reach all of them. The alternative -- rewriting one and
# re-copying over the others -- was rejected because the copies are not always
# identical: the macro can rewrite its own folder between a sync and a rename,
# and a blind copy would silently throw that away.
# --------------------------------------------------------------------------
def source_folder(job_dir: Path) -> Optional[Path]:
    """The folder the macro wrote this job into, from meta.json. None if unknown."""
    try:
        raw = json.loads((Path(job_dir) / "meta.json").read_text(encoding="utf-8"))
    except Exception:
        return None
    val = str((raw or {}).get("source_folder") or "").strip()
    if not val:
        return None
    try:
        p = Path(val)
        return p if p.is_dir() else None
    except OSError:
        return None


def search_bases(job_dir: Path) -> List[Path]:
    """Every folder that can hold a copy of this job's workbooks, best first.

    The registry comes first so a partial failure still leaves the app's own view
    correct, and the source folder is included last because it is the one that
    may be a slow or offline share.
    """
    job_dir = Path(job_dir)
    bases = [job_dir, job_dir / "documents"]

    src = source_folder(job_dir)
    if src:
        bases += [src, src / "documents"]

    out: List[Path] = []
    seen = set()
    for b in bases:
        try:
            if not b.is_dir():
                continue
            key = str(b.resolve()).lower()
        except OSError:
            continue
        if key in seen:
            continue
        seen.add(key)
        out.append(b)
    return out


def find_all(job_dir: Path) -> Dict[str, List[Path]]:
    """{"quote": [...], "steel": [...]} -- one best pick PER FOLDER, deduped by path.

    Scored per folder rather than globally: `documents/` and the macro's folder
    each hold their own copy of the same job's workbook, and both need the edit.
    Two folders that turn out to hold the very same file (a junction, or a source
    folder that IS the registry folder) collapse to one entry.
    """
    job_dir = Path(job_dir)
    idents = job_identities(job_dir)
    out: Dict[str, List[Path]] = {"quote": [], "steel": []}
    seen: Dict[str, set] = {"quote": set(), "steel": set()}

    for base in search_bases(job_dir):
        cands = _candidates_in(base)
        if not cands:
            continue
        # Identities are also collected from the base itself: the macro's folder
        # holds the .x_t / .easm exports that name the job, while the registry
        # folder may only have the C-number.
        local_idents = list(idents)
        for extra in job_identities(base):
            if extra not in local_idents:
                local_idents.append(extra)

        best: Dict[str, Optional[Path]] = {"quote": None, "steel": None}
        best_score = {"quote": 0, "steel": 0}
        for p in cands:
            q = score_quote(p, local_idents)
            if q > best_score["quote"]:
                best_score["quote"], best["quote"] = q, p
            s = score_steel(p, local_idents)
            if s > best_score["steel"]:
                best_score["steel"], best["steel"] = s, p

        # Same file cannot be both, same tie-break as find().
        if best["quote"] is not None and best["quote"] == best["steel"]:
            best["steel"], best_score["steel"] = None, 0
            for p in cands:
                if p == best["quote"]:
                    continue
                s = score_steel(p, local_idents)
                if s > best_score["steel"]:
                    best_score["steel"], best["steel"] = s, p

        for kind, path in best.items():
            if path is None:
                continue
            try:
                key = str(path.resolve()).lower()
            except OSError:
                key = str(path).lower()
            if key in seen[kind]:
                continue
            seen[kind].add(key)
            out[kind].append(path)

    return out


def explain(job_dir: Path) -> List[Dict[str, object]]:
    """Per-candidate scores, for diagnosing a wrong pick."""
    idents = job_identities(job_dir)
    return [
        {
            "name": p.name,
            "quote": score_quote(p, idents),
            "steel": score_steel(p, idents),
        }
        for p in _candidates(job_dir)
    ]
