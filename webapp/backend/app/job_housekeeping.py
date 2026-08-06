"""Tidy a finished job folder: loose CAD into `base\\`, diagnostics into PDF.

WHY THE CSVs MOVE, AND WHY THAT IS SAFE NOW
    Two of these CSVs are live inputs, not reports:

      * XT_Export_CAD_Dimensions.csv is the webapp's source of truth for a job --
        jobs.py reads it for mass, for `has_raw_csv`, and for the whole parts list.
      * XT_Export_BOM_Match_Report.csv is read by sheet_pricing.py.

    So this module used to render them to PDF and leave the CSVs in the job root,
    on the grounds that moving them would take the job down. The danger was real
    but the conclusion was not: the readers now resolve paths through
    find_job_file(), which looks in the job root, then `pdf\\`, then
    `documents\\`. sheet_pricing.py and bms_pairs.py already did this for
    root + documents; it is now one helper used everywhere.

    With that in place, each CSV is rendered to a PDF and then MOVED into
    `pdf\\` beside it. The root ends up holding only deliverables, and every
    reader still finds its input. If you add a new reader of a generated CSV, go
    through find_job_file() -- `job_dir / name` will work until the first tidy.

WHY THE CAD MOVE IS REVERSIBLE
    `base\\<job>.SLDASM` is the macro's exported assembly and it very likely holds
    external references to the imported part files. Those references cannot be read
    out of the binary from here, so whether SolidWorks re-resolves them after the
    move cannot be proved without opening it.

    Moving parts INTO the assembly's own folder is the safe direction -- SolidWorks
    searches the assembly folder when a stored path misses -- but "likely fine" is
    not "verified". Every move is therefore recorded in a manifest and `undo()` puts
    them back exactly, so opening the assembly once is a check with a cheap escape
    rather than a commitment.
"""
from __future__ import annotations

import csv
import json
import re
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional

MANIFEST_NAME = "_cms_tidy_manifest.json"
# Named "pdf" because that is what the shop calls it -- the folder of PDFs.
DIAGNOSTICS_DIR = "pdf"

# Where a job's generated CSVs can be found, in priority order.
#
# The macro always WRITES to the job root. tidy_job() then moves the diagnostic
# CSVs into pdf\, and the email/upload paths copy attachments into documents\.
# So a reader that only looks in the root breaks the moment the folder is tidied
# -- which is why find_job_file() exists and why every reader of these files goes
# through it. sheet_pricing.py and bms_pairs.py already searched root +
# documents\ for exactly this reason; this generalises that to one place.
JOB_FILE_SEARCH_DIRS = ("", DIAGNOSTICS_DIR, "documents")


def find_job_file(job_dir, name: str) -> Optional[Path]:
    """First existing copy of `name` for this job, or None.

    Searches the job root, then pdf\\, then documents\\. Use this for any
    generated CSV instead of `job_dir / name`, so tidying a folder cannot take
    the webapp down with it.
    """
    base = Path(job_dir)
    for sub in JOB_FILE_SEARCH_DIRS:
        candidate = (base / sub / name) if sub else (base / name)
        if candidate.is_file():
            return candidate
    return None


def job_file_exists(job_dir, name: str) -> bool:
    return find_job_file(job_dir, name) is not None

# Imported CAD that belongs beside the exported assembly, not in the job root.
CAD_MOVE_SUFFIXES = {".sldprt", ".sldasm"}
# Neutral CAD is only source material when it is NOT one of the job's own exports.
NEUTRAL_SUFFIXES = {".x_t", ".xt", ".step", ".stp", ".igs", ".iges", ".sldlfp"}

# Diagnostic CSVs the macro writes, in the order they are useful to read.
DIAGNOSTIC_CSVS = (
    "Standard_Quote_Rows_Debug.csv",
    "Standard_Quote_Rows_Debug_BEFORE_STL.csv",
    "XT_Export_CAD_Dimensions.csv",
    "CAD_All_Components_Debug_FINAL.csv",
    "CAD_All_Components_Debug_PRE_ORIENT.csv",
    "Stack_LeaderPin_Analysis.csv",
    "XT_Export_BOM_Match_Report.csv",
    "PDF_Knowledge_Evidence.csv",
    "Job_File_Inventory.csv",
    # Safe to render whoever reads it, because this pass copies and never moves.
    "Purchased Components Quote.csv",
)


# ----------------------------------------------------------------------------
# Loose CAD -> base\
# ----------------------------------------------------------------------------
def _job_base_name(job_dir: Path) -> str:
    """The job's own export prefix, e.g. "ITWMedical-125-C18597".

    A job folder can hold TWO sets of .easm/.dxf/.stl: the customer's originals and
    the macro's exports. C18517 has both "8617-MB.EASM" (customer) and
    "Cameo-8617-C18517.easm" (ours). Taking whichever the filesystem listed first
    picked the customer's, and then "Cameo-8617-C18517.x_t" -- a deliverable -- did
    not match the prefix and was scheduled to be moved into base\\.
    """
    job_dir = Path(job_dir)
    cnum = re.search(r"C-?\d{4,6}", job_dir.name.upper())
    cands: List[str] = []
    for suffix in (".easm", ".dxf", ".stl"):
        for p in job_dir.glob(f"*{suffix}"):
            cands.append(p.stem)
    # The macro names its exports after the job, so the C-number is the tell.
    if cnum:
        for s in cands:
            if cnum.group(0).replace("-", "") in s.upper().replace("-", ""):
                return s
    if cands:
        return max(cands, key=len)
    return job_dir.name


def _deliverable_identities(job_dir: Path) -> List[str]:
    """Lower-case stems that mark a file as this job's own output, not source CAD."""
    job_dir = Path(job_dir)
    out = {_job_base_name(job_dir).lower(), job_dir.name.lower()}
    m = re.search(r"C-?\d{4,6}", job_dir.name.upper())
    if m:
        out.add(m.group(0).lower())
        out.add(m.group(0).replace("-", "").lower())
    return [s for s in out if s]


def plan_cad_move(job_dir: Path) -> List[Dict[str, str]]:
    """Loose CAD files in the job root that belong in base\\."""
    job_dir = Path(job_dir)
    base = job_dir / "base"
    idents = _deliverable_identities(job_dir)

    moves: List[Dict[str, str]] = []
    for p in sorted(job_dir.iterdir()):
        if not p.is_file():
            continue
        suf = p.suffix.lower()
        stem_l = p.stem.lower()

        if suf in CAD_MOVE_SUFFIXES:
            pass  # always source CAD
        elif suf in NEUTRAL_SUFFIXES:
            # A neutral file carrying this job's identity IS a deliverable. Matched
            # anywhere in the stem, not just as a prefix, so "Cameo-8617-C18517.x_t"
            # and "Cameo-8617-C18517_2.x_t" both stay put.
            if any(ident in stem_l for ident in idents):
                continue
        else:
            continue

        moves.append({"from": str(p), "to": str(base / p.name), "name": p.name})
    return moves


def tidy_cad_files(job_dir: Path, dry_run: bool = False) -> Dict[str, Any]:
    """Move loose CAD into base\\, recording every move so it can be undone."""
    job_dir = Path(job_dir)
    base = job_dir / "base"
    report: Dict[str, Any] = {
        "moved": [],
        "skipped": [],
        "errors": [],
        "manifest": None,
        "dry_run": dry_run,
    }

    moves = plan_cad_move(job_dir)
    if not moves:
        report["skipped"].append("no loose CAD files in the job root")
        return report

    if not dry_run:
        base.mkdir(parents=True, exist_ok=True)

    done: List[Dict[str, str]] = []
    for m in moves:
        src, dst = Path(m["from"]), Path(m["to"])
        if dst.exists():
            # Never overwrite: a same-named file already beside the assembly is
            # more likely the one the assembly is using.
            report["skipped"].append(f"{m['name']} (already in base\\)")
            continue
        if dry_run:
            done.append(m)
            continue
        try:
            shutil.move(str(src), str(dst))
            done.append(m)
        except Exception as e:
            report["errors"].append(f"{m['name']}: {e}")

    report["moved"] = [m["name"] for m in done]

    if done and not dry_run:
        mpath = base / MANIFEST_NAME
        existing = []
        if mpath.exists():
            try:
                existing = json.loads(mpath.read_text(encoding="utf-8")).get("moves", [])
            except Exception:
                existing = []
        mpath.write_text(
            json.dumps({"moves": existing + done}, indent=2), encoding="utf-8"
        )
        report["manifest"] = str(mpath)

    return report


def undo_cad_move(job_dir: Path) -> Dict[str, Any]:
    """Put every recorded move back where it came from."""
    job_dir = Path(job_dir)
    mpath = job_dir / "base" / MANIFEST_NAME
    report: Dict[str, Any] = {"restored": [], "skipped": [], "errors": []}
    if not mpath.exists():
        report["skipped"].append("no tidy manifest; nothing to undo")
        return report
    try:
        moves = json.loads(mpath.read_text(encoding="utf-8")).get("moves", [])
    except Exception as e:
        report["errors"].append(f"manifest unreadable: {e}")
        return report

    remaining = []
    for m in moves:
        src, dst = Path(m["to"]), Path(m["from"])
        if not src.exists():
            report["skipped"].append(f"{m['name']} (not in base\\ any more)")
            continue
        if dst.exists():
            report["skipped"].append(f"{m['name']} (already back in the job root)")
            continue
        try:
            shutil.move(str(src), str(dst))
            report["restored"].append(m["name"])
        except Exception as e:
            report["errors"].append(f"{m['name']}: {e}")
            remaining.append(m)

    if remaining:
        mpath.write_text(json.dumps({"moves": remaining}, indent=2), encoding="utf-8")
    else:
        try:
            mpath.unlink()
        except Exception:
            pass
    return report


# ----------------------------------------------------------------------------
# Diagnostic CSV -> PDF
# ----------------------------------------------------------------------------
def _read_csv_rows(path: Path) -> List[List[str]]:
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            with path.open("r", encoding=enc, newline="") as fh:
                return [row for row in csv.reader(fh)]
        except UnicodeDecodeError:
            continue
        except Exception:
            return []
    return []


def csv_to_pdf(
    csv_path: Path,
    pdf_path: Path,
    title: str = "",
    max_cols_per_page: int = 12,
) -> Dict[str, Any]:
    """Render one CSV as a landscape PDF table.

    Wide tables are split into column groups across pages, with the first column
    repeated on each group so a row stays identifiable. That beats shrinking a
    24-column dimension dump until it cannot be read.
    """
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import landscape, letter
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        LongTable,
        PageBreak,
        Paragraph,
        SimpleDocTemplate,
        Spacer,
    )

    rows = _read_csv_rows(csv_path)
    if not rows:
        return {"ok": False, "why": "empty or unreadable", "csv": csv_path.name}

    # Some of these files hold two stacked tables separated by a blank line
    # (Stack_LeaderPin_Analysis is SECTION/KEY/VALUE then a per-part table).
    blocks: List[List[List[str]]] = []
    current: List[List[str]] = []
    for r in rows:
        if not any(str(c).strip() for c in r):
            if current:
                blocks.append(current)
                current = []
            continue
        current.append([("" if c is None else str(c)) for c in r])
    if current:
        blocks.append(current)

    styles = {
        "title": ParagraphStyle("t", fontName="Helvetica-Bold", fontSize=13, leading=16),
        "sub": ParagraphStyle("s", fontName="Helvetica", fontSize=8, leading=10,
                              textColor=colors.HexColor("#555555")),
        "cell": ParagraphStyle("c", fontName="Helvetica", fontSize=6, leading=7.2),
        "head": ParagraphStyle("h", fontName="Helvetica-Bold", fontSize=6, leading=7.2,
                               textColor=colors.white),
    }

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(letter),
        leftMargin=0.35 * inch, rightMargin=0.35 * inch,
        topMargin=0.4 * inch, bottomMargin=0.4 * inch,
        title=title or csv_path.stem,
    )
    avail = doc.width

    flow: List[Any] = [
        Paragraph(title or csv_path.stem, styles["title"]),
        Paragraph(f"Source: {csv_path.name}", styles["sub"]),
        Spacer(1, 8),
    ]

    first_block = True
    for block in blocks:
        header, body = block[0], block[1:]
        ncol = max(len(r) for r in block)
        header = header + [""] * (ncol - len(header))
        body = [r + [""] * (ncol - len(r)) for r in body]

        groups: List[List[int]] = []
        if ncol <= max_cols_per_page:
            groups = [list(range(ncol))]
        else:
            rest = list(range(1, ncol))
            step = max_cols_per_page - 1
            for i in range(0, len(rest), step):
                groups.append([0] + rest[i:i + step])

        for gi, cols in enumerate(groups):
            if not first_block or gi:
                flow.append(PageBreak())
            first_block = False

            if len(groups) > 1:
                flow.append(
                    Paragraph(
                        f"Columns {cols[1]}–{cols[-1]} of {ncol - 1} "
                        f"(part {gi + 1} of {len(groups)}); first column repeated.",
                        styles["sub"],
                    )
                )
                flow.append(Spacer(1, 4))

            data = [[Paragraph(header[c], styles["head"]) for c in cols]]
            for r in body:
                data.append([Paragraph(r[c].replace("&", "&amp;")
                                       .replace("<", "&lt;")
                                       .replace(">", "&gt;"), styles["cell"])
                             for c in cols])

            widths = _column_widths([header] + body, cols, avail)
            tbl = LongTable(data, colWidths=widths, repeatRows=1)
            tbl.setStyle(_table_style(colors))
            flow.append(tbl)

    try:
        doc.build(flow)
    except Exception as e:
        return {"ok": False, "why": str(e), "csv": csv_path.name}
    return {"ok": True, "csv": csv_path.name, "pdf": pdf_path.name,
            "rows": sum(len(b) - 1 for b in blocks)}


def _column_widths(rows: List[List[str]], cols: List[int], avail: float) -> List[float]:
    """Width per column from the longest cell, normalised to the page width."""
    weights = []
    for c in cols:
        longest = 1
        for r in rows[:400]:  # sampling is enough and keeps this fast
            if c < len(r):
                longest = max(longest, len(r[c]))
        weights.append(min(max(longest, 4), 34))
    total = float(sum(weights)) or 1.0
    return [avail * w / total for w in weights]


def _table_style(colors):
    from reportlab.platypus import TableStyle
    return TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#33475b")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1),
             [colors.white, colors.HexColor("#f2f5f8")]),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#b8c2cc")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, -1), 1.5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 1.5),
        ]
    )


def diagnostics_to_pdf(
    job_dir: Path,
    dry_run: bool = False,
    only: Optional[List[str]] = None,
) -> Dict[str, Any]:
    """Render every diagnostic CSV present into `diagnostics\\`. CSVs are kept."""
    job_dir = Path(job_dir)
    out_dir = job_dir / DIAGNOSTICS_DIR
    report: Dict[str, Any] = {"written": [], "skipped": [], "errors": [], "dir": str(out_dir)}

    wanted = list(only) if only else list(DIAGNOSTIC_CSVS)
    present = [job_dir / n for n in wanted if (job_dir / n).is_file()]
    if not present:
        report["skipped"].append("no diagnostic CSVs found in this job folder")
        return report

    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    job_label = _job_base_name(job_dir)
    for src in present:
        pdf = out_dir / (src.stem + ".pdf")
        if dry_run:
            report["written"].append(pdf.name)
            continue
        # Plain hyphen, not an em dash: the built-in Helvetica the PDF uses is
        # WinAnsi-encoded and an em dash comes back as a replacement char in some
        # readers. Not worth a font embed for a separator.
        res = csv_to_pdf(src, pdf, title=f"{job_label} - {src.stem.replace('_', ' ')}")
        if res.get("ok"):
            report["written"].append(f"{pdf.name} ({res.get('rows', 0)} rows)")
        else:
            report["errors"].append(f"{src.name}: {res.get('why')}")
    return report


# Run logs the macro leaves in the job root. These really are reports -- nothing
# reads them back from the job folder -- so they move rather than being rendered.
#   * CMS_Base_Export_Log.txt   the macro's own copy lives in DOWNLOADS_FOLDER
#     (Module6121.bas:972), so the job-folder one is a leftover.
#   * CMS_Module6121_Live_Log.txt  quote_pipeline.py:1356 tails the copy in
#     LOCAL_WORKSPACE, not this one, and the macro re-opens it with Append on the
#     next run, so moving a finished job's copy loses nothing.
LOG_TXT_MOVE = (
    "CMS_Base_Export_Log.txt",
    "CMS_Module6121_Live_Log.txt",
)

# Text files that must NOT move, whatever they look like.
#   cms_zip_extract_done.txt is a sentinel the macro tests for at
#   CurrentJobFolder & "\cms_zip_extract_done.txt" (Module6121.bas:2382) to decide
#   whether the job's ZIP is already unpacked. Moved, the next run re-extracts.
LOG_TXT_KEEP = (
    "cms_zip_extract_done.txt",
)


def logs_to_diagnostics(job_dir: Path, dry_run: bool = False) -> Dict[str, Any]:
    """Move macro run logs into `pdf\\`, leaving control files in the root."""
    job_dir = Path(job_dir)
    out_dir = job_dir / DIAGNOSTICS_DIR
    report: Dict[str, Any] = {"moved": [], "kept": [], "errors": [], "dir": str(out_dir)}

    movable = [job_dir / n for n in LOG_TXT_MOVE if (job_dir / n).is_file()]
    for name in LOG_TXT_KEEP:
        if (job_dir / name).is_file():
            report["kept"].append(f"{name} (control file the macro reads)")
    # Any other .txt is unknown: report it rather than move it blind.
    known = set(LOG_TXT_MOVE) | set(LOG_TXT_KEEP)
    for extra in sorted(job_dir.glob("*.txt")):
        if extra.name not in known:
            report["kept"].append(f"{extra.name} (unrecognised - left in place)")

    if not movable:
        return report
    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    for src in movable:
        dest = out_dir / src.name
        if dry_run:
            report["moved"].append(src.name)
            continue
        try:
            if dest.exists():
                dest.unlink()
            shutil.move(str(src), str(dest))
            report["moved"].append(src.name)
        except Exception as e:
            report["errors"].append(f"{src.name}: {type(e).__name__}: {e}")
    return report


def move_diagnostic_csvs(job_dir: Path, dry_run: bool = False) -> Dict[str, Any]:
    """Move the diagnostic CSVs into `pdf\\` after they have been rendered.

    WHY THIS IS NOW A MOVE
        The header of this module used to argue for keeping the CSVs in the job
        root, because jobs.py and sheet_pricing.py read two of them and a move
        would take the job down. That reasoning was right about the danger and
        wrong about the fix: the readers now resolve through find_job_file(),
        which searches root, pdf\\ and documents\\, so the files can live in
        pdf\\ and still be found. Requiring a tidy folder and requiring working
        readers were never actually in conflict.

        Run this AFTER diagnostics_to_pdf, so each CSV has a PDF beside it.
    """
    job_dir = Path(job_dir)
    out_dir = job_dir / DIAGNOSTICS_DIR
    report: Dict[str, Any] = {"moved": [], "skipped": [], "errors": [], "dir": str(out_dir)}

    present = [job_dir / n for n in DIAGNOSTIC_CSVS if (job_dir / n).is_file()]
    # Any other CSV the macro or a classifier left loose in the root.
    known = {n.lower() for n in DIAGNOSTIC_CSVS}
    for extra in sorted(job_dir.glob("*.csv")):
        if extra.name.lower() not in known and extra.is_file():
            present.append(extra)

    if not present:
        report["skipped"].append("no CSVs in the job root")
        return report
    if not dry_run:
        out_dir.mkdir(parents=True, exist_ok=True)

    for src in present:
        dest = out_dir / src.name
        if dry_run:
            report["moved"].append(src.name)
            continue
        try:
            if dest.exists():
                dest.unlink()
            shutil.move(str(src), str(dest))
            report["moved"].append(src.name)
        except Exception as e:
            report["errors"].append(f"{src.name}: {type(e).__name__}: {e}")
    return report


def tidy_job(job_dir: Path, dry_run: bool = False,
             include_cad: bool = False) -> Dict[str, Any]:
    """Tidy one job folder.

    include_cad is off by default: moving imported .SLDPRT/.stp next to the
    exported assembly is the one pass here that SolidWorks might notice, and
    tidy_cad_files keeps an undo manifest precisely because it is unverified.
    The CSV/log passes are safe and are what runs after every quote.
    """
    out: Dict[str, Any] = {"job_dir": str(job_dir)}
    if include_cad:
        out["cad"] = tidy_cad_files(job_dir, dry_run=dry_run)
    # Render first, then move: the PDF is made from the CSV in the root.
    out["diagnostics"] = diagnostics_to_pdf(job_dir, dry_run=dry_run)
    out["csvs"] = move_diagnostic_csvs(job_dir, dry_run=dry_run)
    out["logs"] = logs_to_diagnostics(job_dir, dry_run=dry_run)
    return out
