"""Write plate renames into the quote sheet and steel sheet.

WHY THIS IS NOT JUST openpyxl
    Both workbooks are legacy BIFF8 `.xls` (Composite Document File V2), not
    `.xlsx`. openpyxl cannot open or write that format at all, and xlrd 2.x is
    read-only. The xlwt/xlutils route can write .xls but drops formulas and
    formatting -- and the quote sheet is formula-driven, so rewriting it that way
    would hand the shop a workbook with the arithmetic stripped out.

    So this drives Excel itself over COM, exactly as Module6121 already does
    (`CreateObject("Excel.Application")`). Same mechanism, same fidelity, no new
    file format risk.

EVERY COPY, NOT JUST THE ONE THE APP READS
    A finished job has the same two workbooks in up to four folders: the
    registry, the registry's documents/ (what the Docs tab serves), and the
    macro's own output folder recorded in meta.json as `source_folder`. This
    used to write to whichever single file `workbooks.find()` returned and
    report, truthfully, "updated 3 cell(s)" -- while the file the estimator
    actually opens still said the old name. See workbooks.find_all().

WHAT IT TOUCHES
    Two cells per renamed plate, matching where the macro writes them:

      * Quote sheet, `QuoteWorksheet`, COLUMN A of the plate's row
        (Module6121: `xlWs.Cells(targetRow, 1).value = stdName(i)`)
      * Steel sheet, `Steel Order` and `Machining Sheet`, COLUMN B, rows 19+
        (Module6121: `ws.Cells(writeRow, 2).value = Replace(stdName(i), Chr(34), "")`)

    Rows are located by MATCHING THE EXISTING TEXT, not by row number. The
    macro's row numbers come from a grade-dependent lookup
    (`StdQuoteRowFor(slot, grade)`) that this module does not have access to, and
    guessing a row number would overwrite an unrelated plate. Matching the text
    that is already in the cell is both safer and self-correcting: if the sheet
    has already been renamed, the new name is found and updated again.

GRACEFUL DEGRADATION
    On a machine without pywin32 or without Excel, nothing is written and the
    caller is told why. The rename is still saved and still flows everywhere
    else, including into the bridge file -- so the next macro run will put it in
    the sheets anyway. Failing to rewrite the .xls today is a delay, not a loss.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

from . import workbooks

# Sheets and columns the macro writes plate names into, 1-based like Excel.
QUOTE_SHEET_NAME = "QuoteWorksheet"
QUOTE_NAME_COL = 1  # column A
STEEL_SHEET_NAMES = ("Steel Order", "Machining Sheet")
STEEL_NAME_COL = 2  # column B
STEEL_FIRST_ROW = 19
MAX_SCAN_ROW = 260


def _norm(s: Any) -> str:
    """Compare names the way a human would: case, spacing and quotes ignored."""
    if s is None:
        return ""
    t = str(s).replace('"', " ").replace("'", " ")
    t = re.sub(r"[^A-Za-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip().lower()


def _find_workbooks(job_dir: Path) -> Dict[str, List[Path]]:
    """EVERY copy of this job's two workbooks. See app/workbooks.py.

    This used to take the first filename in sort order containing "STEEL SHEET".
    On C18500 that was "J000-STEEL SHEET-std.xls" -- a stray copy of the blank
    template, which sorts before "Majestic-5475-C18500 STEEL SHEET.xls" because J
    precedes M. Renames were written into the template, so the job's own steel
    sheet never changed and the rename looked like it had done nothing.

    It then took the ONE right file and left the job's other three copies of it
    stale, which looked the same from the estimator's chair.
    """
    return workbooks.find_all(job_dir)


# QuoteWorksheet data columns: C=qty, D=thickness, E=width, F=length.
_QUOTE_DATA_COLS = (3, 4, 5, 6)


def _quote_row_has_data(ws, row: int) -> bool:
    """True when this QuoteWorksheet row is a filled plate, not a blank heading."""
    for col in _QUOTE_DATA_COLS:
        v = ws.Cells(row, col).Value
        if v is None:
            continue
        if isinstance(v, str):
            if v.strip():
                return True
            continue
        try:
            if float(v) != 0.0:
                return True
        except Exception:
            return True
    return False


def _excel_available() -> Tuple[bool, str]:
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except Exception:
        return False, (
            "pywin32 is not installed, so the .xls workbooks cannot be edited "
            "from Python. The rename is saved and will be written into the "
            "sheets the next time the macro runs. To rewrite them now: "
            "pip install pywin32"
        )
    return True, ""


def rewrite_names(
    job_dir: Path,
    renames: List[Tuple[str, str]],
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Replace old plate names with new ones in both workbooks.

    `renames` is [(old_name, new_name), ...]. Old names are matched loosely, so
    'A Plate', 'a plate' and '"A" Plate' all hit.

    Returns a report; never raises. A failure here must not lose the rename.
    """
    report: Dict[str, Any] = {
        "attempted": len(renames),
        "changed": 0,
        "cells": [],
        "workbooks": {},
        "files_changed": [],
        "skipped_reason": "",
        "ok": False,
    }
    if not renames:
        report["skipped_reason"] = "nothing to rename"
        return report

    books = _find_workbooks(job_dir)
    report["workbooks"] = {k: [str(p) for p in v] for k, v in books.items()}
    if not any(books.values()):
        report["skipped_reason"] = (
            "No quote sheet or steel sheet found for this job yet. Run the macro "
            "first; the rename is saved and will be applied then."
        )
        return report

    ok, why = _excel_available()
    if not ok:
        report["skipped_reason"] = why
        return report

    want = {_norm(old): new for old, new in renames if _norm(old) and new}
    if not want:
        report["skipped_reason"] = "no usable old names to match"
        return report

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        for kind, paths in books.items():
            for path in paths:
                report["changed"] += _rewrite_one(
                    excel, kind, path, want, report, dry_run
                )

        report["ok"] = True
        if report["changed"] == 0 and not report.get("errors"):
            report["skipped_reason"] = (
                "The old names were not found in the workbooks. They may have "
                "been renamed already, or the macro has not written these rows."
            )
    except Exception as e:
        report["skipped_reason"] = f"Excel automation failed: {e}"
    finally:
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    return report


def _rewrite_one(excel, kind: str, path: Path, want: Dict[str, str], report, dry_run) -> int:
    """Apply every wanted rename to ONE workbook file. Never raises.

    Isolated per file so one unreachable copy -- a share that dropped, a
    workbook someone has open -- is reported against that copy and the other
    copies still get the edit.
    """
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path))

        # A workbook someone has open in Excel opens READ-ONLY, and Save() on it
        # does nothing useful with DisplayAlerts off. Say so instead of counting
        # cells that will never reach the file.
        if not dry_run and _is_read_only(wb):
            report.setdefault("errors", []).append(
                f"{path.name}: opened read-only, so nothing could be written. "
                "It is probably open in Excel on this PC — close it and rename again."
            )
            return 0

        touched = 0
        targets = (
            [(QUOTE_SHEET_NAME, QUOTE_NAME_COL, 1)]
            if kind == "quote"
            else [(n, STEEL_NAME_COL, STEEL_FIRST_ROW) for n in STEEL_SHEET_NAMES]
        )

        for sheet_name, col, first_row in targets:
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                continue  # sheet not in this workbook, fine
            for row in range(first_row, MAX_SCAN_ROW + 1):
                cur = ws.Cells(row, col).Value
                key = _norm(cur)
                if not key or key not in want:
                    continue
                new = want[key]
                if _norm(new) == key:
                    continue  # already says what we want

                # DO NOT RENAME AN EMPTY BLOCK LABEL.
                #
                # QuoteWorksheet is pre-labelled by the template, so the text
                # "A Plate" appears twice: once as row 8's empty '"A" Plate'
                # heading in the #1 A-36 block, and once as this job's filled
                # row 27. Renaming the heading rewrites the workbook's own
                # scaffolding and leaves a block that no longer says which
                # plate belongs in it.
                if kind == "quote" and not _quote_row_has_data(ws, row):
                    continue
                report["cells"].append(
                    {
                        "workbook": path.name,
                        "folder": str(path.parent),
                        "sheet": sheet_name,
                        "cell": f"{chr(64 + col)}{row}",
                        "from": str(cur),
                        "to": new,
                    }
                )
                if not dry_run:
                    ws.Cells(row, col).Value = new
                touched += 1

        if touched and not dry_run:
            wb.Save()
            report["files_changed"].append(str(path))
        return touched
    except Exception as e:
        report.setdefault("errors", []).append(f"{path.name}: {e}")
        return 0
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass


def _is_read_only(wb) -> bool:
    try:
        return bool(wb.ReadOnly)
    except Exception:
        return False
