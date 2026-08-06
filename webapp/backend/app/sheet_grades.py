"""Write steel-grade changes into the quote sheet and the steel sheet.

WHY THIS IS SEPARATE FROM sheet_rename.py
    A rename is a text swap: find the cell that says "A Plate", make it say
    something else. Column, sheet and row all stay put.

    A grade change moves the plate BETWEEN BLOCKS of the quote workbook. Each
    grade owns its own band of rows -- #1 A-36 at 6-15, #2 4140 at 22-34, #3 P20
    at 38-43, #7 420-SS at 52-61, ALM 6061 at 68-77 -- and the row a plate sits in
    is what prices it, because the per-pound rate and the formulas live in the
    block. So switching a plate from #2 to #3 means clearing its row in the 4140
    band and writing name/qty/T/W/L into the matching row of the P20 band.

    Two different jobs, so two modules. Same Excel-over-COM mechanism and the same
    "never raise, always report" contract as sheet_rename.

MIRRORS THE MACRO, DELIBERATELY
    QUOTE_ROWS and SPARE_ROWS below are transcriptions of StdQuoteRowFor and
    NextStdSpareQuoteRow in Module6121.bas, and STEEL_TYPE_LABEL of
    StdSteelTypeFor. If those change in the macro they must change here too, or a
    plate the estimator regrades lands in a different row than the same plate
    would after the next macro run.

WHAT IT TOUCHES
    * Steel sheet, `Steel Order` and `Machining Sheet`, rows 19+:
        column B = plate name (used to FIND the row, never written here)
        column H = steel type text, e.g. "#3 P20"
    * Quote sheet, `QuoteWorksheet`:
        column A = plate name, C = qty, D = thickness, E = width, F = length
      The whole 5-cell group moves from the old row to the new one.

SLOT COMES FROM THE ROLE, NOT THE LABEL
    The row within a block is chosen by SLOT (TOPCLAMP, A, B, RAILS, ...). A plate
    the estimator has renamed to "Cavity Insert Backer" no longer says anything a
    name-based slot test can read, so the slot is derived from the plate's ROLE,
    which a rename never changes. The name is used only to locate the current row.
"""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

from . import workbooks

QUOTE_SHEET_NAME = "QuoteWorksheet"
QUOTE_NAME_COL = 1  # A
QUOTE_QTY_COL = 3  # C
QUOTE_T_COL = 4  # D
QUOTE_W_COL = 5  # E
QUOTE_L_COL = 6  # F
# Read, never written. G is the weight formula and H the extended steel price;
# both belong to the row, so moving a plate into another grade's block reprices
# it at that block's per-pound rate. Reported back so the estimator can see the
# money the change moved. See sheet_pricing._load_steel_from_quote_block, which
# reads the same column.
QUOTE_WEIGHT_COL = 7  # G
QUOTE_PRICE_COL = 8  # H

STEEL_SHEET_NAMES = ("Steel Order", "Machining Sheet")
STEEL_NAME_COL = 2  # B
STEEL_TYPE_COL = 8  # H
STEEL_FIRST_ROW = 19

MAX_SCAN_ROW = 260

# StdSteelTypeFor
STEEL_TYPE_LABEL = {
    "A36": "#1 A-36",
    "4140": "#2 4140",
    "P20": "#3 P20",
    "420SS": "#7 420-SS",
    "6061": "ALM 6061",
    "H13": "#5 H13",
    "A2": "#1 A-36",
    "O1": "#1 A-36",
}

# StdQuoteRowFor. Grades with no block of their own (A2, O1) price in the A-36
# band, exactly as the macro's Case "A2", "O1" does.
_A36_ROWS = {
    "TOPCLAMP": 6, "MANIFOLD": 7, "A": 8, "B": 9, "SUPPORT": 10,
    "BOTTOMCLAMP": 11, "RAILS": 12, "PIN": 14, "EJECTOR": 15,
}
QUOTE_ROWS: Dict[str, Dict[str, int]] = {
    "A36": _A36_ROWS,
    "A2": _A36_ROWS,
    "O1": _A36_ROWS,
    "4140": {
        "TOPCLAMP": 22, "BOTTOMCLAMP": 23, "RAILS": 24, "PIN": 25, "SUPPORT": 26,
        "A": 27, "B": 28, "MANIFOLD": 29, "EJECTOR": 30, "X": 31, "Y": 32,
    },
    "P20": {"TOPCLAMP": 38, "A": 39, "B": 40, "X": 41, "Y": 42, "SUPPORT": 43},
    "420SS": {
        "TOPCLAMP": 52, "MANIFOLD": 53, "A": 54, "B": 55, "Y": 56, "SUPPORT": 57,
        "RAILS": 58, "BOTTOMCLAMP": 59, "PIN": 60, "EJECTOR": 61,
    },
    "6061": {
        "TOPCLAMP": 68, "A": 69, "B": 70, "X": 71, "Y": 72, "SUPPORT": 73,
        "RAILS": 74, "BOTTOMCLAMP": 75, "PIN": 76, "EJECTOR": 77,
    },
}

# NextStdSpareQuoteRow. Plain rows before premium ones in every block -- see
# THICK_ROWS below for why the 4140 order changed.
_A36_SPARE = [13, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 17, 18]
SPARE_ROWS: Dict[str, List[int]] = {
    "A36": _A36_SPARE,
    "A2": _A36_SPARE,
    "O1": _A36_SPARE,
    "4140": [24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34],
    "P20": [41, 42, 43, 38, 39, 40, 46, 47, 48],
    "420SS": [52, 53, 54, 55, 56, 57, 58, 59, 60, 61, 62, 63, 64],
    "6061": [68, 69, 70, 71, 72, 73, 74, 75, 76, 77],
}

# --------------------------------------------------------------------------
# THICK-PLATE PRICE BAND -- the yellow rows on QuoteWorksheet.
#
# Transcribed from StdThickMinFor / StdThickRowsFor in Module6121.bas, and they
# must stay identical: the macro chooses the row on the first pass and this
# module chooses it again whenever the estimator changes a plate's steel. If the
# two disagree, changing a grade and changing it back would land the plate in a
# different price band than the macro put it in.
#
# The shop's rule: premium rows take plates 5.875" thick and up, except in the
# #1 A-36 block where the cut is 2.00". Inclusive at the boundary. Applies to
# pot and holder blocks the same as any other plate.
#
#     block      normal $/lb   premium rows   premium $/lb
#     A36           1.38        16, 17, 18       1.58
#     4140          1.88        30 - 34          2.18
#     P20           2.45        46, 47, 48       3.75
#     420SS      3.85/3.95      62, 63, 64       3.95
#     6061          3.50        (none)             --
# --------------------------------------------------------------------------
_A36_THICK = [16, 17, 18]
THICK_ROWS: Dict[str, List[int]] = {
    "A36": _A36_THICK,
    "A2": _A36_THICK,
    "O1": _A36_THICK,
    "4140": [30, 31, 32, 33, 34],
    "P20": [46, 47, 48],
    "420SS": [62, 63, 64],
    "6061": [],
}
_A36_THICK_MIN = 2.0
THICK_MIN: Dict[str, float] = {
    "A36": _A36_THICK_MIN,
    "A2": _A36_THICK_MIN,
    "O1": _A36_THICK_MIN,
    "4140": 5.875,
    "P20": 5.875,
    "420SS": 5.875,
    "6061": 0.0,
}
# Half a thou, so a plate the shop calls 5.875 still counts when geometry
# measured it as 5.874999. Matches THICK_EPS in Module6121.bas.
THICK_EPS = 0.0005


def is_thick_plate(grade: str, thickness) -> bool:
    """Does this plate belong in the premium rows of its grade block?"""
    mn = THICK_MIN.get((grade or "").upper(), 0.0)
    if mn <= 0:
        return False
    try:
        t = float(thickness)
    except (TypeError, ValueError):
        return False
    return t >= mn - THICK_EPS

# Every row any block can use, so "is this row part of a grade block" is answerable.
ALL_BLOCK_ROWS = set()
for _g, _m in QUOTE_ROWS.items():
    ALL_BLOCK_ROWS.update(_m.values())
for _g, _l in SPARE_ROWS.items():
    ALL_BLOCK_ROWS.update(_l)

# Role -> slot. Roles are the stable identifiers from app/roles.py; the macro's
# StdSlotForName reaches the same slots from plate names.
ROLE_SLOT = {
    "top_clamp_plate": "TOPCLAMP",
    "bottom_clamp_plate": "BOTTOMCLAMP",
    # BMS / pot-block roles. Without these a grade change on a pot job wrote the
    # steel type onto the steel sheet and then silently gave up on the quote
    # sheet: slot_for() returned "" for role "tcp" (ROLE_SLOT had only
    # "top_clamp_plate") and the name "TCP" contains no " TOP CLAMP ", so
    # _NAME_SLOT_RULES missed it too. C18328's TCP switched to #1 A-36 on the
    # steel order and stayed in the #2 4140 block at $256.62.
    "tcp": "TOPCLAMP",
    "bcp": "BOTTOMCLAMP",
    # The four pot roles have no slot in any grade block -- the macro writes them
    # into the 4140 block's spare rows. Mapped to "" on purpose, so slot_for()
    # returns a known-empty answer and _do_quote takes the spare-row path instead
    # of guessing a named slot that would land on another plate's row.
    "id_holder": "",
    "od_holder": "",
    "id_pot": "",
    "od_pot": "",
    "a_plate": "A",
    "b_plate": "B",
    "x_plate": "X",
    "y_plate": "Y",
    "cavity_plate": "A",
    "core_plate": "B",
    "support_plate": "SUPPORT",
    "manifold_plate": "MANIFOLD",
    "rail": "RAILS",
    "rail_1": "RAILS",
    "rail_2": "RAILS",
    "riser": "RAILS",
    "pin_plate": "PIN",
    "ejector_plate": "EJECTOR",
    "ejector_retainer_plate": "PIN",
    "bottom_ejector_plate": "PIN",
    "ejector_backup_plate": "PIN",
    "stripper_plate": "B",
    "die_plate": "X",
    "die_backup_plate": "SUPPORT",
    "sc_retainer_plate": "B",
    "sc_backup_plate": "SUPPORT",
}

# Fallback when the role is generic ("steel_plate"): read the label the way
# StdSlotForName reads a plate name. Order matters exactly as it does there.
_NAME_SLOT_RULES: List[Tuple[str, str]] = [
    ("RUNNER STRIPPER", "X"),
    ("STRIPPER", "B"),
    ("TOP CLAMP", "TOPCLAMP"),
    ("BOTTOM CLAMP", "BOTTOMCLAMP"),
    # The abbreviations the BMS steel sheet actually carries. Matched as whole
    # tokens like every other rule here, so they cannot fire inside a longer word.
    ("TCP", "TOPCLAMP"),
    ("BCP", "BOTTOMCLAMP"),
    ("A PLATE", "A"),
    ("CAVITY", "A"),
    ("B PLATE", "B"),
    ("CORE", "B"),
    ("SC RETAINER", "B"),
    ("SC BACKUP", "SUPPORT"),
    ("SC BACK UP", "SUPPORT"),
    ("EJ BACKUP", "PIN"),
    ("EJECTOR BACKUP", "PIN"),
    ("EJECTOR BACK UP", "PIN"),
    ("EJ RET", "EJECTOR"),
    ("BOTTOM EJECTOR", "PIN"),
    ("EJECTOR RETAINER", "PIN"),
    ("PIN PLATE", "PIN"),
    ("RETAINER", "PIN"),
    ("EJECTOR", "EJECTOR"),
    ("SUPPORT", "SUPPORT"),
    ("PILLAR", "SUPPORT"),
    # Both singular and plural, as StdSlotForName does. The macro writes the row
    # "Rails", which does not contain " RAIL " once padded.
    ("RAIL", "RAILS"),
    ("RAILS", "RAILS"),
    ("RISER", "RAILS"),
    ("RISERS", "RAILS"),
    ("DIE BACKUP", "SUPPORT"),
    ("DIE BACK UP", "SUPPORT"),
    ("DIE", "X"),
    ("BACKING", "SUPPORT"),
    ("MANIFOLD", "MANIFOLD"),
    ("X PLATE", "X"),
    ("Y PLATE", "Y"),
]


def _norm(s: Any) -> str:
    """Compare names the way sheet_rename does, so both modules find the same row."""
    if s is None:
        return ""
    t = str(s).replace('"', " ").replace("'", " ")
    t = re.sub(r"[^A-Za-z0-9]+", " ", t)
    return re.sub(r"\s+", " ", t).strip().lower()


def slot_for(role: str, name: str = "") -> str:
    """Quote-block slot for a plate. Role first, label only as a fallback."""
    r = (role or "").strip().lower()
    if r in ROLE_SLOT:
        return ROLE_SLOT[r]
    upper = " " + re.sub(r"[^A-Za-z0-9]+", " ", str(name or "").upper()).strip() + " "
    for token, slot in _NAME_SLOT_RULES:
        if f" {token} " in upper:
            return slot
    return ""


def target_row(grade: str, slot: str) -> int:
    return QUOTE_ROWS.get(grade, {}).get(slot, 0)


def _find_workbooks(job_dir: Path) -> Dict[str, List[Path]]:
    """EVERY copy of this job's two workbooks -- see app/workbooks.py.

    All of them, not just the one the app reads: the registry copy, the
    documents/ copy the Docs tab serves, and the macro's own output folder. A
    grade change that reached only the first looked like it had done nothing.
    """
    return workbooks.find_all(job_dir)


def _excel_available() -> Tuple[bool, str]:
    try:
        import win32com.client  # type: ignore  # noqa: F401
    except Exception:
        return False, (
            "pywin32 is not installed, so the .xls workbooks cannot be edited from "
            "Python. The grade change is saved and will be written into the sheets "
            "the next time the macro runs. To rewrite them now: pip install pywin32"
        )
    return True, ""


def rewrite_grades(
    job_dir: Path,
    changes: List[Dict[str, str]],
    dry_run: bool = False,
) -> Dict[str, Any]:
    """Apply grade changes to both workbooks.

    `changes` is [{"name": <text currently in the sheet>,
                   "role": <stable role id>,
                   "grade": <new grade code>}, ...]

    Returns a report and never raises: a failure here must not lose the override.
    """
    report: Dict[str, Any] = {
        "attempted": len(changes),
        "changed": 0,
        "cells": [],
        "moves": [],
        "workbooks": {},
        "files_changed": [],
        "skipped_reason": "",
        "unresolved": [],
        "ok": False,
    }
    if not changes:
        report["skipped_reason"] = "nothing to regrade"
        return report

    # Index the wanted changes by normalised name, carrying the slot with each.
    #
    # A BLANK SLOT IS NOT A FAILURE. Four of the six BMS roles -- ID/OD Holder,
    # ID/OD Pot -- have no named slot in any grade block, because the macro puts
    # them in the 4140 block's spare rows. This used to report them unresolved
    # and skip the quote sheet entirely; _do_quote now takes the spare-row path,
    # and only a block with no free row is a real failure.
    want: Dict[str, Dict[str, str]] = {}
    for c in changes:
        nm = _norm(c.get("name"))
        grade = (c.get("grade") or "").strip().upper()
        if not nm or grade not in QUOTE_ROWS:
            report["unresolved"].append(
                {"name": c.get("name"), "grade": c.get("grade"), "why": "no name or unknown grade"}
            )
            continue
        slot = slot_for(c.get("role") or "", c.get("name") or "")
        want[nm] = {"grade": grade, "slot": slot, "display": str(c.get("name") or "")}

    if not want:
        report["skipped_reason"] = "no usable changes"
        return report

    books = _find_workbooks(job_dir)
    report["workbooks"] = {k: [str(p) for p in v] for k, v in books.items()}
    if not any(books.values()):
        report["skipped_reason"] = (
            "No quote sheet or steel sheet found for this job yet. Run the macro "
            "first; the grade is saved and will be applied then."
        )
        return report

    ok, why = _excel_available()
    if not ok:
        report["skipped_reason"] = why
        return report

    import pythoncom  # type: ignore
    import win32com.client  # type: ignore

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        for path in books.get("steel") or []:
            report["changed"] += _do_steel(excel, path, want, report, dry_run)
        for path in books.get("quote") or []:
            report["changed"] += _do_quote(excel, path, want, report, dry_run)

        report["ok"] = True
        if report["changed"] == 0 and not report.get("errors"):
            report["skipped_reason"] = (
                "The plate names were not found in the workbooks. They may have been "
                "renamed since, or the macro has not written these rows yet."
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


def _do_steel(excel, path: Path, want, report, dry_run) -> int:
    """Steel sheet: rewrite the steel-type text beside each renamed plate."""
    touched = 0
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path))
        if not dry_run and _is_read_only(wb):
            report.setdefault("errors", []).append(
                f"{path.name}: opened read-only, so nothing could be written. "
                "It is probably open in Excel on this PC — close it and try again."
            )
            return 0
        for sheet_name in STEEL_SHEET_NAMES:
            try:
                ws = wb.Worksheets(sheet_name)
            except Exception:
                continue
            for row in range(STEEL_FIRST_ROW, MAX_SCAN_ROW + 1):
                key = _norm(ws.Cells(row, STEEL_NAME_COL).Value)
                if not key or key not in want:
                    continue
                label = STEEL_TYPE_LABEL.get(want[key]["grade"], "")
                if not label:
                    continue
                cur = ws.Cells(row, STEEL_TYPE_COL).Value
                if _norm(cur) == _norm(label):
                    continue
                report["cells"].append(
                    {
                        "workbook": path.name,
                        "folder": str(path.parent),
                        "sheet": sheet_name,
                        "cell": f"H{row}",
                        "plate": want[key]["display"],
                        "from": "" if cur is None else str(cur),
                        "to": label,
                    }
                )
                if not dry_run:
                    ws.Cells(row, STEEL_TYPE_COL).Value = label
                touched += 1
        if touched and not dry_run:
            wb.Save()
            report["files_changed"].append(str(path))
    except Exception as e:
        report.setdefault("errors", []).append(f"{path.name}: {e}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    return touched


def _do_quote(excel, path: Path, want, report, dry_run) -> int:
    """Quote sheet: move each plate's 5-cell row group into its new grade block."""
    touched = 0
    wb = None
    try:
        wb = excel.Workbooks.Open(str(path))
        if not dry_run and _is_read_only(wb):
            report.setdefault("errors", []).append(
                f"{path.name}: opened read-only, so nothing could be written. "
                "It is probably open in Excel on this PC — close it and try again."
            )
            return 0
        try:
            ws = wb.Worksheets(QUOTE_SHEET_NAME)
        except Exception:
            report.setdefault("errors", []).append(
                f"{path.name}: no '{QUOTE_SHEET_NAME}' sheet"
            )
            return 0

        # Where every named plate currently sits, and which block rows are taken.
        # Read once: a COM round trip per cell is slow and the sheet is small.
        #
        # A FILLED ROW BEATS AN EMPTY LABEL WITH THE SAME NAME.
        #
        # Every grade block is pre-labelled by the template, so QuoteWorksheet holds
        # the text "A Plate" more than once: row 8 is the empty '"A" Plate' label in
        # the #1 A-36 block, row 27 is this job's actual A plate with its qty and
        # sizes. Taking the first match took row 8, moved an empty row into the P20
        # block, and left the real plate where it was -- so changing the steel
        # appeared to do nothing at all.
        current: Dict[str, int] = {}
        occupied: Dict[int, str] = {}
        pending: List[Dict[str, Any]] = []  # moves whose price is read after recalc
        for row in range(1, MAX_SCAN_ROW + 1):
            key = _norm(ws.Cells(row, QUOTE_NAME_COL).Value)
            if not key:
                continue
            has_data = _row_has_data(ws, row)
            if row in ALL_BLOCK_ROWS and has_data:
                occupied[row] = key
            if key in want:
                if key not in current:
                    current[key] = row
                elif has_data and not _row_has_data(ws, current[key]):
                    current[key] = row  # upgrade from a bare label to the real row

        for key, spec in want.items():
            slot = spec["slot"]
            grade = spec["grade"]
            src = current.get(key, 0)

            if not src:
                report["unresolved"].append(
                    {
                        "name": spec["display"],
                        "grade": grade,
                        "why": "not found in the quote sheet, so there is no row to move",
                    }
                )
                continue

            # A plate with no named slot still moves. ID/OD Holder and ID/OD Pot
            # have no slot in any block -- the macro parks them in the 4140
            # block's spare rows -- and skipping them here is what left C18328's
            # pot plates priced in a grade the estimator had changed away from.
            dst = target_row(grade, slot) if slot else 0
            if not dst or (occupied.get(dst) and occupied.get(dst) != key):
                dst = _first_free(SPARE_ROWS.get(grade, []), occupied, key)

            # THICKNESS DECIDES THE PRICE BAND, NOT THE SLOT.
            #
            # The slot map says which plate this is; it says nothing about what
            # the steel costs. It also sends 4140 EJECTOR to row 30 and X/Y to
            # 31/32, which ARE the premium rows -- so without this a thin ejector
            # plate reprices at 2.18/lb the moment its grade is touched.
            thickness = _cell_number(ws, src, QUOTE_T_COL)
            before_band = dst
            dst, band_note = _band_correct(grade, thickness, dst, occupied, key)

            # A BAND CORRECTION THAT COULD NOT HAPPEN IS NEWS.
            #
            # When every premium row is taken the plate stays where it is, which
            # is the right call -- a plate priced low still beats a plate that
            # could not be placed. But the move record below is only written when
            # the row actually changes, so without this the estimator is told
            # nothing at all and the plate quietly prices in the wrong band.
            #
            # It is a real capacity limit, not a rare edge: the 4140 block has
            # five premium rows, and a pot base with four thick pots plus a thick
            # TCP and BCP needs six.
            if band_note and dst == before_band:
                report["unresolved"].append(
                    {"name": spec["display"], "grade": grade, "why": band_note}
                )

            if not dst:
                where = f"slot {slot}" if slot else "an unslotted plate"
                report["unresolved"].append(
                    {
                        "name": spec["display"],
                        "grade": grade,
                        "why": f"the {grade} block has no free row for {where}. Its steel "
                        "type is written on the steel sheet, but the quote row could "
                        "not be moved and the price is still the old grade's.",
                    }
                )
                continue
            if dst == src:
                continue  # already in the right block

            vals = {
                QUOTE_NAME_COL: ws.Cells(src, QUOTE_NAME_COL).Value,
                QUOTE_QTY_COL: ws.Cells(src, QUOTE_QTY_COL).Value,
                QUOTE_T_COL: ws.Cells(src, QUOTE_T_COL).Value,
                QUOTE_W_COL: ws.Cells(src, QUOTE_W_COL).Value,
                QUOTE_L_COL: ws.Cells(src, QUOTE_L_COL).Value,
            }
            price_before = _cell_number(ws, src, QUOTE_PRICE_COL)

            move = {
                "workbook": path.name,
                "folder": str(path.parent),
                "sheet": QUOTE_SHEET_NAME,
                "plate": spec["display"],
                "moved_from_row": src,
                "moved_to_row": dst,
                "grade": grade,
                "grade_label": STEEL_TYPE_LABEL.get(grade, grade),
                "slot": slot or "spare",
                "thickness": thickness,
                "band": "premium" if dst in (THICK_ROWS.get(grade) or []) else "standard",
                "band_note": band_note,
                "price_before": price_before,
                "price_after": None,
            }
            report["cells"].append(move)
            report["moves"].append(move)

            if not dry_run:
                # Write the destination first, then clear the source. If this fails
                # halfway the plate is duplicated, which is visible; clearing first
                # and failing would delete it outright.
                for col, v in vals.items():
                    ws.Cells(dst, col).Value = v
                for col in (QUOTE_NAME_COL, QUOTE_QTY_COL, QUOTE_T_COL, QUOTE_W_COL, QUOTE_L_COL):
                    ws.Cells(src, col).ClearContents()
                pending.append(move)

            occupied.pop(src, None)
            occupied[dst] = key
            touched += 1

        if touched and not dry_run:
            # RECALCULATE BEFORE SAVING, ALWAYS.
            #
            # Price is a formula in column H of the destination row, and the app
            # reads this .xls back with xlrd, which can only see the value Excel
            # last CACHED in the file. A workbook opened with calculation set to
            # manual -- which the template can carry, and which Module6121 also
            # sets while it fills the sheet -- saves the OLD cached price next to
            # the new grade. The number on the Parts tab would then be the price
            # of the block the plate just left.
            _recalculate(excel, wb)
            for move in pending:
                move["price_after"] = _cell_number(ws, move["moved_to_row"], QUOTE_PRICE_COL)
            wb.Save()
            report["files_changed"].append(str(path))
    except Exception as e:
        report.setdefault("errors", []).append(f"{path.name}: {e}")
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
    return touched


def _row_has_data(ws, row: int) -> bool:
    """Does this QuoteWorksheet row hold a real plate, or is it a blank label?

    Qty or any of the three sizes present means the macro filled it in. The
    template's own block labels have all four empty.
    """
    for col in (QUOTE_QTY_COL, QUOTE_T_COL, QUOTE_W_COL, QUOTE_L_COL):
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


def _cell_number(ws, row: int, col: int):
    """A cell's numeric value, or None when it is blank or text."""
    try:
        v = ws.Cells(row, col).Value
    except Exception:
        return None
    if v is None or isinstance(v, str):
        return None
    try:
        return round(float(v), 2)
    except (TypeError, ValueError):
        return None


def _recalculate(excel, wb) -> None:
    """Force formulas to recompute, whatever the workbook's calculation mode.

    xlCalculationAutomatic is   -4105 . Set on the application because that is
    where Excel keeps the mode, and restored is not worth the risk: this is a
    private DispatchEx instance that is quit at the end of the call.
    """
    try:
        excel.Calculation = -4105
    except Exception:
        pass
    for attempt in (
        lambda: wb.Application.CalculateFullRebuild(),
        lambda: wb.Application.CalculateFull(),
        lambda: excel.Calculate(),
    ):
        try:
            attempt()
            return
        except Exception:
            continue


def _is_read_only(wb) -> bool:
    try:
        return bool(wb.ReadOnly)
    except Exception:
        return False


def _band_correct(
    grade: str, thickness, dst: int, occupied: Dict[int, str], key: str
) -> Tuple[int, str]:
    """Move `dst` into the price band this plate's thickness calls for.

    Returns (row, note). The row is unchanged whenever it is already in the
    right band, which is the common case, and unchanged when there is nowhere
    to move to -- a plate priced in the wrong band still beats a plate that
    could not be placed at all, and the note says which happened.

    Mirrors StdBandedQuoteRow in Module6121.bas.
    """
    rows = THICK_ROWS.get(grade) or []
    if not rows or not dst or thickness is None:
        return dst, ""

    want_thick = is_thick_plate(grade, thickness)
    if (dst in rows) == want_thick:
        return dst, ""

    mn = THICK_MIN.get(grade, 0.0)
    if want_thick:
        alt = _first_free(rows, occupied, key)
        if not alt:
            return dst, (
                f'{thickness}" is {mn}" or over, but every premium row in the {grade} '
                "block is taken, so it stays on the standard rate."
            )
        return alt, f'{thickness}" thick — priced on the {grade} premium row {alt}.'

    plain = [r for r in SPARE_ROWS.get(grade, []) if r not in rows]
    alt = _first_free(plain, occupied, key)
    if not alt:
        return dst, (
            f'{thickness}" is under {mn}", but no standard {grade} row is free, so it '
            "stays on the premium rate."
        )
    return alt, f'{thickness}" is under {mn}" — moved off the premium row to {alt}.'


def _first_free(rows: List[int], occupied: Dict[int, str], key: str) -> int:
    for r in rows:
        holder = occupied.get(r)
        if not holder or holder == key:
            return r
    return 0
