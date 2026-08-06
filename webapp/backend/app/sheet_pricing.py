"""Read quote/steel sheet dimensions and CSV prices for a job.

Prices come only from CSV files (never guessed):
  - Shop: Purchased Components Prices.csv
  - Per-job: Purchased Components Quote.csv (written by Module6121)
  - Per-job: Pullcore Prices.csv

Steel/quote Excel workbooks supply stock/finished sizing for plates.
"""
from __future__ import annotations

import csv
import re
from pathlib import Path

from . import config, plate_grades, workbooks

# Role -> keywords to match Component column in price CSVs
ROLE_CSV_KEYWORDS: dict[str, list[str]] = {
    "top_clamp_plate": ["top clamp", "top plate", "top clamping"],
    "a_plate": ["a plate", "a-plate"],
    "b_plate": ["b plate", "b-plate"],
    "stripper_plate": ["stripper"],
    "sc_retainer_plate": ["sc retainer", "retainer plate"],
    "sc_backup_plate": ["sc backup", "backup plate"],
    "support_plate": ["support plate"],
    "bottom_clamp_plate": ["bottom clamp", "bottom plate", "bot clamp"],
    "rail": ["rail"],
    "pin_plate": ["pin plate"],
    "ejector_plate": ["ejector plate", "ej-ret", "ej ret"],
    "bottom_ejector_plate": ["bottom ejector", "ej-backup", "ej backup", "ejector backup"],
    "latch_lock": ["safety strap", "latch lock", "latch-lock", "lss"],
    "leader_pin": ["leader pin", "ldr-pin", "ldr pin"],
    "leader_pin_bushing": ["bushing", "leader bushing", "guide bushing", "lbb"],
    "guided_ejector_bushing": ["ejector bushing"],
    "return_pin": ["return pin"],
    "ejector_pin": ["ejector pin"],
    "support_pillar": ["support pillar", "pillar"],
    "pullcore": ["pullcore", "pull core"],
}

PLATE_SHEET_NAMES: dict[str, list[str]] = {
    "top_clamp_plate": ["top clamp plate", "top clamping plate"],
    "a_plate": ["a plate"],
    "b_plate": ["b plate"],
    "stripper_plate": ["stripper plate"],
    "sc_retainer_plate": ["sc retainer plate", "sc retainer"],
    "sc_backup_plate": ["sc backup plate", "sc backup"],
    "support_plate": ["support plate"],
    "bottom_clamp_plate": ["bottom clamp plate", "bottom clamping plate"],
    "rail": ["rail"],
    "pin_plate": ["pin plate"],
    "ejector_plate": ["ejector plate"],
    "bottom_ejector_plate": ["bottom ejector plate", "ejector backup plate"],
}


def _safe_float(v, default=0.0) -> float:
    try:
        s = str(v).strip().replace("$", "").replace(",", "")
        return float(s) if s else default
    except (TypeError, ValueError):
        return default


def _norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").lower().strip())


def _read_csv_rows(path: Path) -> list[dict]:
    if not path.exists():
        return []
    rows = []
    try:
        with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            for line in f:
                if line.lstrip().startswith("#"):
                    continue
                rows.append(line)
        if not rows:
            return []
        reader = csv.DictReader(rows)
        return list(reader)
    except Exception:
        return []


def load_shop_prices() -> list[dict]:
    return _read_csv_rows(config.PURCHASED_PRICES_CSV)


def load_job_purchased_quote(job_dir: Path) -> list[dict]:
    for name in (
        "Purchased Components Quote.csv",
        "pdf/Purchased Components Quote.csv",
        "documents/Purchased Components Quote.csv",
    ):
        p = job_dir / name
        rows = _read_csv_rows(p)
        if rows:
            return rows
    return []


def load_job_pullcore(job_dir: Path) -> list[dict]:
    for name in ("Pullcore Prices.csv", "documents/Pullcore Prices.csv"):
        p = job_dir / name
        rows = _read_csv_rows(p)
        if rows:
            return rows
    return []


# BMS QuoteWorksheet #2 4140 block (Module6121 FillQuoteWorkbookFromBoundingBox).
# Col A = name, C = qty, D = thickness, E = width, F = length.
# Template formulas often put hours in G and price in H.
# Rows 22/23/31-34 are the canonical pot-block plates; 24-30 are spare/extra.
# ---------------------------------------------------------------------------
# QuoteWorksheet row maps.
#
# TWO MAPS, BECAUSE THE SAME ROWS MEAN DIFFERENT PLATES.
#
# The workbook has one grade block per material and the macro writes a plate into
# the row for its (slot, grade) pair -- see StdQuoteRowFor in Module6121.bas. A
# BMS pot base and a standard PCS base use DIFFERENT blocks, and rows 31-34
# collide outright: in the #2/4140 block they are the ID/OD Holder and Pot slots,
# while a standard base puts its X and Y plates there. Merging the maps would put
# "X Plate" on a BMS holder or "ID Holder" on a standard X plate, so the base type
# has to choose.
#
# Standard rows, from StdQuoteRowFor:
#   A-36 / #1 block  6 TopClamp  7 Manifold  8 A  9 B  10 Support
#                    11 BottomClamp  12 Rails  14 Pin/BottomEjector  15 Ejector
#   P20              38 TopClamp  39 A  40 B  41 X  42 Y  43 Support
#   420SS            52 TopClamp 53 Manifold 54 A 55 B 56 Y 57 Support
#                    58 Rails 59 BottomClamp 60 Pin 61 Ejector
#   6061             68 TopClamp 69 A 70 B 71 X 73 Support 74 Rails
#                    75 BottomClamp 76 Pin 77 Ejector
#
# WHY THIS EXISTS: C18626 is a standard base whose plates landed on rows 6, 10,
# 11, 12, 14, 15, 39 and 40. Only the BMS map was consulted, so every row was
# skipped and the Parts tab showed blank hours and no price -- while the workbook
# held 62.83 hrs / $86.70 for the Top Clamp Plate and $139.93 for the A Plate.
# The macro had done the work; nothing read it.
_STD_QUOTE_STEEL_ROWS = {
    # A-36 / #1 block
    6: ("Top Clamp Plate", "top_clamp_plate"),
    7: ("Manifold Plate", "manifold_plate"),
    8: ("A Plate", "a_plate"),
    9: ("B Plate", "b_plate"),
    10: ("Support Plate", "support_plate"),
    11: ("Bottom Clamp Plate", "bottom_clamp_plate"),
    12: ("Rails", "rail"),
    14: ("Bottom Ejector Plate", "bottom_ejector_plate"),
    15: ("Ejector Plate", "ejector_plate"),
    # 4140 / #2 -- WHERE A STANDARD BASE ACTUALLY LANDS.
    #
    # This block was missing entirely, and it is the one the macro fills. Both
    # hand quotes we have put every plate of a standard base in #2, so
    # STD_DEFAULT_GRADE_ALL in Module6121.bas now defaults there -- and this map
    # still stopped at row 15 and resumed at 38. Nothing read rows 22-34, so on a
    # standard job EVERY steel plate came back at $0.00 and the Parts tab printed
    # "sheet" in the price column. C18500's five plates are worth $755.61 in the
    # workbook and the app showed nothing for any of them.
    #
    # Slots transcribed from StdQuoteRowFor's 4140 band, same source as
    # sheet_grades.QUOTE_ROWS["4140"]. Note 31/32 are "X"/"Y" here and ID/OD
    # Holder in the BMS map -- the two maps cannot be merged, which is why there
    # are two.
    22: ("Top Clamp Plate", "top_clamp_plate"),
    23: ("Bottom Clamp Plate", "bottom_clamp_plate"),
    24: ("Rails", "rail"),
    25: ("Bottom Ejector Plate", "bottom_ejector_plate"),
    26: ("Support Plate", "support_plate"),
    27: ("A Plate", "a_plate"),
    28: ("B Plate", "b_plate"),
    29: ("Manifold Plate", "manifold_plate"),
    30: ("Ejector Plate", "ejector_plate"),
    31: ('"X" Plate', "stripper_plate"),
    32: ('"Y" Plate', "sc_retainer_plate"),
    # P20
    38: ("Top Clamp Plate", "top_clamp_plate"),
    39: ("A Plate", "a_plate"),
    40: ("B Plate", "b_plate"),
    41: ('"X" Plate', "stripper_plate"),
    42: ('"Y" Plate', "sc_retainer_plate"),
    43: ("Support Plate", "support_plate"),
    # 420SS
    52: ("Top Clamp Plate", "top_clamp_plate"),
    53: ("Manifold Plate", "manifold_plate"),
    54: ("A Plate", "a_plate"),
    55: ("B Plate", "b_plate"),
    56: ('"Y" Plate', "sc_retainer_plate"),
    57: ("Support Plate", "support_plate"),
    58: ("Rails", "rail"),
    59: ("Bottom Clamp Plate", "bottom_clamp_plate"),
    60: ("Bottom Ejector Plate", "bottom_ejector_plate"),
    61: ("Ejector Plate", "ejector_plate"),
    # 6061
    68: ("Top Clamp Plate", "top_clamp_plate"),
    69: ("A Plate", "a_plate"),
    70: ("B Plate", "b_plate"),
    71: ('"X" Plate', "stripper_plate"),
    73: ("Support Plate", "support_plate"),
    74: ("Rails", "rail"),
    75: ("Bottom Clamp Plate", "bottom_clamp_plate"),
    76: ("Bottom Ejector Plate", "bottom_ejector_plate"),
    77: ("Ejector Plate", "ejector_plate"),
}

# The spare rows of every block, so a plate the macro parked on one -- or that a
# grade change moved onto one -- is still read and still priced. Generic on
# purpose: they carry no slot, so the name in column A is the only thing that
# says which plate they hold, and _role_from_name reads exactly that.
for _row in (13, 16, 17, 18, 33, 34, 44, 45, 46, 47, 48, 62, 63, 64, 72, 78, 79, 80):
    _STD_QUOTE_STEEL_ROWS.setdefault(_row, ("", "steel_plate"))

_BMS_QUOTE_STEEL_ROWS = {
    22: ("TCP", "tcp"),
    23: ("BCP", "bcp"),
    24: ("", "steel_plate"),
    25: ("", "steel_plate"),
    26: ("", "steel_plate"),
    27: ("", "steel_plate"),
    28: ("", "steel_plate"),
    29: ("", "steel_plate"),
    30: ("", "steel_plate"),
    31: ("ID Holder", "id_holder"),
    32: ("OD Holder", "od_holder"),
    33: ("ID Pot", "id_pot"),
    34: ("OD Pot", "od_pot"),
}

# A POT JOB IS NOT LOCKED TO THE 4140 BLOCK.
#
# The macro writes every BMS plate into rows 22-34, which is why the map above
# stops there. But the estimator can now change a plate's steel from the Parts
# tab, and sheet_grades.rewrite_grades physically MOVES that plate's row into
# the block for its new grade -- TCP at #1 A-36 lands on row 6.
#
# With the map ending at 34 the moved plate had no quote row the reader could
# see: its price came back 0.00 and the Parts tab showed "sheet" where the
# money used to be. The plate did not vanish, because _load_steel_from_j000
# still lists it off the steel sheet -- it just lost its price, which is worse
# than vanishing, since a $0 line still totals.
#
# Every other block's rows are therefore readable too. They carry no default
# name and the generic role, so an empty template label is skipped on qty and a
# real plate is identified by the name the macro (or the move) wrote in column A
# -- which is what _BMS_NAME_TO_ROLE resolves back to a role anyway.
for _row in (
    list(range(6, 19))       # #1 A-36, including the 16-18 premium rows
    + list(range(38, 49))    # #3 P20, including 46-48
    + list(range(52, 65))    # #7 420-SS, including 62-64
    + list(range(68, 81))    # ALM 6061
):
    _BMS_QUOTE_STEEL_ROWS.setdefault(_row, ("", "steel_plate"))

_BMS_NAME_TO_ROLE = {
    "tcp": "tcp",
    "bcp": "bcp",
    "id holder": "id_holder",
    "od holder": "od_holder",
    "id pot": "id_pot",
    "od pot": "od_pot",
    "id pot block": "id_pot",
    "od pot block": "od_pot",
    "id holder block": "id_holder",
    "od holder block": "od_holder",
}

# THE NAME IN THE CELL BEATS THE ROW NUMBER.
#
# The row maps assign a role by POSITION, which is a guess about which slot the
# macro used. StdQuoteRowFor routes by (slot, grade) and several slots share a
# grade block, so the guess can be wrong -- and when it is, the merge onto the
# J000 row list matches the wrong plate.
#
# C18597 is exactly that. The macro wrote Die Plate to P20 row 41 and Stripper
# Plate to row 42, while the positional map called row 41 "stripper_plate". So the
# Stripper Plate row took DIE PLATE's hours and price ($216.65), and Die Plate and
# Die Backup Plate -- which had no quote row carrying their role -- came back at
# $0.00 each. Net: $284 of steel silently missing from a $1,476 sheet, with one
# plate's price sitting on another plate's row.
#
# The macro writes the plate NAME into column A of the row it chose. That name is
# authoritative; the row number is not. This map turns it back into a role.
_STD_NAME_TO_ROLE = {
    "top clamp plate": "top_clamp_plate",
    "top clamping plate": "top_clamp_plate",
    "manifold plate": "manifold_plate",
    "backing plate": "support_plate",
    "a plate": "a_plate",
    "cavity plate": "a_plate",
    "b plate": "b_plate",
    "core plate": "b_plate",
    "stripper plate": "stripper_plate",
    "runner stripper plate": "stripper_plate",
    "die plate": "die_plate",
    "die backup plate": "die_backup_plate",
    "die back up plate": "die_backup_plate",
    "support plate": "support_plate",
    "sc retainer plate": "sc_retainer_plate",
    "sc backup plate": "sc_backup_plate",
    "bottom clamp plate": "bottom_clamp_plate",
    "bottom clamping plate": "bottom_clamp_plate",
    "rails": "rail",
    "rail": "rail",
    "pin plate": "pin_plate",
    "ejector plate": "ejector_plate",
    "bottom ejector plate": "bottom_ejector_plate",
    "ejector backup plate": "bottom_ejector_plate",
    "ejector back-up plate": "bottom_ejector_plate",
    "ejector retainer plate": "ejector_plate",
    '"x" plate': "stripper_plate",
    '"y" plate': "sc_retainer_plate",
    "x plate": "stripper_plate",
    "y plate": "sc_retainer_plate",
}

_BMS_ROLE_LABELS = {
    "tcp": "TCP",
    "bcp": "BCP",
    "id_holder": "ID Holder",
    "od_holder": "OD Holder",
    "id_pot": "ID Pot",
    "od_pot": "OD Pot",
    "steel_plate": "Steel Plate",
}


def _cell_str(val) -> str:
    if val is None:
        return ""
    return str(val).strip()


def _canonical_steel_name(name: str, role: str, default_name: str = "") -> str:
    """Never return a blank steel description — UI shows '--' otherwise."""
    n = _cell_str(name)
    if n and n not in {"--", "-", "None", "none"}:
        return n
    if default_name:
        return default_name
    return _BMS_ROLE_LABELS.get(role, "") or role or "Steel Plate"


def _iter_workbook_rows(wb_path: Path, sheet_names: tuple[str, ...] | None = None):
    """Yield (sheet_name, 1-based_row_index, cell_values_list)."""
    suffix = wb_path.suffix.lower()
    if suffix == ".xls":
        try:
            import xlrd  # type: ignore
        except ImportError:
            return
        try:
            book = xlrd.open_workbook(str(wb_path))
        except Exception:
            return
        for name in book.sheet_names():
            if sheet_names and name not in sheet_names:
                low = name.lower()
                if not any(s.lower() in low for s in sheet_names):
                    continue
            sh = book.sheet_by_name(name)
            for r in range(sh.nrows):
                yield name, r + 1, list(sh.row_values(r))
        return

    try:
        import openpyxl  # type: ignore
    except ImportError:
        return
    try:
        # data_only=True needs Excel to have calculated+saved formulas.
        # Fall back to formula workbook if cached values are missing.
        wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    except Exception:
        return
    try:
        for name in wb.sheetnames:
            if sheet_names and name not in sheet_names:
                low = name.lower()
                if not any(s.lower() in low for s in sheet_names):
                    continue
            ws = wb[name]
            for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=280, values_only=True), start=1):
                yield name, r_idx, list(row) if row else []
    finally:
        wb.close()


def _score_quote_workbook(path: Path) -> int:
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
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        score += 5
    # Prefer shorter names (fewer date/initials suffixes).
    score -= max(0, len(path.name) - 40) // 5
    return score


def _score_steel_workbook(path: Path) -> int:
    nm = path.name.upper()
    if "PURCHASED" in nm:
        return -1000
    score = 0
    if "STEEL" in nm and "SHEET" in nm:
        score += 80
    if "J000" in nm:
        score += 40
    if "MACHINING" in nm:
        score += 20
    # Quote_Steel_Grinding also matches *steel* — de-prioritize vs true J000.
    if "QUOTE" in nm and "GRIND" in nm:
        score -= 30
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        score += 5
    score -= max(0, len(path.name) - 40) // 5
    return score


def _iter_candidate_workbooks(job_dir: Path):
    seen: set[str] = set()
    for sub in ("", "documents"):
        base = job_dir / sub if sub else job_dir
        if not base.exists():
            continue
        for pat in ("*.xlsx", "*.xlsm", "*.xls"):
            for hit in sorted(base.glob(pat)):
                key = hit.name.lower()
                if key.startswith("~$"):
                    continue
                if key in seen:
                    continue
                seen.add(key)
                yield hit


def _find_workbook(job_dir: Path) -> Path | None:
    """Best quote workbook (Quote_Steel_Grinding), else best steel sheet."""
    quote = _find_quote_workbook(job_dir)
    if quote:
        return quote
    return _find_steel_workbook(job_dir)


# Both pickers now defer to app/workbooks.py, which weighs the JOB'S OWN NAME
# above everything else.
#
# The scorers above could not tell this job's steel sheet from a stray copy of the
# blank template, and on C18500 they preferred the stray one: "J000-STEEL SHEET-
# std.xls" scored 120 (+80 steel sheet, +40 J000) against 80 for the real
# "Majestic-5475-C18500 STEEL SHEET.xls". The Parts tab then listed the template's
# leftovers -- a plate named "0", a row of 23s -- beside the five real plates.
#
# _score_quote_workbook / _score_steel_workbook are kept because they are the
# documented history of what the filenames mean, and workbooks.py builds on the
# same signals.
def _find_quote_workbook(job_dir: Path) -> Path | None:
    return workbooks.find(job_dir).get("quote")


def _find_steel_workbook(job_dir: Path) -> Path | None:
    return workbooks.find(job_dir).get("steel")


def _renamed_name_to_role(job_dir: Path) -> dict[str, str]:
    """{normalised renamed text: role} for every rename saved on this job.

    A RENAME MUST NOT DELETE A PRICE, AND IT WAS DOING EXACTLY THAT.

    Both loaders below turn the text in the cell back into a role, using
    _BMS_NAME_TO_ROLE / _STD_NAME_TO_ROLE, and _merge_steel_hours_prices then
    pairs the steel sheet against the quote block on an EXACT role match.

    Rename "OD Holder" to "OD Holders" and the steel sheet's row no longer maps:
    "od holders" is in neither table, so it came back as the catch-all
    "steel_plate" while the quote row -- which still had a positional fallback --
    stayed "od_holder". No match, no merge, and a $391.49 plate showed $0.00 on
    the Parts tab with the word "sheet" where its price had been.

    That is the same closing-the-loop problem plate_names.sheet_name_aliases
    already solves for stock dimensions: whatever the estimator renamed a plate
    to is, by definition, another name for that plate's role.
    """
    out: dict[str, str] = {}
    try:
        from . import plate_names
    except Exception:
        return out
    try:
        overrides = plate_names.get_overrides(job_dir)
    except Exception:
        return out
    for key, name in overrides.items():
        if not key.startswith(plate_names.ROLE_PREFIX):
            continue
        role = key[len(plate_names.ROLE_PREFIX):].strip()
        norm = _norm(name)
        if role and norm:
            out[norm] = role
    return out


def _role_from_name(name: str, extra: dict[str, str] | None = None) -> str:
    """Role for a plate name: this job's renames first, then the shared tables.

    The bare plural is tried last. The macro writes "Rails" where the role tables
    say "rail", and an estimator typing "OD Holders" for a pair of them is the
    same shape of edit -- neither should cost the plate its role.
    """
    cn = _norm(name)
    if not cn:
        return ""
    for table in (extra or {}, _BMS_NAME_TO_ROLE, _STD_NAME_TO_ROLE):
        hit = table.get(cn)
        if hit:
            return hit
    if cn.endswith("s"):
        singular = cn[:-1].strip()
        for table in (extra or {}, _BMS_NAME_TO_ROLE, _STD_NAME_TO_ROLE):
            hit = table.get(singular)
            if hit:
                return hit
    return ""


def load_steel_plate_lines(job_dir: Path, base_type: str = "bms") -> list[dict]:
    """Steel plates from J000 (names) merged with QuoteWorksheet hours/price.

    Module6121 writes plate *names* reliably on the J000 Steel Order sheet.
    The QuoteWorksheet grade blocks hold stock dims + hours/price. Prefer J000 for
    the row list so the UI never shows blank Description='--'.

    `base_type` selects which QuoteWorksheet grade block is read. It defaults to
    "bms" so an old caller keeps its previous behaviour exactly; the quote builder
    passes the job's real base type.
    """
    quote_wb = _find_quote_workbook(job_dir)
    steel_wb = _find_steel_workbook(job_dir)

    # Fed to BOTH loaders, so a renamed plate resolves to the same role on each
    # side and _merge_steel_hours_prices can still pair them. See
    # _renamed_name_to_role.
    renamed = _renamed_name_to_role(job_dir)

    j000_lines = _load_steel_from_j000(steel_wb, renamed) if steel_wb else []
    # Quote workbook sometimes *is* the only file; also try J000 sheets there.
    if not j000_lines and quote_wb and quote_wb != steel_wb:
        j000_lines = _load_steel_from_j000(quote_wb, renamed)

    quote_lines = (
        _load_steel_from_quote_block(quote_wb, base_type, renamed) if quote_wb else []
    )

    if j000_lines:
        return _merge_steel_hours_prices(j000_lines, quote_lines)

    if quote_lines:
        # Ensure every line has a visible name.
        for line in quote_lines:
            line["component"] = _canonical_steel_name(
                line.get("component") or "",
                line.get("role") or "",
            )
        return quote_lines

    return _load_steel_from_bom_match(job_dir)


def _merge_steel_hours_prices(base_lines: list[dict], quote_lines: list[dict]) -> list[dict]:
    """Copy hours/price from quote block onto J000 rows matched by role (then dims)."""
    if not quote_lines:
        return base_lines

    by_role: dict[str, list[dict]] = {}
    for q in quote_lines:
        role = q.get("role") or ""
        by_role.setdefault(role, []).append(q)

    used: set[int] = set()
    out: list[dict] = []
    for line in base_lines:
        merged = dict(line)
        merged["component"] = _canonical_steel_name(
            merged.get("component") or "",
            merged.get("role") or "",
        )
        role = merged.get("role") or ""
        candidates = by_role.get(role) or []
        best = None
        best_idx = -1
        best_score = -1.0
        for i, q in enumerate(candidates):
            if id(q) in used:
                continue
            score = 0.0
            for key in ("thickness", "width", "length"):
                a = float(merged.get(key) or 0)
                b = float(q.get(key) or 0)
                if a > 0 and b > 0 and abs(a - b) < 0.35:
                    score += 1.0
            if float(q.get("price") or 0) > 0:
                score += 0.25
            if float(q.get("hours") or 0) > 0:
                score += 0.25
            if score > best_score:
                best_score = score
                best = q
                best_idx = i
        if best is not None and best_score >= 1.0:
            used.add(id(best))
            if best.get("hours"):
                merged["hours"] = best["hours"]
            if float(best.get("price") or 0) > 0:
                merged["price"] = best["price"]
                merged["price_source"] = best.get("price_source") or merged.get("price_source")
            # The steel sheet's own "STEEL TYPE" text wins; the quote block the
            # row sits in fills in when that cell is blank.
            if not merged.get("grade") and best.get("grade"):
                merged["grade"] = best["grade"]
            # Keep J000 finished dims (source of truth for T/W/L). Quote stock
            # sizes can swap axes vs the steel sheet — do not overwrite.
        out.append(merged)
    return out


def _load_steel_from_quote_block(
    wb_path: Path | None,
    base_type: str = "bms",
    renamed: dict[str, str] | None = None,
) -> list[dict]:
    """Steel rows out of the QuoteWorksheet grade blocks.

    `base_type` picks the row map -- "bms" for a pot base, anything else for a
    standard/PCS base. See the note above _STD_QUOTE_STEEL_ROWS for why one map
    cannot serve both.
    """
    if not wb_path:
        return []
    row_map = (
        _BMS_QUOTE_STEEL_ROWS
        if (base_type or "").strip().lower() in ("bms", "pot", "pot_block")
        else _STD_QUOTE_STEEL_ROWS
    )
    lines: list[dict] = []
    for sheet_name, row_idx, cells in _iter_workbook_rows(
        wb_path, ("QuoteWorksheet", "Quote")
    ):
        # Prefer the real QuoteWorksheet; skip fuzzy "Quote Summary" etc. when
        # the exact sheet exists in this workbook iteration order.
        if sheet_name not in ("QuoteWorksheet", "Quote") and "quote" not in sheet_name.lower():
            continue
        if row_idx not in row_map:
            continue
        default_name, role = row_map[row_idx]
        while len(cells) < 10:
            cells.append(None)

        # Col A preferred; some templates put the label in Col B.
        raw_name = _cell_str(cells[0]) or _cell_str(cells[1])
        name = _canonical_steel_name(raw_name, role, default_name)
        qty = _safe_float(cells[2], 0.0)
        thickness = _parse_fraction(cells[3])
        width = _parse_fraction(cells[4])
        length = _parse_fraction(cells[5])
        hours = _safe_float(cells[6], 0.0)
        price = _safe_float(cells[7], 0.0)

        # Skip empty / zero-qty template rows (flipper blanks, etc.).
        if qty <= 0 and thickness <= 0 and width <= 0 and length <= 0:
            continue
        if qty <= 0:
            continue
        # Skip spare rows that have no real name and no default (blank extras).
        if not name or (not default_name and role == "steel_plate" and not raw_name):
            if role == "steel_plate" and not raw_name:
                continue

        # Name first (the macro wrote it into this row), row position only as a
        # fallback for a blank/spare row. See _STD_NAME_TO_ROLE.
        role_key = (
            _role_from_name(raw_name, renamed)
            or _role_from_name(name, renamed)
            or role
        )
        if role_key in _BMS_ROLE_LABELS and (not raw_name or role_key in {"tcp", "bcp", "id_holder", "od_holder", "id_pot", "od_pot"}):
            # Keep canonical BMS labels for the six pot-block slots.
            if default_name:
                name = default_name
        cu_in = 0.0
        if thickness > 0 and width > 0 and length > 0:
            cu_in = round(qty * thickness * width * length, 2)

        lines.append(
            {
                "component": name,
                "role": role_key,
                "qty": qty,
                "thickness": thickness,
                "width": width,
                "length": length,
                "hours": hours if hours > 0 else None,
                "cu_in": cu_in if cu_in > 0 else None,
                "price": round(price, 2) if price > 0 else 0.0,
                # The block this row sits in IS its grade -- that is the whole
                # reason a grade change moves the row. Second-best to the steel
                # sheet's own type text, and the only source when that is blank.
                "grade": grade_for_quote_row(row_idx),
                "price_source": f"quote_workbook:{sheet_name}:row{row_idx}",
                "source_sheet": sheet_name,
            }
        )
    return lines


# Which grade block each QuoteWorksheet row belongs to. Transcribed from the same
# StdQuoteRowFor bands app/sheet_grades.py writes into, and the reason a moved row
# reprices: the per-pound rate lives in the block, not in the plate.
_QUOTE_ROW_GRADE_BANDS = (
    (6, 19, "A36"),
    (20, 35, "4140"),
    (36, 49, "P20"),
    (50, 65, "420SS"),
    (66, 81, "6061"),
)


def grade_for_quote_row(row_idx: int) -> str:
    for lo, hi, grade in _QUOTE_ROW_GRADE_BANDS:
        if lo <= row_idx <= hi:
            return grade
    return ""



# Labels that live in the plate block but are not plates.
_NON_PLATE_LABELS = {
    "total",
    "total # of plates",
    "total number of plates",
    "plate",
    "sizes",
    "qty",
    "steel type",
    "grinder notes",
    "notes",
}


def _is_plate_row(name: str, qty: float, t: float, w: float, l: float) -> bool:
    """Does this row describe an actual plate?

    The dimension checks upstream (qty > 0, all three sizes > 0) are not enough. A
    steel sheet that was filled in by hand, or left half-populated from a previous
    job, carries rows that pass all of them and are still not plates. C18500's
    folder had a stray template whose Machining Sheet held:

        row 21    1 | 0  | 3.625 | 11.875 | 9.875
        row 25   23 | 23 | 23    | 23     | 23

    Both survived every existing check and reached the Parts & Pricing tab as
    plates -- one named "0", one named "23" with every column 23.
    """
    n = _norm(name)
    if not n:
        return False
    if n in _NON_PLATE_LABELS:
        return False
    if n.startswith("total "):
        return False

    # A plate name is not a bare number. "0", "23", "0.0" are spilled values or a
    # count row, never a description.
    if re.fullmatch(r"[0-9]+(?:[.,][0-9]+)?", n):
        return False

    # One value smeared across the whole row: qty and all three sizes identical.
    # No real plate is 23 thick by 23 by 23 in a quantity of 23.
    vals = [qty, t, w, l]
    if all(v > 0 for v in vals) and max(vals) - min(vals) < 1e-6:
        return False

    return True


def _load_steel_from_j000(
    wb_path: Path | None, renamed: dict[str, str] | None = None
) -> list[dict]:
    """J000 Steel Order / Machining Sheet: A=Qty B=Name C=T E=W G=L H=Steel type."""
    if not wb_path:
        return []
    # Collected PER SHEET, not merged, because the tabs disagree and one of them
    # can be stale. See the sheet-preference note where this is returned.
    per_sheet: dict[str, list[dict]] = {}
    seen_per_sheet: dict[str, set[str]] = {}
    for sheet_name, row_idx, cells in _iter_workbook_rows(
        wb_path, ("Steel Order", "Machining Sheet", "Steel")
    ):
        if row_idx < 19:
            continue
        lines = per_sheet.setdefault(sheet_name, [])
        seen = seen_per_sheet.setdefault(sheet_name, set())
        while len(cells) < 8:
            cells.append(None)
        name = _cell_str(cells[1])
        if not name:
            continue
        # SAME NAME->ROLE RESOLUTION AS THE QUOTE BLOCK, OR THE MERGE CANNOT MATCH.
        #
        # _merge_steel_hours_prices pairs the two sources on an EXACT role match.
        # PLATE_SHEET_NAMES below has no entry for the die stack, so J000 called
        # Die Plate and Die Backup Plate the generic "steel_plate" while the quote
        # block called them "die_plate"/"die_backup_plate" -- no match, and both
        # plates came back at $0.00 on C18597 despite the workbook holding $216.65
        # and $134.23 for them.
        role = _role_from_name(name, renamed)
        # Also accept standard plate names via PLATE_SHEET_NAMES keywords.
        if not role:
            cn = _norm(name)
            for r, names in PLATE_SHEET_NAMES.items():
                if any(n in cn for n in names):
                    role = r
                    break
        if not role:
            # Still show named steel rows Module6121 wrote (extras like Flipper).
            role = "steel_plate"
        # Column H is the STEEL TYPE the shop orders against ("#2 4140"). Read so
        # the Parts tab's steel picker shows what the sheet SAYS instead of
        # falling back to the 4140 default for every row -- which made a plate
        # the macro put in #1 A-36 look like an unedited #2, and made the
        # estimator's own change look like it had not stuck.
        grade = plate_grades.clean_grade(_cell_str(cells[7]))
        qty = _safe_float(cells[0], 1.0)
        if qty <= 0:
            continue
        thickness = _parse_fraction(cells[2])
        width = _parse_fraction(cells[4])
        length = _parse_fraction(cells[6])
        if thickness <= 0 or width <= 0 or length <= 0:
            continue
        if not _is_plate_row(name, qty, thickness, width, length):
            continue
        key = _norm(name)
        if key in seen:
            continue
        seen.add(key)
        cu_in = round(qty * thickness * width * length, 2)
        lines.append(
            {
                "component": _canonical_steel_name(name, role),
                "role": role,
                "qty": qty,
                "thickness": thickness,
                "width": width,
                "length": length,
                "hours": None,
                "cu_in": cu_in,
                "price": 0.0,
                "grade": grade,
                "price_source": f"steel_workbook:{sheet_name}:row{row_idx}",
                "source_sheet": sheet_name,
            }
        )

    # ONE TAB WINS. THEY ARE NOT TWO HALVES OF THE SAME LIST.
    #
    # Both tabs hold the same plate list, so merging them can only add rows that
    # one tab has and the other does not -- and that difference is stale template
    # content, not extra steel.
    #
    # FillStandardBaseSteel writes rows 19..18+StdCount and never clears below, and
    # the J000 template ships with a previous job's rows already in it. C18500 is a
    # 5-plate job: its Steel Order tab holds exactly those 5 (rows 24-28 blank),
    # while its Machining Sheet still carries the template's rows 24-26 --
    # Rails 1.688 x 2.5 x 12, Pin Plate 0.5 x 7.375 x 11.985, Ejector Plate
    # 1 x 7.375 x 11.985 -- and a total of 9. Merging put three plates on the quote
    # that are not in the mold.
    #
    # Steel Order is preferred because it is the sheet the shop actually orders
    # from, and the one whose row count the macro's own log reports.
    for preferred in ("Steel Order", "Machining Sheet", "Steel"):
        for name, rows in per_sheet.items():
            if rows and preferred.lower() in name.lower():
                return rows
    for rows in per_sheet.values():
        if rows:
            return rows
    return []


def _load_steel_from_bom_match(job_dir: Path) -> list[dict]:
    """Fallback: XT_Export_BOM_Match_Report.csv QuoteName + CAD dims."""
    for name in (
        "XT_Export_BOM_Match_Report.csv",
        "pdf/XT_Export_BOM_Match_Report.csv",
        "documents/XT_Export_BOM_Match_Report.csv",
    ):
        rows = _read_csv_rows(job_dir / name)
        if not rows:
            continue
        lines: list[dict] = []
        for row in rows:
            qn = _cell_str(row.get("QuoteName") or row.get("quoteName"))
            if not qn:
                continue
            role = _BMS_NAME_TO_ROLE.get(_norm(qn), "")
            if not role:
                continue
            qty = _safe_float(row.get("Qty") or row.get("QTY") or 1, 1.0)
            thickness = _safe_float(row.get("CAD_Thickness") or row.get("BOM_Thickness"))
            width = _safe_float(row.get("CAD_Width") or row.get("BOM_Width"))
            length = _safe_float(row.get("CAD_Length") or row.get("BOM_Length"))
            if thickness <= 0 or width <= 0 or length <= 0:
                continue
            cu_in = round(qty * thickness * width * length, 2)
            lines.append(
                {
                    "component": qn,
                    "role": role,
                    "qty": qty,
                    "thickness": thickness,
                    "width": width,
                    "length": length,
                    "hours": None,
                    "cu_in": cu_in,
                    "price": 0.0,
                    "price_source": "bom_match_report",
                    "source_sheet": "",
                }
            )
        if lines:
            return lines
    return []


def load_pullcore_lines(job_dir: Path) -> list[dict]:
    """Pull cores & keys from Pullcore Prices.csv (Module6121 WritePullcorePriceFile)."""
    rows = load_job_pullcore(job_dir)
    lines: list[dict] = []
    for row in rows:
        name = _cell_str(
            row.get("Pull Core / Key")
            or row.get("Description")
            or row.get("Component")
        )
        if not name or name.upper() in ("TOTAL", "RATE ($/IN3)", "RATE"):
            continue
        if name.upper().startswith("RATE"):
            continue
        qty = _safe_float(row.get("Qty") or row.get("QTY") or 1, 1.0)
        thickness = _safe_float(row.get("Thickness"))
        width = _safe_float(row.get("Width"))
        length = _safe_float(row.get("Length"))
        cu_in = _safe_float(row.get("Cu In") or row.get("Cu. In.") or row.get("CuIn"))
        price = _safe_float(row.get("Price USD") or row.get("Price"))
        if cu_in <= 0 and thickness > 0 and width > 0 and length > 0:
            cu_in = round(qty * thickness * width * length, 2)
        lines.append(
            {
                "component": name,
                "role": "pullcore",
                "qty": qty,
                "thickness": thickness,
                "width": width,
                "length": length,
                "hours": None,
                "cu_in": cu_in if cu_in > 0 else None,
                "price": round(price, 2),
                "price_source": "job_csv:Pullcore Prices.csv",
                "material": _cell_str(row.get("Material")),
            }
        )
    return lines


# QuoteWorksheet "Components" block, 1-based columns, exactly as
# WritePurchasedToComponentsArea in Module6121.bas writes them:
#   K=Name  L=STD  M=QTY  N=Dia  O=Length  P=Width  Q=Part No.  R=Unit Price
#   S=Total (a template formula, M*R)
_COMP_FIRST_ROW = 4
_COMP_LAST_ROW = 47
_COMP_COL_NAME = 11
_COMP_COL_QTY = 13
_COMP_COL_DIA = 14
_COMP_COL_LENGTH = 15
_COMP_COL_WIDTH = 16
_COMP_COL_PARTNO = 17
_COMP_COL_PRICE = 18
_COMP_COL_TOTAL = 19

# Pre-printed category labels the template ships with. When the macro has written
# real components these are overwritten; when it has not, they are all that is
# there and they are not purchased parts.
_COMP_TEMPLATE_LABELS = {
    "leader pins", "l p bushings", "return pins", "sprue bushing", "locating ring",
    "support pillars", "guided ejec pins", "guided ejec bush", "side locks",
    "slide retainers", "angle pins", "o 7 32", "r 1 2", "components",
}


def load_purchased_from_quote_workbook(job_dir: Path) -> list[dict]:
    """Purchased components and their prices out of the quote workbook itself.

    WHY THIS EXISTS
        Purchased lines used to come only from "Purchased Components Quote.csv".
        Module6121 writes that CSV in the same pass that fills the Components block
        of the quote workbook -- and both passes were switched off for standard
        (non-BMS) jobs. So on C18517 the web app listed three purchased parts with
        no price and no CSV to read: leader pins, bushings, hardware, all "--".

        The quote workbook is the file the estimator actually works in, so it is
        also the right place to read a price back from -- including a price typed
        in by hand after the macro ran, which no CSV will ever know about.
    """
    wb = _find_quote_workbook(job_dir)
    if not wb:
        return []

    lines: list[dict] = []
    for sheet_name, row_idx, cells in _iter_workbook_rows(wb, ("QuoteWorksheet",)):
        if row_idx < _COMP_FIRST_ROW or row_idx > _COMP_LAST_ROW:
            continue
        while len(cells) < _COMP_COL_TOTAL:
            cells.append(None)

        name = _cell_str(cells[_COMP_COL_NAME - 1])
        if not name:
            continue
        if _norm(name) in _COMP_TEMPLATE_LABELS:
            continue

        qty = _safe_float(cells[_COMP_COL_QTY - 1], 0.0)
        unit = _safe_float(cells[_COMP_COL_PRICE - 1], 0.0)
        total = _safe_float(cells[_COMP_COL_TOTAL - 1], 0.0)

        # An untouched template row carries a name but no quantity and no money.
        if qty <= 0 and unit <= 0 and total <= 0:
            continue

        if total <= 0 and unit > 0:
            total = unit * max(qty, 1.0)
        if unit <= 0 and total > 0 and qty > 0:
            unit = total / qty

        lines.append(
            {
                "component": name,
                "role": "purchased_component",
                "qty": qty or 1.0,
                "thickness": _safe_float(cells[_COMP_COL_WIDTH - 1]) or None,
                "width": _safe_float(cells[_COMP_COL_DIA - 1]) or None,
                "length": _safe_float(cells[_COMP_COL_LENGTH - 1]) or None,
                "hours": None,
                "cu_in": None,
                "price": round(total, 2),
                "price_source": f"quote_workbook:{sheet_name}:row{row_idx}",
                "vendor": "",
                "part_number": _cell_str(cells[_COMP_COL_PARTNO - 1]),
                "unit_price": round(unit, 2),
                "category": name,
            }
        )
    return lines


def load_purchased_lines(job_dir: Path) -> list[dict]:
    """Purchased components: the job CSV first, then the quote workbook.

    The CSV is preferred where it exists because it carries the vendor and the
    part number the macro matched. The workbook fills in what the CSV does not
    have -- which on a standard job is everything, and on any job is any price the
    estimator typed straight into the sheet.
    """
    rows = load_job_purchased_quote(job_dir)
    lines: list[dict] = []
    for row in rows:
        desc = _cell_str(row.get("Description") or row.get("Component"))
        comp = _cell_str(row.get("Component") or desc)
        if not desc and not comp:
            continue
        if (desc or comp).upper() == "TOTAL":
            continue
        qty = _safe_float(row.get("QTY") or row.get("Qty") or 1, 1.0)
        unit = _safe_float(row.get("UnitPrice"))
        ext = _safe_float(row.get("Extended"))
        if ext <= 0 and unit > 0:
            ext = unit * max(qty, 1)
        lines.append(
            {
                "component": desc or comp,
                "role": "purchased_component",
                "qty": qty,
                "thickness": None,
                "width": None,
                "length": None,
                "hours": None,
                "cu_in": None,
                "price": round(ext, 2),
                "price_source": "job_csv:Purchased Components Quote.csv",
                "vendor": _cell_str(row.get("Vendor")),
                "part_number": _cell_str(row.get("PartNumber")),
                "unit_price": unit,
                "category": comp,
            }
        )

    # Fold in the workbook's Components block: rows the CSV does not have at all,
    # and prices for CSV rows that came back at zero.
    wb_lines = load_purchased_from_quote_workbook(job_dir)
    if not wb_lines:
        return lines
    if not lines:
        return wb_lines

    by_key = {_norm(l.get("component") or ""): l for l in lines}
    for w in wb_lines:
        key = _norm(w.get("component") or "")
        hit = by_key.get(key)
        if hit is None:
            lines.append(w)
            continue
        if float(hit.get("price") or 0) <= 0 < float(w.get("price") or 0):
            hit["price"] = w["price"]
            hit["unit_price"] = w.get("unit_price") or hit.get("unit_price")
            hit["price_source"] = w["price_source"]
        if not hit.get("part_number") and w.get("part_number"):
            hit["part_number"] = w["part_number"]
    return lines


def load_quote_summary(job_dir: Path) -> dict:
    """Best-effort summary totals from QuoteWorksheet (hours / price / commission)."""
    wb = _find_workbook(job_dir)
    summary = {
        "total_hours": None,
        "total_price_rough": None,
        "total_price_finish": None,
        "commission_pct": None,
        "commission_rough": None,
        "commission_finish": None,
        "grand_total_rough": None,
        "grand_total_finish": None,
    }
    if not wb:
        return summary

    for sheet_name, row_idx, cells in _iter_workbook_rows(
        wb, ("QuoteWorksheet", "Quote")
    ):
        if not cells:
            continue
        label = _norm(_cell_str(cells[0]))
        nums = [_safe_float(c) for c in cells[1:6] if _safe_float(c) != 0]
        if "total hours" in label and nums:
            summary["total_hours"] = nums[0]
        elif label == "commission" or label.startswith("commission"):
            # Often: Commission | 6% | $545 | $550
            pct = None
            for c in cells[1:4]:
                s = _cell_str(c)
                if s.endswith("%"):
                    pct = _safe_float(s.replace("%", ""))
            money = [_safe_float(c) for c in cells[1:5] if _safe_float(c) > 1]
            if pct is not None:
                summary["commission_pct"] = pct
            if len(money) >= 1:
                summary["commission_rough"] = money[0]
            if len(money) >= 2:
                summary["commission_finish"] = money[1]
        elif label == "total price":
            money = [_safe_float(c) for c in cells[1:5] if _safe_float(c) > 1]
            # First Total Price row is subtotal; later one (after commission) is grand.
            if summary["total_price_rough"] is None and money:
                summary["total_price_rough"] = money[0]
                if len(money) >= 2:
                    summary["total_price_finish"] = money[1]
            elif money:
                summary["grand_total_rough"] = money[0]
                if len(money) >= 2:
                    summary["grand_total_finish"] = money[1]
    return summary


def _match_shop_price(role: str, component_name: str, shop_rows: list[dict]) -> tuple[float, str]:
    """Return (unit_price, csv_component_matched)."""
    comp_norm = _norm(component_name)
    keywords = ROLE_CSV_KEYWORDS.get(role, [])
    best_price = 0.0
    best_match = ""
    is_plate_role = role.endswith("_plate") or role in ("rail", "pin_plate")

    for row in shop_rows:
        csv_comp = _norm(row.get("Component") or "")
        price = _safe_float(row.get("UnitPrice"))
        if price <= 0 or not csv_comp:
            continue

        # Plates must not pick up pin/bushing hardware rows.
        if is_plate_role and any(x in csv_comp for x in ("pin", "bushing", "strap", "ring", "insulation")):
            if "plate" not in csv_comp:
                continue
        if role == "leader_pin" and "pin" not in csv_comp:
            continue
        if role == "leader_pin_bushing" and "bush" not in csv_comp:
            continue

        if csv_comp in comp_norm or comp_norm in csv_comp:
            return price, row.get("Component", "")

        for kw in keywords:
            if len(kw) < 3:
                continue
            if kw in csv_comp or (kw in comp_norm and not is_plate_role):
                if price > best_price:
                    best_price = price
                    best_match = row.get("Component", "")

    return best_price, best_match


def _match_job_purchased(
    component_name: str, job_rows: list[dict]
) -> tuple[float, str]:
    comp_norm = _norm(component_name)
    for row in job_rows:
        csv_comp = _norm(row.get("Component") or "")
        if not csv_comp:
            continue
        unit = _safe_float(row.get("UnitPrice"))
        qty = _safe_float(row.get("QTY") or row.get("Qty") or 1, 1.0)
        ext = _safe_float(row.get("Extended"))
        if ext > 0:
            return ext, row.get("Component", "")
        if unit > 0:
            return unit * max(qty, 1), row.get("Component", "")
        if csv_comp in comp_norm or comp_norm in csv_comp:
            return unit * max(qty, 1), row.get("Component", "")
    return 0.0, ""


def _parse_fraction(val) -> float:
    s = str(val).strip()
    if not s:
        return 0.0
    m = re.match(r"^(\d+)\s+(\d+)/(\d+)$", s)
    if m:
        return int(m.group(1)) + int(m.group(2)) / int(m.group(3))
    m = re.match(r"^(\d+)/(\d+)$", s)
    if m:
        return int(m.group(1)) / int(m.group(2))
    return _safe_float(s)


def read_sheet_dimensions(
    job_dir: Path, extra_aliases: dict[str, list] | None = None
) -> dict[str, dict]:
    """Parse quote/steel Excel for plate names and T/W/L (stock sizes on quote sheet).

    `extra_aliases` maps role -> additional name strings to match, and exists to
    keep user renames from breaking pricing.

    The scan below finds a plate's stock dimensions by looking for a known
    keyword ("a plate", "support plate", ...) in the row. Once a rename has been
    written into the workbook, that keyword is gone -- so the dimensions for that
    role silently vanish and `price_for_part` falls back to the CAD bounding box.
    A rename that was meant to be cosmetic would quietly change the price.
    Feeding the overrides back in as aliases for their role closes that loop.
    """
    wb_path = _find_workbook(job_dir)
    if not wb_path:
        return {}
    try:
        import openpyxl  # type: ignore
    except ImportError:
        return {}

    out: dict[str, dict] = {}
    try:
        wb = openpyxl.load_workbook(wb_path, read_only=True, data_only=True)
    except Exception:
        return {}

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=250, values_only=True):
            if not row:
                continue
            cells = [str(c).strip() if c is not None else "" for c in row]
            for role, names in PLATE_SHEET_NAMES.items():
                if role in out:
                    continue
                # A renamed plate matches on its new name as well as the stock
                # keywords for its role.
                search = list(names) + [
                    _norm(a) for a in (extra_aliases or {}).get(role, [])
                ]
                for cell in cells:
                    cn = _norm(cell)
                    if any(n in cn for n in search):
                        nums = [_parse_fraction(c) for c in cells if _parse_fraction(c) > 0]
                        if len(nums) >= 3:
                            out[role] = {
                                "thickness": nums[0],
                                "width": nums[1],
                                "length": nums[2],
                                "sheet_name": cell,
                                "source_sheet": sheet_name,
                            }
                        break
    wb.close()
    return out


def price_for_part(
    row: dict,
    shop_rows: list[dict],
    job_purchased: list[dict],
    sheet_dims: dict[str, dict],
) -> dict:
    """Return pricing dict with price, source, and sheet dimensions when available."""
    role = row.get("role", "")
    component = row.get("Component") or row.get("component") or ""
    quote_flag = bool(row.get("quote") or row.get("Quote"))

    dims = sheet_dims.get(role, {})
    thickness = dims.get("thickness") or _safe_float(row.get("Thickness"))
    width = dims.get("width") or _safe_float(row.get("Width"))
    length = dims.get("length") or _safe_float(row.get("Length"))

    if not quote_flag:
        return {
            "price": 0.0,
            "price_source": "",
            "thickness": thickness,
            "width": width,
            "length": length,
        }

    price = 0.0
    source = ""

    if job_purchased:
        price, match = _match_job_purchased(component, job_purchased)
        if price > 0:
            source = f"job_csv:{match}"

    if price <= 0 and shop_rows:
        price, match = _match_shop_price(role, component, shop_rows)
        if price > 0:
            source = f"shop_csv:{match}"

    if price <= 0:
        source = "no_csv_price"

    return {
        "price": round(price, 2),
        "price_source": source,
        "thickness": thickness,
        "width": width,
        "length": length,
    }
