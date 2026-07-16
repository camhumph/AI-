#!/usr/bin/env python3
"""Reverse-train the geometry classifier from completed quote + steel sheets.

For each historical job folder the script:
  1. Finds XT_Export_CAD_Dimensions.csv (CAD geometry ground truth input)
  2. Reads the finished quote worksheet + J000 steel sheet (ground truth names)
  3. Matches steel-sheet plate names to CAD components by size + position
  4. Writes *_CORRECT_ME.csv training files and a summary manifest

Usage:
  python train_from_quote_sheets.py --jobs-root "C:\\CMS_Local_Workspace\\AI_Jobs"
  python train_from_quote_sheets.py --manifest training_manifest.csv

Manifest CSV columns (header row required):
  job_id,job_folder,xt_csv,quote_sheet,steel_sheet
Any column can be blank; the script will search common filenames inside job_folder.
"""
from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from pathlib import Path

BASE = Path(__file__).resolve().parent
OUT_DIR = BASE / "outputs" / "training"
DATA_DIR = BASE / "data" / "training"

# Plate names as they appear on quote/steel sheets -> classifier roles
PLATE_NAME_TO_ROLE = {
    "top clamp plate": "top_clamp_plate",
    "top clamping plate": "top_clamp_plate",
    "a plate": "a_plate",
    "b plate": "b_plate",
    "stripper plate": "stripper_plate",
    "support plate": "support_plate",
    "bottom clamp plate": "bottom_clamp_plate",
    "bottom clamping plate": "bottom_clamp_plate",
    "sc retainer plate": "sc_retainer_plate",
    "sc backup plate": "sc_backup_plate",
    "ejector plate": "ejector_plate",
    "bottom ejector plate": "bottom_ejector_plate",
    "ejector retainer plate": "ejector_plate",
    "ejector backup plate": "bottom_ejector_plate",
    "pin plate": "pin_plate",
    "rail": "rail",
    "rail top": "rail",
    "rail bottom": "rail",
    # BMS / pot-block steel-sheet names (J000 Steel Order / Machining Sheet)
    "tcp": "tcp",
    "top clamping": "tcp",
    "bcp": "bcp",
    "bottom clamping": "bcp",
    "id holder": "id_holder",
    "od holder": "od_holder",
    "id pot": "id_pot",
    "od pot": "od_pot",
    "id pot block": "id_pot",
    "od pot block": "od_pot",
}

# CMS J000 steel sheet column layout (1-based Excel columns, rows start at 19):
#   A(1)=Qty  B(2)=Name  C(3)=Thickness/Height  E(5)=Width  G(7)=Length  H(8)=Steel type
# QuoteWorksheet #2 4140 block uses: C=Qty D=Thickness E=Width F=Length
BMS_STEEL_COL_MAP = {
    "qty_col": 1,
    "name_col": 2,
    "thickness_height_col": 3,  # Height on steel sheet = plate thickness (smallest bbox dim)
    "width_col": 5,
    "length_col": 7,
    "steel_type_col": 8,
    "first_data_row": 19,
    "rule": (
        "CAD bbox dims are sorted Largest→Middle→Smallest as Length, Width, Thickness. "
        "On the J000 steel sheet: Col C = Thickness/Height, Col E = Width, Col G = Length. "
        "Never swap Width and Length; never put thickness into the Length column."
    ),
}

HARDWARE_NAME_TO_ROLE = {
    "leader pin": "leader_pin",
    "leader bushing": "leader_pin_bushing",
    "leader pin bushing": "leader_pin_bushing",
    "latch lock": "latch_lock",
    "latch-lock": "latch_lock",
    "safety strap": "latch_lock",
    "return pin": "return_pin",
    "ejector pin": "ejector_pin",
    "support pillar": "support_pillar",
}

XT_CANDIDATES = (
    "XT_Export_CAD_Dimensions.csv",
    "raw/XT_Export_CAD_Dimensions.csv",
)
QUOTE_CANDIDATES = (
    "*quote*.xls*",
    "*Quote*.xls*",
    "Quote Steel*.xls*",
    "documents/*quote*.xls*",
)
STEEL_CANDIDATES = (
    "*steel*.xls*",
    "STEEL SHEET.xls*",
    "STEEL SHEET.xlsx",
    "*J000*.xls*",
    "documents/*steel*.xls*",
)


def _norm_name(s: str) -> str:
    s = (s or "").lower().strip()
    s = re.sub(r"[_\-/]+", " ", s)
    s = re.sub(r"\s+", " ", s)
    return s


def role_for_sheet_name(name: str) -> str:
    n = _norm_name(name)
    if n in PLATE_NAME_TO_ROLE:
        return PLATE_NAME_TO_ROLE[n]
    for key, role in PLATE_NAME_TO_ROLE.items():
        if key in n:
            return role
    for key, role in HARDWARE_NAME_TO_ROLE.items():
        if key in n:
            return role
    return ""


def safe_float(v) -> float:
    try:
        return float(str(v).strip())
    except Exception:
        return 0.0


def _glob_one(folder: Path, pattern: str) -> Path | None:
    if "*" in pattern:
        hits = sorted(folder.glob(pattern))
        return hits[0] if hits else None
    p = folder / pattern
    return p if p.exists() else None


def find_xt_csv(folder: Path, explicit: str = "") -> Path | None:
    if explicit:
        p = Path(explicit)
        return p if p.exists() else None
    for rel in XT_CANDIDATES:
        p = _glob_one(folder, rel)
        if p:
            return p
    for p in sorted(folder.rglob("XT_Export_CAD_Dimensions.csv")):
        return p
    return None


def find_workbook(folder: Path, patterns: tuple[str, ...], explicit: str = "") -> Path | None:
    if explicit:
        p = Path(explicit)
        return p if p.exists() else None
    for pat in patterns:
        p = _glob_one(folder, pat)
        if p:
            return p
    return None


def read_xt_rows(xt_path: Path) -> list[dict]:
    with xt_path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        return list(csv.DictReader(f))


def _parse_fraction_inch(val) -> float:
    s = str(val).strip()
    if not s:
        return 0.0
    if re.match(r"^[\d.]+$", s):
        return safe_float(s)
    m = re.match(r"^(\d+)\s+(\d+)/(\d+)$", s)
    if m:
        return int(m.group(1)) + int(m.group(2)) / int(m.group(3))
    m = re.match(r"^(\d+)/(\d+)$", s)
    if m:
        return int(m.group(1)) / int(m.group(2))
    return safe_float(s)


def _extract_plates_from_rows(row_values_list: list, sheet_name: str) -> list[dict]:
    """Shared plate extraction from a list of row cell lists.

    CMS J000 Steel Order / Machining Sheet layout (1-based):
      A=Qty B=Name C=Thickness/Height  E=Width  G=Length  H=Steel
    QuoteWorksheet #2 block: C=Qty D=Thickness E=Width F=Length
    Empty spacer columns are skipped when collecting numeric dims.
    """
    found: list[dict] = []
    is_steel = any(k in (sheet_name or "").lower() for k in ("steel", "machining"))
    for row in row_values_list:
        if not row:
            continue
        cells = [str(c).strip() if c is not None and str(c).strip() else "" for c in row]
        if not any(cells):
            continue
        name = ""
        for i, cell in enumerate(cells):
            role = role_for_sheet_name(cell)
            if role and not name:
                name = cell
                # Prefer fixed CMS steel columns when present (C/E/G = indices 2/4/6).
                t = w = l = 0.0
                if is_steel and len(cells) > 6:
                    t = _parse_fraction_inch(cells[2] if len(cells) > 2 else "")
                    w = _parse_fraction_inch(cells[4] if len(cells) > 4 else "")
                    l = _parse_fraction_inch(cells[6] if len(cells) > 6 else "")
                if t <= 0 or w <= 0 or l <= 0:
                    nums = [
                        _parse_fraction_inch(c)
                        for c in cells[i + 1 : i + 10]
                        if _parse_fraction_inch(c) > 0
                    ]
                    if len(nums) >= 3:
                        t, w, l = nums[0], nums[1], nums[2]
                    elif len(nums) == 2:
                        w, l = nums[0], nums[1]
                found.append(
                    {
                        "sheet_name": name,
                        "role": role,
                        "thickness": t,
                        "width": w,
                        "length": l,
                        "height": t,  # CMS steel sheet: Height column = Thickness
                        "source_sheet": sheet_name,
                        "dim_layout": "C=T/H E=W G=L" if is_steel else "sequential",
                    }
                )
                break
        if not name:
            for cell in cells:
                role = role_for_sheet_name(cell)
                if role:
                    nums = [_parse_fraction_inch(c) for c in cells if _parse_fraction_inch(c) > 0]
                    found.append(
                        {
                            "sheet_name": cell,
                            "role": role,
                            "thickness": nums[0] if len(nums) > 0 else 0,
                            "width": nums[1] if len(nums) > 1 else 0,
                            "length": nums[2] if len(nums) > 2 else 0,
                            "height": nums[0] if len(nums) > 0 else 0,
                            "source_sheet": sheet_name,
                            "dim_layout": "sequential",
                        }
                    )
                    break
    return found


def _sheet_name_matches(sheet_name: str) -> bool:
    target_sheets = ("QuoteWorksheet", "Steel Order", "Machining Sheet", "Quote", "Steel")
    if sheet_name in target_sheets:
        return True
    low = sheet_name.lower()
    return any(k in low for k in ("quote", "steel", "machining", "grind"))


def _read_sheet_plates_xls(path: Path) -> list[dict]:
    """Read legacy Excel .xls (97-2003) — common on CMS steel/quote sheets."""
    try:
        import xlrd  # type: ignore
    except ImportError:
        return []
    plates: list[dict] = []
    try:
        book = xlrd.open_workbook(str(path))
    except Exception:
        return []
    for sheet_name in book.sheet_names():
        if not _sheet_name_matches(sheet_name):
            continue
        sh = book.sheet_by_name(sheet_name)
        rows = [sh.row_values(i) for i in range(min(250, sh.nrows))]
        plates.extend(_extract_plates_from_rows(rows, sheet_name))
    return _dedupe_plates(plates)


def read_sheet_plates(xlsx_path: Path) -> list[dict]:
    """Extract named plate rows from quote or steel Excel workbooks (.xls or .xlsx)."""
    suffix = xlsx_path.suffix.lower()
    if suffix == ".xls":
        return _read_sheet_plates_xls(xlsx_path)

    try:
        import openpyxl  # type: ignore
    except ImportError:
        return _read_sheet_plates_csv_fallback(xlsx_path)

    plates: list[dict] = []
    try:
        wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    except Exception:
        return _read_sheet_plates_csv_fallback(xlsx_path)

    for sheet_name in wb.sheetnames:
        if not _sheet_name_matches(sheet_name):
            continue
        ws = wb[sheet_name]
        rows = list(ws.iter_rows(min_row=1, max_row=250, values_only=True))
        plates.extend(_extract_plates_from_rows(rows, sheet_name))
    wb.close()
    return _dedupe_plates(plates)


def _read_sheet_plates_csv_fallback(path: Path) -> list[dict]:
    """If openpyxl is missing, try a sidecar CSV export of the steel sheet."""
    csv_path = path.with_suffix(".csv")
    if not csv_path.exists():
        return []
    plates = []
    with csv_path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        for row in csv.reader(f):
            for cell in row:
                role = role_for_sheet_name(cell)
                if role:
                    nums = [_parse_fraction_inch(c) for c in row if _parse_fraction_inch(c) > 0]
                    plates.append(
                        {
                            "sheet_name": cell,
                            "role": role,
                            "thickness": nums[0] if nums else 0,
                            "width": nums[1] if len(nums) > 1 else 0,
                            "length": nums[2] if len(nums) > 2 else 0,
                            "source_sheet": csv_path.name,
                        }
                    )
                    break
    return _dedupe_plates(plates)


def _dedupe_plates(plates: list[dict]) -> list[dict]:
    seen: set[str] = set()
    out = []
    for p in plates:
        key = p["role"]
        if key in seen:
            continue
        seen.add(key)
        out.append(p)
    return out


def _size_close(a: float, b: float, tol: float = 0.06) -> bool:
    if a <= 0 or b <= 0:
        return False
    return abs(a - b) <= tol or abs(a - b) / max(a, b) <= 0.03


def _sorted_positive_dims(*vals: float) -> list[float]:
    return sorted(v for v in vals if v > 0)


def _match_plate_to_row_score(plate: dict, row: dict) -> float:
    """Score how well a steel-sheet plate row matches one XT component.

    Steel sheets often list W×L×T in varying column order and use fractions
    like 1 3/8 — compare sorted dimension triplets, not fixed T/W/L columns.
    """
    t = safe_float(row.get("Thickness"))
    w = safe_float(row.get("Width"))
    l = safe_float(row.get("Length"))
    pt = plate["thickness"]
    pw = plate["width"]
    pl = plate["length"]

    score = 0.0
    plate_sorted = _sorted_positive_dims(pt, pw, pl)
    row_sorted = _sorted_positive_dims(t, w, l)
    if len(plate_sorted) == 3 and len(row_sorted) == 3:
        if all(_size_close(a, b) for a, b in zip(plate_sorted, row_sorted)):
            score += 9

    # Best oriented assignment (handles steel column order differences)
    from itertools import permutations

    plate_vals = [v for v in (pt, pw, pl) if v > 0]
    if len(plate_vals) == 3:
        for ot, ow, ol in set(permutations(plate_vals)):
            oriented = 0.0
            if _size_close(t, ot):
                oriented += 3
            if _size_close(w, ow) or _size_close(w, ol):
                oriented += 2
            if _size_close(l, ol) or _size_close(l, ow):
                oriented += 2
            score = max(score, oriented)
    elif len(plate_vals) == 2:
        matched = 0
        for pv in plate_vals:
            for rv in (t, w, l):
                if _size_close(pv, rv):
                    matched += 1
                    break
        score = max(score, float(matched * 3))

    comp_name = _norm_name(row.get("Component", ""))
    role_phrase = plate["role"].replace("_", " ")
    if role_phrase in comp_name or _norm_name(plate["sheet_name"]) in comp_name:
        score += 4
    return score


def match_components(xt_rows: list[dict], plates: list[dict]) -> list[dict]:
    """Match CAD components to steel-sheet plate names by dimensions."""
    assignments: dict[str, str] = {}
    used_indices: set[str] = set()

    for plate in plates:
        role = plate["role"]
        if not role:
            continue
        best_idx = ""
        best_score = -1.0
        for row in xt_rows:
            idx = row.get("Index", "")
            if idx in used_indices:
                continue
            score = _match_plate_to_row_score(plate, row)
            if score > best_score:
                best_score = score
                best_idx = idx
        if best_idx and best_score >= 5:
            assignments[best_idx] = role
            used_indices.add(best_idx)

    # Shop-token hints for anything still unmatched (same tokens as qwen_classify_xt_csv)
    for row in xt_rows:
        idx = row.get("Index", "")
        if idx in assignments:
            continue
        name = _norm_name(row.get("Component", "")).upper().replace(" ", "-")
        token_map = (
            ("A-PLATE", "a_plate"),
            ("B-PLATE", "b_plate"),
            ("SC-RETAINER-PLATE", "sc_retainer_plate"),
            ("SC-BACKUP-PLATE", "sc_backup_plate"),
            ("CLAMP-PLATE", "bottom_clamp_plate"),
            ("EJ-BACKUP-PLATE", "bottom_ejector_plate"),
            ("EJ-RET-PLATE", "ejector_plate"),
            ("EJECTOR PLATE", "ejector_plate"),
            ("EJECTOR-PLATE", "ejector_plate"),
            ("LDR-PIN", "leader_pin"),
            ("LATCH-LOCK", "latch_lock"),
            ("RAIL", "rail"),
        )
        for token, role in token_map:
            if token in name:
                assignments[idx] = role
                break

    training = []
    for row in xt_rows:
        idx = row.get("Index", "")
        training.append(
            {
                "Index": idx,
                "Component": row.get("Component", ""),
                "Thickness": row.get("Thickness", ""),
                "Width": row.get("Width", ""),
                "Length": row.get("Length", ""),
                "CorrectRole": assignments.get(idx, ""),
                "Source": "steel_sheet" if idx in assignments else "",
            }
        )
    return training


def process_job(
    job_id: str,
    folder: Path,
    xt_csv: str = "",
    quote_sheet: str = "",
    steel_sheet: str = "",
) -> dict:
    xt_path = find_xt_csv(folder, xt_csv)
    steel_path = find_workbook(folder, STEEL_CANDIDATES, steel_sheet)
    quote_path = find_workbook(folder, QUOTE_CANDIDATES, quote_sheet)
    sheet_path = steel_path or quote_path

    if not sheet_path:
        return {"job_id": job_id, "status": "skipped", "reason": "no quote/steel sheet (.xls or .xlsx) found"}

    plates = read_sheet_plates(sheet_path)
    if not plates:
        return {
            "job_id": job_id,
            "status": "skipped",
            "reason": f"steel/quote sheet found ({sheet_path.name}) but no plate names parsed — check sheet format",
            "sheet": str(sheet_path),
        }

    if not xt_path:
        OUT_DIR.mkdir(parents=True, exist_ok=True)
        labels_path = OUT_DIR / f"{job_id}_STEEL_LABELS.json"
        labels_path.write_text(
            json.dumps({"job_id": job_id, "plates": plates, "sheet": str(sheet_path)}, indent=2),
            encoding="utf-8",
        )
        return {
            "job_id": job_id,
            "status": "ok_steel_only",
            "reason": "steel sheet OK — add XT_Export_CAD_Dimensions.csv (run macro once) for full training",
            "sheet": str(sheet_path),
            "plates_found": len(plates),
            "output": str(labels_path),
        }

    xt_rows = read_xt_rows(xt_path)
    training = match_components(xt_rows, plates)
    matched = sum(1 for r in training if r["CorrectRole"])

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_name = f"{job_id}_CORRECT_ME.csv"
    out_path = OUT_DIR / out_name
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=["Index", "Component", "Thickness", "Width", "Length", "CorrectRole", "Source"],
        )
        writer.writeheader()
        writer.writerows(training)

    return {
        "job_id": job_id,
        "status": "ok",
        "xt_csv": str(xt_path),
        "sheet": str(sheet_path),
        "plates_found": len(plates),
        "components_matched": matched,
        "total_components": len(xt_rows),
        "output": str(out_path),
        "accuracy_pct": round(100 * matched / max(len(xt_rows), 1), 1),
    }


def _scan_jobs_root(root: Path) -> list[dict]:
    rows = []
    if not root.exists():
        return rows
    for child in sorted(root.iterdir()):
        if child.is_dir():
            rows.append({"job_id": child.name, "job_folder": str(child)})
    return rows


def _load_manifest(path: Path) -> list[dict]:
    with path.open("r", newline="", encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def run_training(
    manifest_path: str | None = None,
    jobs_root: str | None = None,
    scan: bool = True,
) -> dict:
    entries: list[dict] = []
    if manifest_path:
        entries = _load_manifest(Path(manifest_path))
    elif scan and jobs_root:
        entries = _scan_jobs_root(Path(jobs_root))

    results = []
    for entry in entries:
        job_id = entry.get("job_id") or Path(entry.get("job_folder", "")).name
        folder = Path(entry.get("job_folder") or (Path(jobs_root or "") / job_id))
        results.append(
            process_job(
                job_id,
                folder,
                entry.get("xt_csv", ""),
                entry.get("quote_sheet", ""),
                entry.get("steel_sheet", ""),
            )
        )

    summary = {
        "jobs_processed": len(results),
        "jobs_ok": sum(1 for r in results if r.get("status") in ("ok", "ok_bms", "ok_steel_only")),
        "jobs_skipped": sum(1 for r in results if r.get("status") == "skipped"),
        "results": results,
        "output_dir": str(OUT_DIR),
    }
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    (DATA_DIR / "last_training_run.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    return summary


def status() -> dict:
    last_path = DATA_DIR / "last_training_run.json"
    if last_path.exists():
        try:
            return json.loads(last_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {"jobs_processed": 0, "output_dir": str(OUT_DIR)}


def main():
    parser = argparse.ArgumentParser(description="Train classifier from quote/steel sheets")
    parser.add_argument("--jobs-root", default="", help="Scan this folder for job subfolders")
    parser.add_argument("--manifest", default="", help="CSV manifest of jobs and file paths")
    parser.add_argument("--no-scan", action="store_true", help="Require --manifest")
    args = parser.parse_args()
    result = run_training(
        manifest_path=args.manifest or None,
        jobs_root=args.jobs_root or None,
        scan=not args.no_scan,
    )
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
