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
QUOTE_CANDIDATES = ("*quote*.xls*", "*Quote*.xls*", "documents/*quote*.xls*")
STEEL_CANDIDATES = ("*steel*.xls*", "*J000*.xls*", "documents/*steel*.xls*")


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


def read_sheet_plates(xlsx_path: Path) -> list[dict]:
    """Extract named plate rows from quote or steel Excel workbooks."""
    try:
        import openpyxl  # type: ignore
    except ImportError:
        return _read_sheet_plates_csv_fallback(xlsx_path)

    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    plates: list[dict] = []
    target_sheets = ("QuoteWorksheet", "Steel Order", "Machining Sheet", "Quote", "Steel")
    for sheet_name in wb.sheetnames:
        if sheet_name not in target_sheets and not any(
            k in sheet_name.lower() for k in ("quote", "steel", "machining")
        ):
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=200, values_only=True):
            if not row:
                continue
            cells = [str(c).strip() if c is not None else "" for c in row]
            name = ""
            t = w = l = 0.0
            for i, cell in enumerate(cells):
                role = role_for_sheet_name(cell)
                if role and not name:
                    name = cell
                    nums = [safe_float(c) for c in cells[i + 1 : i + 8] if _parse_fraction_inch(c) > 0]
                    if len(nums) >= 3:
                        t, w, l = nums[0], nums[1], nums[2]
                    elif len(nums) == 2:
                        w, l = nums[0], nums[1]
                    plates.append(
                        {
                            "sheet_name": name,
                            "role": role,
                            "thickness": t or _parse_fraction_inch(cells[i + 2] if i + 2 < len(cells) else 0),
                            "width": w,
                            "length": l,
                            "source_sheet": sheet_name,
                        }
                    )
                    break
            if not name:
                for cell in cells:
                    role = role_for_sheet_name(cell)
                    if role:
                        nums = [_parse_fraction_inch(c) for c in cells if _parse_fraction_inch(c) > 0]
                        plates.append(
                            {
                                "sheet_name": cell,
                                "role": role,
                                "thickness": nums[0] if len(nums) > 0 else 0,
                                "width": nums[1] if len(nums) > 1 else 0,
                                "length": nums[2] if len(nums) > 2 else 0,
                                "source_sheet": sheet_name,
                            }
                        )
                        break
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
                    nums = [safe_float(c) for c in row if safe_float(c) > 0]
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
            t = safe_float(row.get("Thickness"))
            w = safe_float(row.get("Width"))
            l = safe_float(row.get("Length"))
            score = 0
            if _size_close(t, plate["thickness"]):
                score += 3
            if _size_close(w, plate["width"]) or _size_close(w, plate["length"]):
                score += 2
            if _size_close(l, plate["length"]) or _size_close(l, plate["width"]):
                score += 2
            comp_name = _norm_name(row.get("Component", ""))
            if role.replace("_", " ") in comp_name or plate["sheet_name"].lower() in comp_name:
                score += 4
            if score > best_score:
                best_score = score
                best_idx = idx
        if best_idx and best_score >= 3:
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
    if not xt_path:
        return {"job_id": job_id, "status": "skipped", "reason": "no XT CSV"}

    steel_path = find_workbook(folder, STEEL_CANDIDATES, steel_sheet)
    quote_path = find_workbook(folder, QUOTE_CANDIDATES, quote_sheet)
    sheet_path = steel_path or quote_path
    if not sheet_path:
        return {"job_id": job_id, "status": "skipped", "reason": "no quote/steel sheet"}

    xt_rows = read_xt_rows(xt_path)
    plates = read_sheet_plates(sheet_path)
    if not plates:
        return {"job_id": job_id, "status": "skipped", "reason": "no plate names parsed from sheet"}

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
        "jobs_ok": sum(1 for r in results if r.get("status") == "ok"),
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
