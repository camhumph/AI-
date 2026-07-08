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

from . import config

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


def _find_workbook(job_dir: Path) -> Path | None:
    patterns = ("*quote*.xls*", "*Quote*.xls*", "*steel*.xls*", "*J000*.xls*")
    for sub in ("", "documents"):
        base = job_dir / sub if sub else job_dir
        if not base.exists():
            continue
        for pat in patterns:
            hits = sorted(base.glob(pat))
            if hits:
                return hits[0]
    return None


def read_sheet_dimensions(job_dir: Path) -> dict[str, dict]:
    """Parse quote/steel Excel for plate names and T/W/L (stock sizes on quote sheet)."""
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
                for cell in cells:
                    cn = _norm(cell)
                    if any(n in cn for n in names):
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
