"""Configurable pricing engine.

IMPORTANT: The unit rates shipped here are placeholders so the UI has a
working "Total Price" out of the box. They are not real CMS price-book
numbers (this repo has no pricing/rate data in it). Edit them from the
Settings page, or PUT /api/pricing so real CMS rates can be dropped in.

Pricing modes per role:
  flat        price = rate                          (e.g. a fixed hardware unit price)
  per_cuin    price = rate * (Thickness*Width*Length)  (rough material-volume proxy for plates)
  per_inch    price = rate * Length                   (rails/pins priced by length)
"""
import json

from . import config
from .roles import role_label

DEFAULT_RATES = {
    "top_clamp_plate": {"mode": "per_cuin", "rate": 0.85, "minimum": 120.0},
    "a_plate": {"mode": "per_cuin", "rate": 1.10, "minimum": 150.0},
    "b_plate": {"mode": "per_cuin", "rate": 1.10, "minimum": 150.0},
    "stripper_plate": {"mode": "per_cuin", "rate": 0.95, "minimum": 120.0},
    "sc_retainer_plate": {"mode": "per_cuin", "rate": 0.90, "minimum": 100.0},
    "sc_backup_plate": {"mode": "per_cuin", "rate": 0.90, "minimum": 100.0},
    "support_plate": {"mode": "per_cuin", "rate": 0.80, "minimum": 110.0},
    "bottom_clamp_plate": {"mode": "per_cuin", "rate": 0.85, "minimum": 120.0},
    "full_footprint_plate": {"mode": "per_cuin", "rate": 0.85, "minimum": 100.0},
    "rail": {"mode": "per_inch", "rate": 3.25, "minimum": 35.0},
    "rail_1": {"mode": "per_inch", "rate": 3.25, "minimum": 35.0},
    "rail_2": {"mode": "per_inch", "rate": 3.25, "minimum": 35.0},
    "pin_plate": {"mode": "per_cuin", "rate": 0.90, "minimum": 90.0},
    "ejector_plate": {"mode": "per_cuin", "rate": 0.95, "minimum": 90.0},
    "bottom_ejector_plate": {"mode": "per_cuin", "rate": 0.95, "minimum": 90.0},
    "ejector_retainer_plate": {"mode": "per_cuin", "rate": 0.95, "minimum": 90.0},
    "ejector_backup_plate": {"mode": "per_cuin", "rate": 0.95, "minimum": 90.0},
    "latch_lock": {"mode": "flat", "rate": 145.0, "minimum": 0.0},
    "leader_pin": {"mode": "flat", "rate": 28.0, "minimum": 0.0},
    "leader_pin_bushing": {"mode": "flat", "rate": 16.0, "minimum": 0.0},
    "guided_ejector_bushing": {"mode": "flat", "rate": 14.0, "minimum": 0.0},
    "return_pin": {"mode": "flat", "rate": 18.0, "minimum": 0.0},
    "ejector_pin": {"mode": "flat", "rate": 9.0, "minimum": 0.0},
    "support_pillar": {"mode": "flat", "rate": 22.0, "minimum": 0.0},
    "pullcore": {"mode": "flat", "rate": 210.0, "minimum": 0.0},
    "insert_or_core_detail": {"mode": "flat", "rate": 0.0, "minimum": 0.0},
    "hardware_other": {"mode": "flat", "rate": 6.5, "minimum": 0.0},
    "ignore": {"mode": "flat", "rate": 0.0, "minimum": 0.0},
}


def _safe_float(value, default=0.0):
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


# Map hardware roles to Component names in the shop's real
# "Purchased Components Prices.csv" (the same file Module6121 reads).
# When that CSV has a non-zero UnitPrice for the component, it overrides the
# placeholder flat rate for the role.
PURCHASED_CSV_ROLE_MAP = {
    "leader_pin": "Leader Pin",
    "leader_pin_bushing": "Bushing",
    "guided_ejector_bushing": "Ejector Bushing",
    "latch_lock": "Safety Strap",
    "support_pillar": "Support Pillar",
}


def _purchased_csv_prices() -> dict:
    """Component name -> first non-zero UnitPrice from the shop price CSV."""
    import csv as _csv

    prices = {}
    path = config.PURCHASED_PRICES_CSV
    if not path.exists():
        return prices
    try:
        with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            for row in _csv.DictReader(l for l in f if not l.lstrip().startswith("#")):
                comp = (row.get("Component") or "").strip()
                price = _safe_float(row.get("UnitPrice"))
                if comp and price > 0 and comp not in prices:
                    prices[comp] = price
    except Exception:
        pass
    return prices


def load_rates() -> dict:
    rates = dict(DEFAULT_RATES)

    # Real shop hardware prices beat placeholder flat rates.
    shop_prices = _purchased_csv_prices()
    for role, comp in PURCHASED_CSV_ROLE_MAP.items():
        if comp in shop_prices and role in rates:
            rates[role] = {**rates[role], "mode": "flat", "rate": shop_prices[comp]}

    # User-edited rates (Settings page) beat everything.
    if config.PRICING_CONFIG_PATH.exists():
        try:
            saved = json.loads(config.PRICING_CONFIG_PATH.read_text(encoding="utf-8"))
            rates.update(saved)
        except Exception:
            pass
    return rates


def save_rates(rates: dict) -> dict:
    config.PRICING_CONFIG_PATH.write_text(json.dumps(rates, indent=2), encoding="utf-8")
    return load_rates()


def price_row(row: dict, rates: dict) -> float:
    role = row.get("role", "")
    if not row.get("quote", False):
        return 0.0
    spec = rates.get(role, {"mode": "flat", "rate": 0.0, "minimum": 0.0})
    mode = spec.get("mode", "flat")
    rate = _safe_float(spec.get("rate", 0.0))
    minimum = _safe_float(spec.get("minimum", 0.0))

    t = _safe_float(row.get("Thickness") or row.get("t"))
    w = _safe_float(row.get("Width") or row.get("w"))
    l = _safe_float(row.get("Length") or row.get("l"))

    if mode == "per_cuin":
        price = rate * max(t * w * l, 0.0)
    elif mode == "per_inch":
        price = rate * max(l, 0.0)
    else:
        price = rate

    return round(max(price, minimum if price > 0 or minimum > 0 else 0.0), 2)


def build_quote_sheet(job: dict) -> dict:
    rates = load_rates()
    line_items = []
    total = 0.0
    for row in job.get("parts", []):
        price = price_row(row, rates)
        total += price
        line_items.append(
            {
                "index": row.get("index"),
                "component": row.get("Component") or row.get("component"),
                "role": row.get("role"),
                "role_label": row.get("role_label") or role_label(row.get("role", "")),
                "role_group": row.get("role_group"),
                "confidence": row.get("confidence") or row.get("Confidence"),
                "quote": bool(row.get("quote") or row.get("Quote")),
                "price": price,
            }
        )
    return {
        "job_id": job.get("job_id"),
        "line_items": line_items,
        "total_price": round(total, 2),
        "quoted_part_count": sum(1 for li in line_items if li["quote"]),
        "total_part_count": len(line_items),
    }
