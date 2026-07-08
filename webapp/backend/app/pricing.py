"""Pricing engine — all prices from CSV files, never guessed.

Sources (in priority order per line item):
  1. Job folder: Purchased Components Quote.csv (Module6121 output)
  2. Shop: Purchased Components Prices.csv (same file the macro reads)
  3. Dimensions: quote/steel Excel workbook when present, else CAD export

Plate steel sizing on the website matches the quote/steel sheet when an Excel
workbook is in the job folder.
"""
from __future__ import annotations

import json
from pathlib import Path

from . import config, sheet_pricing
from .roles import role_label


def load_rates() -> dict:
    """Legacy Settings endpoint — returns shop CSV snapshot for display only."""
    shop = sheet_pricing.load_shop_prices()
    rates = {}
    for row in shop:
        comp = (row.get("Component") or "").strip()
        price = sheet_pricing._safe_float(row.get("UnitPrice"))
        if comp and price > 0:
            rates[comp] = {"mode": "flat", "rate": price, "minimum": 0.0, "source": "csv"}
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


def _job_dir(job_id: str) -> Path:
    safe = job_id.strip().replace("..", "").replace("/", "_")
    return config.JOBS_ROOT / safe


def build_quote_sheet(job: dict) -> dict:
    job_id = job.get("job_id", "")
    job_dir = _job_dir(job_id)

    shop_rows = sheet_pricing.load_shop_prices()
    job_purchased = sheet_pricing.load_job_purchased_quote(job_dir)
    sheet_dims = sheet_pricing.read_sheet_dimensions(job_dir)

    line_items = []
    total = 0.0
    csv_priced = 0
    missing = 0

    for row in job.get("parts", []):
        priced = sheet_pricing.price_for_part(row, shop_rows, job_purchased, sheet_dims)
        price = priced["price"]
        total += price
        if price > 0:
            csv_priced += 1
        elif row.get("quote") or row.get("Quote"):
            missing += 1

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
                "price_source": priced.get("price_source", ""),
                "thickness": priced.get("thickness"),
                "width": priced.get("width"),
                "length": priced.get("length"),
            }
        )

    return {
        "job_id": job_id,
        "line_items": line_items,
        "total_price": round(total, 2),
        "quoted_part_count": sum(1 for li in line_items if li["quote"]),
        "total_part_count": len(line_items),
        "csv_priced_count": csv_priced,
        "missing_csv_price_count": missing,
        "pricing_source": "Purchased Components Prices.csv"
        + (" + job Purchased Components Quote.csv" if job_purchased else ""),
        "shop_csv": str(config.PURCHASED_PRICES_CSV),
        "has_steel_sheet_dims": bool(sheet_dims),
    }
