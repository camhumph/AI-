"""Pricing engine — all prices from CSV / quote workbook, never guessed.

Sources (in priority order per line item):
  1. Job folder: Purchased Components Quote.csv (Module6121 output)
  2. Job folder: Pullcore Prices.csv (Module6121 output)
  3. Quote / steel Excel workbook (BMS #2 steel block + summary)
  4. Shop: Purchased Components Prices.csv (same file the macro reads)
  5. Dimensions: quote/steel Excel workbook when present, else CAD export

Parts & Pricing tab sections mirror the quote workbook:
  - Steel Plates / Mold Base
  - Pull Cores & Keys
  - Purchased Components
"""
from __future__ import annotations

import json
from pathlib import Path

from . import config, sheet_pricing
from .roles import role_group, role_label


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


def _line_from_section_row(
    row: dict,
    *,
    index: str,
    section: str,
    role_group_name: str,
) -> dict:
    role = row.get("role") or ""
    return {
        "index": index,
        "section": section,
        "component": row.get("component") or "",
        "role": role,
        "role_label": role_label(role) if role else (row.get("component") or ""),
        "role_group": role_group_name,
        "confidence": "HIGH",
        "quote": True,
        "price": float(row.get("price") or 0),
        "price_source": row.get("price_source") or "",
        "thickness": row.get("thickness"),
        "width": row.get("width"),
        "length": row.get("length"),
        "qty": row.get("qty"),
        "cu_in": row.get("cu_in"),
        "hours": row.get("hours"),
        "vendor": row.get("vendor") or "",
        "part_number": row.get("part_number") or "",
        "unit_price": row.get("unit_price"),
        "material": row.get("material") or "",
        "category": row.get("category") or "",
    }


def build_quote_sheet(job: dict) -> dict:
    job_id = job.get("job_id", "")
    job_dir = _job_dir(job_id)
    base_type = (job.get("base_type") or "").lower()

    shop_rows = sheet_pricing.load_shop_prices()
    job_purchased = sheet_pricing.load_job_purchased_quote(job_dir)
    sheet_dims = sheet_pricing.read_sheet_dimensions(job_dir)

    steel_rows = sheet_pricing.load_steel_plate_lines(job_dir)
    pullcore_rows = sheet_pricing.load_pullcore_lines(job_dir)
    purchased_rows = sheet_pricing.load_purchased_lines(job_dir)
    summary = sheet_pricing.load_quote_summary(job_dir)

    steel_items = [
        _line_from_section_row(
            r,
            index=f"S{i}",
            section="steel",
            role_group_name="Steel Plates / Mold Base",
        )
        for i, r in enumerate(steel_rows, start=1)
        if (r.get("component") or r.get("role"))
    ]
    # Guarantee a visible description even if the workbook Col A was blank.
    for item in steel_items:
        if not (item.get("component") or "").strip():
            item["component"] = item.get("role_label") or item.get("role") or "Steel Plate"

    pullcore_items = [
        _line_from_section_row(
            r,
            index=f"K{i}",
            section="pullcore",
            role_group_name="Pull Cores & Keys",
        )
        for i, r in enumerate(pullcore_rows, start=1)
        if (r.get("component") or "").strip()
    ]
    purchased_items = [
        _line_from_section_row(
            r,
            index=f"P{i}",
            section="purchased",
            role_group_name="Purchased Components",
        )
        for i, r in enumerate(purchased_rows, start=1)
        if (r.get("component") or "").strip()
    ]

    line_items: list[dict] = []
    total = 0.0
    csv_priced = 0
    missing = 0

    # Classified CAD parts (standard mold bases). Skip on BMS — steel/pullcore/purchased
    # sections are the source of truth and avoid empty AI A/B/rail rows.
    classified_items: list[dict] = []
    if base_type != "bms":
        for row in job.get("parts", []):
            role = row.get("role") or ""
            component = (row.get("Component") or row.get("component") or "").strip()
            label = (row.get("role_label") or role_label(role) or role or "").strip()
            display = component or label
            # Drop blank junk rows that become "Other Hardware --"
            if not display or display in {"--", "-"}:
                continue

            priced = sheet_pricing.price_for_part(row, shop_rows, job_purchased, sheet_dims)
            price = priced["price"]
            total += price
            if price > 0:
                csv_priced += 1
            elif row.get("quote") or row.get("Quote"):
                missing += 1

            item = {
                "index": row.get("index"),
                "section": "classified",
                "component": display,
                "role": role,
                "role_label": label or role_label(role),
                "role_group": row.get("role_group") or role_group(role),
                "confidence": row.get("confidence") or row.get("Confidence"),
                "quote": bool(row.get("quote") or row.get("Quote")),
                "price": price,
                "price_source": priced.get("price_source", ""),
                "thickness": priced.get("thickness"),
                "width": priced.get("width"),
                "length": priced.get("length"),
                "qty": sheet_pricing._safe_float(row.get("Qty") or row.get("QTY") or 1, 1.0),
            }
            classified_items.append(item)
            line_items.append(item)

    # Macro sections — always surface when present (BMS and standard).
    for item in steel_items + pullcore_items + purchased_items:
        price = float(item.get("price") or 0)
        total += price
        if price > 0:
            csv_priced += 1
        else:
            # Steel hours/price may be Excel-formula-only (data_only needs a prior Excel save).
            # Don't count steel blanks as "missing CSV price".
            if item.get("section") != "steel":
                missing += 1
        line_items.append(item)

    # Prefer workbook grand total when available (includes machining + commission).
    display_total = total
    if summary.get("grand_total_finish"):
        display_total = float(summary["grand_total_finish"])
    elif summary.get("grand_total_rough"):
        display_total = float(summary["grand_total_rough"])
    elif summary.get("total_price_finish"):
        display_total = float(summary["total_price_finish"])
    elif summary.get("total_price_rough"):
        display_total = float(summary["total_price_rough"])

    sources = []
    if steel_items:
        sources.append("quote/steel workbook")
    if pullcore_items:
        sources.append("Pullcore Prices.csv")
    if purchased_items:
        sources.append("Purchased Components Quote.csv")
    if classified_items:
        sources.append("Purchased Components Prices.csv")

    return {
        "job_id": job_id,
        "line_items": line_items,
        "sections": {
            "steel": steel_items,
            "pullcore": pullcore_items,
            "purchased": purchased_items,
            "classified": classified_items,
        },
        "steel_plates": steel_items,
        "pullcore_components": pullcore_items,
        "purchased_components": purchased_items,
        "summary": summary,
        "total_price": round(display_total, 2),
        "section_total_price": round(total, 2),
        "quoted_part_count": sum(1 for li in line_items if li.get("quote")),
        "total_part_count": len(line_items),
        "csv_priced_count": csv_priced,
        "missing_csv_price_count": missing,
        "pricing_source": " + ".join(sources) if sources else "Purchased Components Prices.csv",
        "shop_csv": str(config.PURCHASED_PRICES_CSV),
        "has_steel_sheet_dims": bool(sheet_dims) or bool(steel_items),
    }
