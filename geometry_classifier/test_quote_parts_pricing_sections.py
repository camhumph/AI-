"""Regression: Parts & Pricing surfaces steel + pull cores + purchased like the quote workbook."""

from __future__ import annotations

import csv
import tempfile
from pathlib import Path


def _write_csv(path: Path, header: list[str], rows: list[list]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        for row in rows:
            w.writerow(row)


def _write_quote_xlsx(path: Path) -> None:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QuoteWorksheet"

    # Summary block (labels in col A; values in C/D like the shop sheet)
    ws["A1"] = "Total Hours"
    ws["C1"] = 22.5
    ws["A2"] = "Total Price"
    ws["C2"] = 9085
    ws["D2"] = 9166
    ws["A3"] = "Commission"
    ws["B3"] = "6%"
    ws["C3"] = 545
    ws["D3"] = 550
    ws["A4"] = "Total Price"
    ws["C4"] = 9631
    ws["D4"] = 9716

    # BMS #2 steel block (rows 22-34): A=name C=qty D=T E=W F=L G=hours H=price
    steel = [
        (22, "TCP", 1, 1.650, 15.900, 18.400, 158, 298),
        (23, "BCP", 1, 1.650, 15.900, 18.400, 158, 298),
        (24, "Flipper Angle Plate Key", 0, 0, 0, 0, 0, 0),
        (31, "ID Holder", 1, 6.400, 6.750, 14.250, 182, 397),
        (32, "OD Holder", 1, 7.000, 7.500, 14.250, 221, 481),
        (33, "ID Pot", 1, 5.250, 6.000, 7.500, 71, 154),
        (34, "OD Pot", 1, 4.800, 6.100, 7.000, 61, 134),
    ]
    for row, name, qty, t, w, l, hours, price in steel:
        ws.cell(row, 1, name)
        ws.cell(row, 3, qty)
        ws.cell(row, 4, t)
        ws.cell(row, 5, w)
        ws.cell(row, 6, l)
        ws.cell(row, 7, hours)
        ws.cell(row, 8, price)

    wb.save(path)
    wb.close()


def test_quote_sheet_sections_match_workbook_breakdown(tmp_path: Path | None = None):
    from webapp.backend.app import config, pricing

    root = Path(tempfile.mkdtemp(prefix="cms_quote_sections_")) if tmp_path is None else tmp_path
    job_id = "C18699"
    job_dir = root / job_id
    job_dir.mkdir(parents=True)

    _write_quote_xlsx(job_dir / "BMS Quote Steel.xlsx")

    _write_csv(
        job_dir / "Pullcore Prices.csv",
        ["Pull Core / Key", "Qty", "Thickness", "Width", "Length", "Material", "Cu In", "Price USD"],
        [
            ["ID Pullcore Keys", 1, 0.954, 1.250, 3.800, "4140", 4.53, 398.77],
            ["OD Pullcore Keys", 1, 0.954, 1.250, 3.800, "4140", 4.53, 398.77],
            ["Odte Cam", 1, 1.750, 2.750, 5.760, "4140", 27.72, 2439.36],
            ["Idte Cam", 1, 1.750, 2.750, 5.760, "4140", 27.72, 2439.36],
            ["TOTAL", "", "", "", "", "", 64.50, 5676.26],
            ["RATE ($/in3)", "", "", "", "", "", "", 88],
        ],
    )

    _write_csv(
        job_dir / "Purchased Components Quote.csv",
        ["Component", "Vendor", "PartNumber", "Description", "QTY", "UnitPrice", "Extended"],
        [
            ["Leader Pin", "DME CO", "5213GL", "TOP LEADER PINS", 4, 25.46, 101.84],
            ["Bushing", "DME CO", "5503", "BOTTOM BUSHINGS", 4, 21.69, 86.76],
            ["Safety Strap", "PCS", "LSS-300", "SAFETY STRAPS", 1, 20.00, 20.00],
            ["Bushing", "MCMASTER-CARR", "6391K255", "BOTTOM EJ. BUSHING", 2, 21.69, 43.38],
            ["Insulation", "JACO", "HT200", "1/4 THK SPACER INSULATION (TOP & BTM)", 2, 0.00, 0.00],
            ["Insulation", "JACO", "HT200", "mold base (j-blk, cam, ht200)", 1, 0.00, 0.00],
            ["TOTAL", "", "", "", "", "", 251.98],
        ],
    )

    old_root = config.JOBS_ROOT
    config.JOBS_ROOT = root
    try:
        sheet = pricing.build_quote_sheet({"job_id": job_id, "base_type": "bms", "parts": []})
    finally:
        config.JOBS_ROOT = old_root

    steel = sheet["sections"]["steel"]
    pull = sheet["sections"]["pullcore"]
    purch = sheet["sections"]["purchased"]

    assert [s["component"] for s in steel] == [
        "TCP",
        "BCP",
        "ID Holder",
        "OD Holder",
        "ID Pot",
        "OD Pot",
    ], steel
    assert steel[0]["qty"] == 1
    assert steel[0]["thickness"] == 1.65
    assert steel[0]["hours"] == 158
    assert steel[0]["price"] == 298
    # Zero-qty flipper row must be skipped
    assert all("flipper" not in s["component"].lower() for s in steel)

    assert len(pull) == 4
    assert pull[0]["component"] == "ID Pullcore Keys"
    assert pull[0]["cu_in"] == 4.53
    assert pull[0]["price"] == 398.77
    assert abs(sum(p["price"] for p in pull) - 5676.26) < 0.05

    assert len(purch) == 6
    assert purch[0]["vendor"] == "DME CO"
    assert purch[0]["part_number"] == "5213GL"
    assert purch[0]["qty"] == 4
    assert purch[0]["unit_price"] == 25.46
    assert purch[0]["price"] == 101.84

    # Summary from workbook
    assert sheet["summary"]["total_hours"] == 22.5
    assert sheet["summary"]["commission_pct"] == 6
    assert sheet["summary"]["grand_total_finish"] == 9716
    assert sheet["total_price"] == 9716

    print("OK: quote sheet sections match workbook breakdown")


if __name__ == "__main__":
    test_quote_sheet_sections_match_workbook_breakdown()
