#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
cms_price_lookup.py  -  CMS purchased-component price updater.

WHY THIS EXISTS
---------------
The DME eStore (store.dme.net, "GenAlpha / Equip360" platform) renders its
prices with JavaScript - the price is NOT in the page HTML, so a plain HTTP
GET (or the SolidWorks VBA macro) cannot read it. This script drives a real
(headless) browser so the page's JavaScript runs and the LIST price appears,
then it parses the dollar amount and writes it into the price-list CSV that
the SolidWorks quoting macro already reads.

It runs LOGGED OUT on purpose, so you get the public list price (the high one)
rather than your discounted logged-in price.

ONE-TIME SETUP (on the quoting PC, in a cmd window)
---------------------------------------------------
    pip install playwright
    playwright install chromium

USAGE
-----
    # fill in any rows whose price is 0 (the normal run):
    python cms_price_lookup.py

    # re-price every DME row (refresh all):
    python cms_price_lookup.py --all

    # look up a single part and print it (no CSV write):
    python cms_price_lookup.py --part 5211GL

    # point at a specific CSV, or watch the browser work (debug):
    python cms_price_lookup.py --csv "C:\\Users\\lenovo\\Documents\\Trust\\Purchased Components Prices.csv"
    python cms_price_lookup.py --part 5211GL --headful

The CSV columns are:  Component,Vendor,PartNumber,Description,Unit,UnitPrice,Notes
Only DME rows are looked up automatically (McMaster / others block bots and are
left as-is; enter those by hand in the CSV).
"""

import argparse
import csv
import os
import re
import sys
import io

# ---- Configuration --------------------------------------------------------

# Where the macro looks for the price list (first existing path wins). Keep these
# in sync with the macro's FindPurchasedPriceFile() search order.
DEFAULT_CSV_CANDIDATES = [
    r"C:\Users\lenovo\Documents\Trust\Purchased Components Prices.csv",
    r"C:\Users\lenovo\Downloads\Purchased Components Prices.csv",
    r"C:\CMS_Local_Workspace\Purchased Components Prices.csv",
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "Purchased Components Prices.csv"),
]

# DME search page for a part number (logged out = list price).
DME_SEARCH_URL = "https://store.dme.net/search?keyword={pn}"

# Match $1,234.56  or  1234.56 USD  (price always has 2 decimals).
PRICE_RE = re.compile(r"\$\s*([0-9][0-9,]*\.[0-9]{2})|([0-9][0-9,]*\.[0-9]{2})\s*USD", re.I)

PAGE_TIMEOUT_MS = 30000     # max wait for navigation
PRICE_WAIT_MS = 12000       # max extra wait for the price JS to render


# ---- Price parsing --------------------------------------------------------

def parse_price(text):
    """Return the first plausible dollar value found in text, else None."""
    if not text:
        return None
    for m in PRICE_RE.finditer(text):
        raw = m.group(1) or m.group(2)
        raw = raw.replace(",", "")
        try:
            val = float(raw)
        except ValueError:
            continue
        if 0 < val < 1_000_000:
            return round(val, 2)
    return None


# ---- Browser lookup -------------------------------------------------------

def lookup_dme_price(page, part_no):
    """Render the DME page for part_no and return its list price (float) or None."""
    url = DME_SEARCH_URL.format(pn=part_no)
    try:
        page.goto(url, timeout=PAGE_TIMEOUT_MS, wait_until="domcontentloaded")
    except Exception as e:
        print(f"    nav error for {part_no}: {e}")
        return None

    # The price is injected by JS after load. Poll the rendered text for a price.
    import time
    deadline = time.time() + PRICE_WAIT_MS / 1000.0
    price = None
    while time.time() < deadline:
        try:
            body = page.inner_text("body")
        except Exception:
            body = ""
        price = parse_price(body)
        if price:
            break
        page.wait_for_timeout(750)

    # If the search landed on a results list rather than the item page, click the
    # first matching product and try once more.
    if price is None:
        try:
            link = page.query_selector(f"a:has-text('{part_no}')")
            if link:
                link.click()
                page.wait_for_timeout(2500)
                price = parse_price(page.inner_text("body"))
        except Exception:
            pass

    return price


def run_lookups(targets, headful=False):
    """targets: list of (part_no, vendor). Returns dict part_no -> price."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("ERROR: Playwright is not installed.\n"
              "  Run:  pip install playwright\n"
              "        playwright install chromium")
        sys.exit(2)

    results = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=not headful)
        ctx = browser.new_context(
            user_agent=("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                        "AppleWebKit/537.36 (KHTML, like Gecko) "
                        "Chrome/122.0 Safari/537.36"),
            locale="en-US",
        )
        page = ctx.new_page()
        for part_no, vendor in targets:
            v = (vendor or "").upper()
            if "DME" not in v:
                print(f"  skip {part_no}: vendor '{vendor}' not auto-looked-up")
                continue
            print(f"  looking up DME {part_no} ...", end=" ", flush=True)
            price = lookup_dme_price(page, part_no)
            if price:
                print(f"${price:.2f}")
                results[part_no] = price
            else:
                print("not found")
        browser.close()
    return results


# ---- CSV read / write (preserves comments + header) -----------------------

def find_csv(explicit):
    if explicit:
        return explicit
    for c in DEFAULT_CSV_CANDIDATES:
        if os.path.isfile(c):
            return c
    return None


def read_rows(path):
    """Return (lines, header_idx, data) where data is list of (idx, cols)."""
    with io.open(path, "r", encoding="utf-8-sig", newline="") as f:
        lines = f.read().splitlines()
    header_idx = None
    data = []
    for i, line in enumerate(lines):
        s = line.strip()
        if not s or s.startswith("#"):
            continue
        cols = next(csv.reader([line]))
        if header_idx is None and "Component" in line and "Price" in line:
            header_idx = i
            continue
        data.append((i, cols))
    return lines, header_idx, data


def write_price(lines, idx, cols, price):
    while len(cols) < 7:
        cols.append("")
    cols[5] = f"{price:.2f}"
    out = io.StringIO()
    csv.writer(out, lineterminator="").writerow(cols)
    lines[idx] = out.getvalue()


# ---- Main -----------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Update CMS purchased-component prices from DME (logged out).")
    ap.add_argument("--csv", help="path to the price-list CSV")
    ap.add_argument("--all", action="store_true", help="refresh every DME row, not just price==0")
    ap.add_argument("--part", help="look up a single part number and print it (no CSV write)")
    ap.add_argument("--headful", action="store_true", help="show the browser (debug)")
    args = ap.parse_args()

    if args.part:
        res = run_lookups([(args.part, "DME")], headful=args.headful)
        if res:
            print(f"\n{args.part} = ${res[args.part]:.2f}")
        else:
            print(f"\n{args.part}: no price found")
        return

    path = find_csv(args.csv)
    if not path:
        print("ERROR: could not find the price-list CSV. Pass --csv <path>.")
        sys.exit(1)
    print(f"Price list: {path}")

    lines, header_idx, data = read_rows(path)

    # Decide which rows to look up.
    targets = []
    row_by_part = {}
    for idx, cols in data:
        if len(cols) < 6:
            continue
        vendor = cols[1].strip()
        part_no = cols[2].strip()
        try:
            price = float((cols[5] or "0").replace("$", "").replace(",", "") or 0)
        except ValueError:
            price = 0.0
        if not part_no:
            continue
        if "DME" not in vendor.upper():
            continue
        if args.all or price <= 0:
            targets.append((part_no, vendor))
            row_by_part[part_no] = (idx, cols)

    if not targets:
        print("Nothing to look up (all DME rows already priced; use --all to refresh).")
        return

    print(f"Looking up {len(targets)} DME part(s)...")
    results = run_lookups(targets, headful=args.headful)

    updated = 0
    for part_no, price in results.items():
        idx, cols = row_by_part[part_no]
        write_price(lines, idx, cols, price)
        updated += 1

    if updated:
        with io.open(path, "w", encoding="utf-8", newline="") as f:
            f.write("\r\n".join(lines) + "\r\n")   # Windows CRLF so the VBA macro can read it
        print(f"\nUpdated {updated} price(s) in {path}")
    else:
        print("\nNo prices updated.")


if __name__ == "__main__":
    main()
