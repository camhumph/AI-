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

from . import config, mold_groups, name_learning, plate_grades, plate_names, sheet_pricing
from .roles import bom_label, role_group, role_label


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
        # What the workbook says this plate's steel is, before any override. Read
        # from the steel sheet's STEEL TYPE column, or from the quote block the
        # row sits in. Blank falls back to the macro default below.
        "grade": row.get("grade") or "",
    }


def build_quote_sheet(job: dict) -> dict:
    job_id = job.get("job_id", "")
    job_dir = _job_dir(job_id)
    base_type = (job.get("base_type") or "").lower()

    shop_rows = sheet_pricing.load_shop_prices()
    job_purchased = sheet_pricing.load_job_purchased_quote(job_dir)

    # Renamed plates keep their stock dimensions -- and therefore their price --
    # by matching on the new name as well as the role's stock keywords.
    role_of_index = {
        str(p.get("index")): (p.get("role") or "") for p in job.get("parts", [])
    }
    sheet_dims = sheet_pricing.read_sheet_dimensions(
        job_dir, extra_aliases=plate_names.sheet_name_aliases(job_dir, role_of_index)
    )

    # The base type picks which QuoteWorksheet grade block the steel rows come
    # from: a standard base writes rows 6-15/38-43, a BMS pot base rows 22-34, and
    # rows 31-34 mean different plates in each. See _STD_QUOTE_STEEL_ROWS.
    steel_rows = sheet_pricing.load_steel_plate_lines(
        job_dir, job.get("base_type") or "standard"
    )
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

    # WHAT THE SHOP ALREADY TOLD US THIS PLATE IS CALLED.
    #
    # Applied before the per-job overrides below, so a rename typed on THIS job
    # still wins -- a global preference is the weaker signal of the two.
    #
    # Only swaps a name this app generated, never one the workbook chose for
    # itself: apply_to_generated compares against the role's built-in label and
    # leaves anything else alone. So learning that od_holder is "OD Holders" does
    # not touch a row the macro deliberately called something different.
    for item in steel_items:
        role = (item.get("role") or "").strip()
        if not role:
            continue
        generated = sheet_pricing._BMS_ROLE_LABELS.get(role) or role_label(role)
        item["component"] = name_learning.apply_to_generated(
            role, base_type, item.get("component") or "", generated
        )
        item["role_label"] = name_learning.label_for(role, base_type, item.get("role_label") or "")

    # Apply renames to the STEEL rows as well.
    #
    # Steel rows come out of the workbook keyed by role, while an override is
    # stored against a CAD index, so they need bridging by role. Without this a
    # rename would show on the CAD row and not on the priced steel row -- and
    # since the steel row is the one that survives de-duplication below, the
    # rename would appear to do nothing at all.
    override_by_role: dict[str, str] = {}
    # Role-keyed renames come straight from the store -- a steel row can be
    # renamed without any CAD part existing for it at all.
    for key, nm in plate_names.get_overrides(job_dir).items():
        if key.startswith(plate_names.ROLE_PREFIX):
            override_by_role[key[len(plate_names.ROLE_PREFIX):]] = nm
    for row in job.get("parts", []):
        nm = (row.get("name_override") or "").strip()
        r = (row.get("role") or "").strip()
        if nm and r and r not in override_by_role:
            override_by_role[r] = nm
    _all_ov = plate_names.get_overrides(job_dir)
    for item in steel_items:
        # Role first, then the displayed NAME. The name key is what makes a rename
        # land on a steel row whose workbook name sheet_pricing could not map to a
        # real role -- every one of those falls back to role "steel_plate", so
        # keying on the role would rename them all together.
        r = (item.get("role") or "").strip()
        nm = None
        if r not in plate_names.GENERIC_ROLES:
            nm = override_by_role.get(r)
        if not nm:
            nm = _all_ov.get(plate_names.name_key(item.get("component") or ""))
        if nm:
            item["role_label_original"] = item.get("component")
            item["component"] = nm
            item["name_override"] = nm

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
    # Roles the STEEL SHEET already prices. A plate that appears there must not
    # also appear as a CAD row, or every plate shows up twice -- once named
    # "B Plate" with a price, once named "2223605_B-PLATE" with none. That is the
    # "things are doubled up" the table was showing.
    #
    # The steel row wins because it comes from the workbook the shop actually
    # sends out, and it carries the price. The CAD row is the same physical plate
    # with a worse name and no price.
    #
    # Counted, not just flagged: a base with two A plates and two B plates (a
    # stack or a T-series) has TWO legitimate steel rows per role, so suppress at
    # most as many CAD rows as the sheet priced.
    # Matched on ROLE + THICKNESS, not role alone.
    #
    # Role alone suppressed whichever CAD row came first, which on a two-base job
    # is arbitrary: the steel sheet priced the 5.875" B plate (base 2) while the
    # surviving CAD row was the 3.879" one (base 1), and both then claimed to be
    # the same plate. Thickness is what actually identifies which plate a steel
    # row is, so that is the join.
    def _thk(v) -> float:
        try:
            return round(float(v or 0), 3)
        except (TypeError, ValueError):
            return 0.0

    steel_by_role_thk: dict[tuple, list] = {}
    for it in steel_items:
        r = (it.get("role") or "").strip()
        if r:
            steel_by_role_thk.setdefault((r, _thk(it.get("thickness"))), []).append(it)
    # Roles the sheet priced but with no usable thickness -- fall back to role.
    steel_role_budget: dict[str, int] = {}
    for (r, t), lst in steel_by_role_thk.items():
        if t == 0.0:
            steel_role_budget[r] = steel_role_budget.get(r, 0) + len(lst)

    classified_items: list[dict] = []
    suppressed_dupes = 0
    if base_type != "bms":
        for row in job.get("parts", []):
            role = row.get("role") or ""
            component = (row.get("Component") or row.get("component") or "").strip()
            label = (row.get("role_label") or role_label(role) or role or "").strip()
            # A user rename beats everything, including the CAD component name.
            #
            # This path computes `display` independently of get_job(), and it
            # normally prefers the CAD name over the role label. So without this
            # branch a rename would show up in the 3D gallery, Geometry and
            # Machining tabs but NOT in the parts table -- the one place the user
            # typed it. get_job() has already stamped the override onto the row.
            override = (row.get("name_override") or "").strip()

            # WHICH NAME TO SHOW.
            #
            # This used to be `override or component or label`, so the raw CAD
            # path won and the table read "2223605_B-PLATE" instead of "B Plate".
            # The shop's own name for the part is the useful one on a plate row.
            #
            # Hardware is the opposite: "LDR-PIN_2OD-X-7-3-4_PCS" IS the part
            # number, and collapsing it to "Leader Pin" would throw away the size
            # someone has to order. So plates prefer the role label, hardware
            # keeps the CAD name.
            # Decided from the ROLE, not the role group. An earlier version
            # tested group names against a hardcoded set that did not match what
            # roles.py actually returns ("Steel Plates / Mold Base", not "Mold
            # Base Plates"), so the branch never fired and every plate kept its
            # raw CAD path. Role suffixes cannot drift out of sync like that.
            is_plate = role.endswith("_plate") or role in {"rail", "rail_1", "rail_2"}
            # Prefer the name the shop's BOM uses. Its DESCRIPTION column is what
            # the floor works from, and for two roles it disagrees with the label
            # this app derived: EJ-RET-PLATE is an "Ejector Retainer Plate" and
            # EJ-BACKUP-PLATE an "Ejector Back-Up Plate". See roles.bom_label.
            bom = bom_label(role, component)
            if override:
                display = override
            elif bom:
                display = bom
            elif is_plate and label:
                display = label
            else:
                display = component or label
            # Drop blank junk rows that become "Other Hardware --"
            if not display or display in {"--", "-"}:
                continue

            # Already on the steel sheet with a price? Then this CAD row is the
            # same plate a second time. Matched on role AND thickness so the
            # right one is suppressed on a job with two bases.
            # The CAD row's own measured thickness. `priced` is not computed until
            # after this check, and using it here raised UnboundLocalError.
            row_thk = _thk(row.get("Thickness"))
            twin = steel_by_role_thk.get((role, row_thk))
            if role and twin:
                # Hand the steel row this part's mold group and CAD index, so it
                # can be labelled and renamed as the plate it actually is.
                st = twin.pop(0)
                st["_cad_index"] = row.get("index")
                if not twin:
                    steel_by_role_thk.pop((role, row_thk), None)
                suppressed_dupes += 1
                continue
            if role and steel_role_budget.get(role, 0) > 0:
                steel_role_budget[role] -= 1
                suppressed_dupes += 1
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
                # Present only when renamed, so the UI can mark the row and offer
                # to revert. `role` above is unchanged, which is what keeps the
                # price on the same quote row.
                "name_override": override or None,
                "role_label_original": row.get("role_label_original") or None,
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

    # ------------------------------------------------------------------
    # DROP CAD ROWS FOR PLATES THE STEEL SHEET ALREADY LISTS.
    #
    # Both sections describe the same steel from different sources: the steel
    # table from the quote/steel workbook, the classified list from CAD. Where a
    # plate is in both, it was shown TWICE under two different names -- on C18626
    # "Rails (x2)" sat in the steel table while a "Rails" section repeated one
    # "Rail", and "Bottom Ejector Plate" was repeated as "Ejector Back-Up Plate"
    # under "Ejector Assembly". Two rows, one plate, two names, and no way to tell
    # from the table that they are the same thing.
    #
    # The steel sheet wins because it is what the shop quotes from. Hardware,
    # inserts and core details are NOT dropped -- they never appear on the steel
    # sheet, so they are the only place those parts are visible at all.
    #
    # Matched on role, NOT on size: the steel row carries STOCK dimensions
    # (10.15 x 12.15) and the CAD row FINISHED (9.875 x 11.875), so a size test
    # would never fire and this whole pass would silently do nothing.
    _ROLE_ALIASES = {
        "ejector_backup_plate": "bottom_ejector_plate",
        "ejector_retainer_plate": "bottom_ejector_plate",
        "rail_1": "rail",
        "rail_2": "rail",
    }

    def _canon_role(r: str) -> str:
        r = (r or "").strip()
        return _ROLE_ALIASES.get(r, r)

    steel_roles = {_canon_role(it.get("role")) for it in steel_items}
    steel_roles.discard("")
    if steel_roles:
        dropped = [it for it in classified_items if _canon_role(it.get("role")) in steel_roles]
        if dropped:
            kept_ids = {id(it) for it in dropped}
            classified_items = [
                it for it in classified_items if _canon_role(it.get("role")) not in steel_roles
            ]
            line_items = [it for it in line_items if id(it) not in kept_ids]
            # Their prices came out of `total` with them; the steel row keeps its own.
            for it in dropped:
                p = float(it.get("price") or 0)
                if p > 0:
                    total -= p
                    csv_priced -= 1

    # Disambiguate plates that share a role.
    #
    # This base has TWO B plates (5.875" and 3.879") and two A plates. Once both
    # are shown by role label they both read "B Plate", which is worse than the
    # raw CAD name it replaced -- two identical-looking rows at different
    # thicknesses. Where a role repeats, the thickness goes in the name, because
    # that is what actually tells them apart on the floor.
    from collections import Counter

    # WHICH MOLD BASE. C18616 holds two complete bases, and the CAD says so with
    # a `_2` suffix. Naming both plates "B Plate" is worse than the raw CAD name
    # it replaced, so where a job has more than one base the plate says which.
    groups = mold_groups.detect_groups(job.get("parts", []))
    n_groups = mold_groups.group_count(groups)

    # Counted across BOTH sections. After de-duplication each role appears at
    # most once in the CAD list, so counting only that list found nothing --
    # meanwhile the table still showed "B Plate" in the steel section and a second
    # "B Plate" in the CAD section at a different thickness. The collision is
    # between sections, so that is where it has to be detected.
    role_counts = Counter(
        (it.get("role") or "") for it in steel_items + classified_items
    )
    for it in classified_items:
        if it.get("name_override"):
            continue
        r = it.get("role") or ""
        if role_counts.get(r, 0) < 2:
            continue
        # Only plates get a mold qualifier. Eight leader pins across two bases do
        # not need "(Mold 1)" on each -- they are hardware, bought by quantity.
        is_plate = r.endswith("_plate") or r in {"rail", "rail_1", "rail_2"}
        grp = groups.get(str(it.get("index"))) if is_plate else None
        if grp and n_groups > 1:
            it["component"] = mold_groups.label_with_group(it["component"], grp, n_groups)
            it["mold_group"] = grp
            continue
        # No marker to go on: fall back to the thickness, which is what actually
        # tells two same-role plates apart on the floor.
        thk = it.get("thickness")
        if thk:
            it["component"] = f"{it['component']} — {float(thk):g}\" thk"

    # Steel rows take the mold group of the CAD plate they were matched to above.
    #
    # An earlier version assumed steel rows were always base 1. On this job the
    # sheet priced the 5.875" B plate, which is base 2 — so the label said Mold 1
    # for a base-2 plate, sitting next to a CAD row also saying Mold 1. Taking the
    # group from the matched part means the label follows the steel, not a guess.
    if n_groups > 1:
        for it in steel_items:
            if it.get("name_override"):
                continue
            r = it.get("role") or ""
            if not (r.endswith("_plate") or r in {"rail", "rail_1", "rail_2"}):
                continue
            grp = groups.get(str(it.get("_cad_index")))
            if grp:
                it["component"] = mold_groups.label_with_group(it["component"], grp, n_groups)
                it["mold_group"] = grp
    # ------------------------------------------------------------------
    # STEEL GRADE, and the estimator's override of it.
    #
    # The macro now defaults a standard base to #2 throughout, matching both hand
    # quotes we have (C18597, C18619). Where a job is not #2 throughout, the
    # override in plate_grades.json says so per plate or per role, and it is
    # applied here so the Parts tab, the API and anything reading this sheet all
    # agree. `grade_source` records which it was, because a repriced plate the
    # estimator did not intend is worse than one they did.
    grade_overrides = plate_grades.get_overrides(job_dir)
    for it in steel_items + classified_items:
        base_grade = plate_grades.clean_grade(it.get("grade") or "") or "4140"
        ov = plate_grades.grade_for(
            grade_overrides,
            it.get("_cad_index") or it.get("index"),
            it.get("role") or "",
            it.get("component") or "",
        )
        it["grade"] = ov or base_grade
        it["grade_label"] = plate_grades.GRADE_LABELS.get(it["grade"], it["grade"])
        it["grade_source"] = "override" if ov else "default"
        it["grade_options"] = list(plate_grades.VALID_GRADES)

    for it in steel_items:
        it.pop("_cad_index", None)

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
        # Reported rather than silent: if this is ever wrong, someone needs to be
        # able to see that rows were hidden and why.
        "suppressed_duplicate_cad_rows": suppressed_dupes,
        "plate_name_overrides_applied": len(override_by_role),
    }
