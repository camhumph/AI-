import csv
import json
from pathlib import Path


ROOTS = [
    Path(r"C:\CMS_Local_Workspace"),
    Path(r"C:\CMS_AI"),
]

OUT_DIR = Path(r"C:\CMS_AI\geometry_classifier\data")
OUT_DIR.mkdir(parents=True, exist_ok=True)

RAW_OUT = OUT_DIR / "all_xt_export_rows.csv"
JOB_OUT = OUT_DIR / "job_geometry_summary.csv"
JSONL_OUT = OUT_DIR / "geometry_ai_examples.jsonl"


def safe_float(value):
    try:
        return float(str(value).strip())
    except Exception:
        return 0.0


def job_name_from_path(path):
    parts = path.parts
    for part in reversed(parts):
        up = part.upper()
        if up.startswith("C") and up[1:].isdigit():
            return part
        if up.startswith("J") and up[1:].isdigit():
            return part
        if up.startswith("CMS_ACTIVE_"):
            return part
    return path.parent.name


def read_rows(path):
    rows = []
    with path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        reader = csv.DictReader(f)
        for row in reader:
            rows.append(row)
    return rows


def row_dims(row):
    t = safe_float(row.get("Thickness"))
    w = safe_float(row.get("Width"))
    l = safe_float(row.get("Length"))
    vol = safe_float(row.get("BBoxVolume_cuin"))
    cx = safe_float(row.get("CenterX"))
    cy = safe_float(row.get("CenterY"))
    cz = safe_float(row.get("CenterZ"))
    return t, w, l, vol, cx, cy, cz


def is_round_like(t, w, l):
    dims = sorted([t, w, l])
    if dims[0] <= 0:
        return False
    dia = (dims[0] + dims[1]) / 2.0
    return abs(dims[0] - dims[1]) <= max(0.08, dia * 0.10)


def summarize_job(rows):
    usable = []
    for row in rows:
        t, w, l, vol, cx, cy, cz = row_dims(row)
        if t > 0 and w > 0 and l > 0:
            usable.append((row, t, w, l, vol, cx, cy, cz))
    if not usable:
        return None

    base_foot = max(w * l for _, t, w, l, vol, cx, cy, cz in usable)
    base_row = max(usable, key=lambda x: x[2] * x[3])
    base_w, base_l = base_row[2], base_row[3]

    full = []
    rails = []
    ejector = []
    round_parts = []
    for row, t, w, l, vol, cx, cy, cz in usable:
        fp = w * l
        if t >= 0.4 and fp >= 0.82 * base_foot:
            full.append((row, t, w, l, cy))
        elif (
            t >= 0.5
            and l >= 0.60 * base_l
            and w <= 0.65 * base_w
            and max(w / t, t / w) >= 1.35
            and l / max(w, 0.001) >= 2.25
        ):
            rails.append((row, t, w, l, cy))
        elif t >= 0.4 and l >= 0.60 * base_l and w >= 0.35 * base_w and fp >= 0.15 * base_foot:
            ejector.append((row, t, w, l, cy))

        if is_round_like(t, w, l):
            round_parts.append((row, t, w, l, vol))

    # Heuristic for standard/non-BMS mold bases: several same-footprint plates,
    # usually with rails/ejector stack/leader hardware.
    likely_standard = len(full) >= 3 and (len(rails) >= 1 or len(ejector) >= 1 or len(round_parts) >= 4)

    return {
        "row_count": len(usable),
        "base_width": round(base_w, 3),
        "base_length": round(base_l, 3),
        "full_plate_count": len(full),
        "rail_count": len(rails),
        "ejector_stack_count": len(ejector),
        "round_part_count": len(round_parts),
        "likely_standard_non_bms": likely_standard,
    }


def main():
    csv_paths = []
    for root in ROOTS:
        if root.exists():
            csv_paths.extend(root.rglob("XT_Export_CAD_Dimensions.csv"))

    seen = set()
    unique_paths = []
    for path in csv_paths:
        key = str(path).lower()
        if key not in seen and path.stat().st_size > 250:
            seen.add(key)
            unique_paths.append(path)

    raw_fields = [
        "SourcePath",
        "Job",
        "LikelyStandardNonBms",
        "Index",
        "Component",
        "Qty",
        "Thickness",
        "Width",
        "Length",
        "BBoxVolume_cuin",
        "Mass_or_Vol",
        "CenterX",
        "CenterY",
        "CenterZ",
    ]

    summaries = []
    with RAW_OUT.open("w", newline="", encoding="utf-8") as raw_f, JSONL_OUT.open("w", encoding="utf-8") as jsonl_f:
        raw_writer = csv.DictWriter(raw_f, fieldnames=raw_fields)
        raw_writer.writeheader()

        for path in sorted(unique_paths):
            try:
                rows = read_rows(path)
            except Exception:
                continue
            summary = summarize_job(rows)
            if not summary:
                continue

            job = job_name_from_path(path)
            summaries.append({"Job": job, "SourcePath": str(path), **summary})

            for row in rows:
                out = {field: "" for field in raw_fields}
                out["SourcePath"] = str(path)
                out["Job"] = job
                out["LikelyStandardNonBms"] = "TRUE" if summary["likely_standard_non_bms"] else "FALSE"
                for key in raw_fields:
                    if key in row:
                        out[key] = row.get(key, "")
                raw_writer.writerow(out)

            if summary["likely_standard_non_bms"]:
                compact_rows = []
                for row in rows[:120]:
                    t, w, l, vol, cx, cy, cz = row_dims(row)
                    compact_rows.append(
                        {
                            "index": row.get("Index", ""),
                            "component": row.get("Component", ""),
                            "t": t,
                            "w": w,
                            "l": l,
                            "cx": cx,
                            "cy": cy,
                            "cz": cz,
                            "volume": vol,
                        }
                    )
                prompt = (
                    "Classify these mold-base CAD parts. Exact shop-standard name tokens "
                    "(A-PLATE, B-PLATE, SC-RETAINER, SC-BACKUP, EJ-RET, EJ-BACKUP, RAIL, "
                    "LDR-PIN, LBB, PLC75/LATCH-LOCK/SAFETY-STRAP) are strong anchor evidence; "
                    "only fall back to geometry when names are generic or missing. Use "
                    "dimensions, center positions, stack order, full-footprint plates, rails, "
                    "ejector stack, leader pins, bushings, support pillars, and parting-line "
                    "logic. Anchor stack orientation from rails/ejector stack first; leader "
                    "pins only decide orientation when rails/ejector plates are missing."
                )
                jsonl_f.write(
                    json.dumps(
                        {
                            "job": job,
                            "source": str(path),
                            "summary": summary,
                            "prompt": prompt,
                            "rows": compact_rows,
                        },
                        ensure_ascii=False,
                    )
                    + "\n"
                )

    with JOB_OUT.open("w", newline="", encoding="utf-8") as f:
        fields = [
            "Job",
            "SourcePath",
            "row_count",
            "base_width",
            "base_length",
            "full_plate_count",
            "rail_count",
            "ejector_stack_count",
            "round_part_count",
            "likely_standard_non_bms",
        ]
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        writer.writerows(summaries)

    print(f"Found {len(unique_paths)} XT export CSV files")
    print(f"Wrote {RAW_OUT}")
    print(f"Wrote {JOB_OUT}")
    print(f"Wrote {JSONL_OUT}")
    print(f"Likely standard/non-BMS jobs: {sum(1 for s in summaries if s['likely_standard_non_bms'])}")


if __name__ == "__main__":
    main()
