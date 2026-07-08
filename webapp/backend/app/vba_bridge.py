"""AI <-> Module6121 bridge.

Every time a job is classified (or its quote sheet is viewed), we write a
flat, VBA-friendly export of the AI-resolved part names/roles so Module6121
(or any macro) can read it and fill part names into the quoting workbook
without re-deriving geometry itself.

Two formats are written side by side:
  <job_id>_part_names.csv   -- trivial to parse from VBA with plain
                                Open/Line Input (no JSON library needed)
  <job_id>_part_names.json  -- for macros/tools that do have a JSON parser,
                                or for the optional HTTP bridge endpoint

CSV columns: Index,Component,Role,ResolvedName,Quote,Price,SecondaryParting

See webapp/vba/AI_Bridge_Import.bas for a ready-to-paste VBA subroutine that
reads this CSV and can be wired into Module6121's part-name-assignment logic.
"""
import csv
import json

from . import config
from .roles import role_label


def _resolved_name(role: str, role_counts: dict) -> str:
    label = role_label(role)
    role_counts[role] = role_counts.get(role, 0) + 1
    count = role_counts[role]
    # Only number roles that legitimately repeat per job (pins, bushings,
    # rails, latch locks). Plates stay singular ("A Plate", not "A Plate 1").
    singular_roles = {
        "a_plate", "b_plate", "top_clamp_plate", "bottom_clamp_plate",
        "sc_retainer_plate", "sc_backup_plate", "support_plate",
        "ejector_plate", "bottom_ejector_plate", "stripper_plate",
    }
    if role in singular_roles:
        return label
    return f"{label} {count}"


def write_bridge_files(job_id: str, quote_sheet: dict, job_analysis: dict) -> dict:
    config.VBA_BRIDGE_DIR.mkdir(parents=True, exist_ok=True)
    csv_path = config.VBA_BRIDGE_DIR / f"{job_id}_part_names.csv"
    json_path = config.VBA_BRIDGE_DIR / f"{job_id}_part_names.json"

    role_counts = {}
    rows = []
    for item in quote_sheet.get("line_items", []):
        role = item.get("role", "")
        rows.append(
            {
                "Index": item.get("index"),
                "Component": item.get("component"),
                "Role": role,
                "ResolvedName": _resolved_name(role, role_counts),
                "Quote": "TRUE" if item.get("quote") else "FALSE",
                "Price": item.get("price", 0.0),
                "SecondaryPartingLine": "TRUE"
                if job_analysis.get("sequenced_latch_lock_base") and role == "latch_lock"
                else "FALSE",
            }
        )

    with csv_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(
            f,
            fieldnames=[
                "Index", "Component", "Role", "ResolvedName", "Quote", "Price",
                "SecondaryPartingLine",
            ],
        )
        writer.writeheader()
        writer.writerows(rows)

    payload = {
        "job_id": job_id,
        "sequenced_latch_lock_base": bool(job_analysis.get("sequenced_latch_lock_base")),
        "parting_line": job_analysis.get("parting_line", ""),
        "total_price": quote_sheet.get("total_price", 0.0),
        "parts": rows,
    }
    json_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")

    return {
        "csv_path": str(csv_path),
        "json_path": str(json_path),
        "part_count": len(rows),
    }


def read_bridge_json(job_id: str):
    json_path = config.VBA_BRIDGE_DIR / f"{job_id}_part_names.json"
    if not json_path.exists():
        return None
    return json.loads(json_path.read_text(encoding="utf-8"))
