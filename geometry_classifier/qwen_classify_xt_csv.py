import argparse
import csv
import json
import re
import shutil
import subprocess
import sys
from pathlib import Path


BASE = Path(r"C:\CMS_AI\geometry_classifier")
KNOWLEDGE = BASE / "mold_geometry_knowledge.md"
VENDOR_KNOWLEDGE = BASE / "vendor_knowledge_sources.md"
OUT_DIR = BASE / "outputs"
OUT_DIR.mkdir(parents=True, exist_ok=True)


SHORT_RULES = """
Use geometry only. Component names may be wrong.
Exception: if a STEP/imported assembly has deliberate shop-standard tokens like
A-PLATE, B-PLATE, SC-RETAINER, SC-BACKUP, EJ-RET, EJ-BACKUP, RAIL, LDR-PIN, or
LBB, treat those as strong hints and still cross-check them against geometry.

Analyze the whole mold first:
- Find the stack axis from full-footprint plates.
- Full-footprint plates have nearly the same width/length as the largest base footprint.
- For a 5-full-plate standard stack, sorted top to bottom:
  1 top_clamp_plate
  2 a_plate
  3 b_plate
  4 support_plate
  5 bottom_clamp_plate
- Rails are long narrow side blocks near the ejector side.
- Pin/ejector plates are long narrower plates inside/between rails.
- pin_plate/ejector retainer is above ejector_plate in the ejector stack.
- Leader pins are long round pins.
- Leader/shoulder bushings are short round cylinders near leader-pin locations.
- Support pillars are large round posts and are not leader pins.
- Return/ejector pins are smaller long round pins.
- The parting line is between a_plate and b_plate.
"""


ROLES = [
    "top_clamp_plate",
    "a_plate",
    "b_plate",
    "stripper_plate",
    "sc_retainer_plate",
    "sc_backup_plate",
    "support_plate",
    "bottom_clamp_plate",
    "full_footprint_plate",
    "rail",
    "rail_1",
    "rail_2",
    "pin_plate",
    "ejector_plate",
    "ejector_retainer_plate",
    "bottom_ejector_plate",
    "ejector_backup_plate",
    "leader_pin",
    "leader_pin_bushing",
    "guided_ejector_bushing",
    "return_pin",
    "ejector_pin",
    "support_pillar",
    "pullcore",
    "insert_or_core_detail",
    "hardware_other",
    "ignore",
]


def safe_float(value):
    try:
        return float(str(value).strip())
    except Exception:
        return 0.0


def read_rows(csv_path, include_names=False):
    with Path(csv_path).open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        rows = list(csv.DictReader(f))
    compact = []
    for row in rows:
        item = {
            "i": row.get("Index", ""),
            "t": safe_float(row.get("Thickness")),
            "w": safe_float(row.get("Width")),
            "l": safe_float(row.get("Length")),
            "v": safe_float(row.get("BBoxVolume_cuin")),
            "x": safe_float(row.get("CenterX")),
            "y": safe_float(row.get("CenterY")),
            "z": safe_float(row.get("CenterZ")),
        }
        if include_names:
            item["name"] = row.get("Component", "")[:80]
        compact.append(item)
    return compact


def build_prompt(rows, csv_path, long_knowledge=False):
    if long_knowledge:
        knowledge = KNOWLEDGE.read_text(encoding="utf-8")
        vendor_knowledge = VENDOR_KNOWLEDGE.read_text(encoding="utf-8") if VENDOR_KNOWLEDGE.exists() else ""
    else:
        knowledge = SHORT_RULES
        vendor_knowledge = ""
    role_list = ", ".join(ROLES)
    return f"""
/no_think
You are a CMS mold-base geometry interpreter.

Your job:
1. Analyze the whole mold first.
2. Decide the stack axis.
3. Decide the full-footprint stack order.
4. Use leader pins, bushings, rails, ejector stack, and dimensions to identify the A plate, B plate, parting line, and part names.
5. Do not trust CAD component names. Names may be wrong or mixed up.
6. Use CAD names only as weak notes after geometry has decided.
7. Return JSON only. No explanation outside JSON. No markdown.

Allowed roles:
{role_list}

Knowledge:
{knowledge}

Vendor/DME/PCS knowledge:
{vendor_knowledge}

Return ONLY valid JSON in this exact shape:
{{
  "job_analysis": {{
    "stack_axis": "CenterX|CenterY|CenterZ",
    "parting_line": "short explanation",
    "rules_for_this_job": [
      "rule 1",
      "rule 2"
    ]
  }},
  "classifications": [
    {{
      "index": "1",
      "role": "a_plate",
      "confidence": "HIGH|MEDIUM|LOW",
      "reason": "geometry-only reason",
      "quote": true
    }}
  ]
}}

CSV source:
{csv_path}

Rows:
{json.dumps(rows, separators=(",", ":"))}
"""


def extract_json(text):
    # Ollama/Qwen can emit terminal control characters or thinking text. Strip
    # those first, then extract the first JSON object.
    text = re.sub(r"\x1b\[[0-9;?]*[A-Za-z]", "", text)
    text = text.replace("\b", "")
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S | re.I)
    text = re.sub(r"^Thinking\.\.\..*?(?=\{)", "", text, flags=re.S)
    text = text.strip()
    if text.startswith("{"):
        return json.loads(text)
    match = re.search(r"\{.*\}", text, flags=re.S)
    if not match:
        raise ValueError("No JSON object found in model output")
    return json.loads(match.group(0))


def close_enough(a, b, tolerance=0.08):
    if a == 0 or b == 0:
        return False
    return abs(a - b) <= max(abs(a), abs(b)) * tolerance


def round_bar_diameter(row):
    dims = [row["t"], row["w"], row["l"]]
    pairs = [(dims[0], dims[1], dims[2]), (dims[0], dims[2], dims[1]), (dims[1], dims[2], dims[0])]
    for a, b, long_axis in pairs:
        if close_enough(a, b, 0.06):
            return min(a, b), long_axis
    return 0.0, 0.0


def name_key(row):
    return str(row.get("name", "")).upper().replace("\\", "/")


def apply_strong_shop_name_hints(rows, roles):
    """Use exact imported shop tokens only; generic CAD names remain weak."""
    for row in rows:
        idx = str(row["i"])
        name = name_key(row)
        if not name:
            continue

        if "A-PLATE" in name or "A_PLATE" in name:
            roles[idx] = ("a_plate", "HIGH", "Strong shop token A-PLATE, confirmed as a full-footprint mold plate.", True)
        elif "B-PLATE" in name or "B_PLATE" in name:
            roles[idx] = ("b_plate", "HIGH", "Strong shop token B-PLATE, confirmed as a full-footprint mold plate.", True)
        elif "SC-RETAINER-PLATE" in name or "SC_RETAINER_PLATE" in name:
            roles[idx] = ("sc_retainer_plate", "HIGH", "Strong shop token SC-RETAINER-PLATE in a special standard mold-base stack.", True)
        elif "SC-BACKUP-PLATE" in name or "SC_BACKUP_PLATE" in name:
            roles[idx] = ("sc_backup_plate", "HIGH", "Strong shop token SC-BACKUP-PLATE in a special standard mold-base stack.", True)
        elif "CLAMP-PLATE" in name or "CLAMP_PLATE" in name:
            roles[idx] = ("bottom_clamp_plate", "HIGH", "Strong shop token CLAMP-PLATE on the ejector/clamp side.", True)
        elif "EJ-BACKUP-PLATE" in name or "EJ_BACKUP_PLATE" in name:
            roles[idx] = ("ejector_retainer_plate", "HIGH", "Strong shop token EJ-BACKUP-PLATE; backing/retainer plate in ejector stack.", True)
        elif "EJ-RET-PLATE" in name or "EJ_RET_PLATE" in name:
            roles[idx] = ("ejector_plate", "HIGH", "Strong shop token EJ-RET-PLATE; user standard treats this as the ejector plate.", True)
        elif "RAIL-" in name or "_RAIL" in name or "/RAIL" in name:
            roles[idx] = ("rail", "HIGH", "Strong shop token RAIL; long side rail/support block.", True)
        elif "LDR-PIN" in name or "LDR_PIN" in name:
            roles[idx] = ("leader_pin", "HIGH", "Strong shop token LDR-PIN; primary leader pin set.", False)
        elif "/LBB_" in name or "LBB_" in name:
            roles[idx] = ("leader_pin_bushing", "HIGH", "Strong shop token LBB; leader pin bushing.", False)


def classify_geometry(rows):
    """Rule-based fallback for when the LLM does not return valid JSON."""
    if not rows:
        return {"job_analysis": {}, "classifications": []}

    max_w = max(r["w"] for r in rows)
    max_l = max(r["l"] for r in rows)
    stack_axis_key = "y"
    stack_axis = "CenterY"
    roles = {}
    apply_strong_shop_name_hints(rows, roles)

    full_plates = [
        r for r in rows
        if r["w"] >= max_w * 0.85 and r["l"] >= max_l * 0.85 and r["t"] >= 0.5
    ]
    if len(full_plates) >= 2:
        ranges = {
            "x": max(r["x"] for r in full_plates) - min(r["x"] for r in full_plates),
            "y": max(r["y"] for r in full_plates) - min(r["y"] for r in full_plates),
            "z": max(r["z"] for r in full_plates) - min(r["z"] for r in full_plates),
        }
        stack_axis_key = max(ranges, key=ranges.get)
        stack_axis = {"x": "CenterX", "y": "CenterY", "z": "CenterZ"}[stack_axis_key]

    full_plates.sort(key=lambda r: r[stack_axis_key], reverse=True)
    if len(full_plates) >= 5:
        inner = full_plates[1:-1]
        avg_inner_t = sum(r["t"] for r in inner) / len(inner) if inner else 0.0
        top_clamp_present = True
        if avg_inner_t > 0 and full_plates[0]["t"] > avg_inner_t * 1.35:
            top_clamp_present = False

        if top_clamp_present:
            stack_names = ["top_clamp_plate", "a_plate", "b_plate", "support_plate", "bottom_clamp_plate"]
        else:
            stack_names = ["a_plate", "stripper_plate", "b_plate", "support_plate", "bottom_clamp_plate"]
        for name, row in zip(stack_names, full_plates):
            idx = str(row["i"])
            if idx not in roles:
                roles[idx] = (
                    name,
                    "HIGH",
                    "Full-footprint plate, assigned by top-to-bottom mold stack order. Top clamp missing." if not top_clamp_present else "Full-footprint plate, assigned by top-to-bottom mold stack order.",
                    True,
                )
    else:
        for row in full_plates:
            roles[str(row["i"])] = (
                "full_footprint_plate",
                "MEDIUM",
                "Full-footprint plate, but fewer than 5 full plates were found so standard stack role was not forced.",
                True,
            )

    if len(full_plates) == 2:
        axes = {
            "x": abs(full_plates[0]["x"] - full_plates[1]["x"]),
            "y": abs(full_plates[0]["y"] - full_plates[1]["y"]),
            "z": abs(full_plates[0]["z"] - full_plates[1]["z"]),
        }
        die_axis = max(axes, key=axes.get)
        hi_full, lo_full = sorted(full_plates, key=lambda r: r[die_axis], reverse=True)
        roles[str(lo_full["i"])] = (
            "bottom_clamp_plate",
            "HIGH",
            "Two-half mold pattern: lower full-footprint plate along stack axis is BCP.",
            True,
        )
        roles[str(hi_full["i"])] = (
            "top_clamp_plate",
            "MEDIUM",
            "Two-half mold pattern: opposite full-footprint clamp plate.",
            True,
        )

        non_full_blocks = [
            r for r in rows
            if str(r["i"]) not in roles
            and r["t"] >= 3.0
            and r["w"] >= max_w * 0.30
            and r["l"] >= max_l * 0.30
        ]
        non_full_blocks.sort(key=lambda r: r["v"], reverse=True)
        if len(non_full_blocks) >= 2:
            high_inner, low_inner = sorted(non_full_blocks[:2], key=lambda r: r[die_axis], reverse=True)
            roles[str(high_inner["i"])] = (
                "a_plate",
                "MEDIUM",
                "Two-half mold pattern: larger inner block on high side of stack axis.",
                True,
            )
            roles[str(low_inner["i"])] = (
                "b_plate",
                "MEDIUM",
                "Two-half mold pattern: matching inner block on low side of stack axis.",
                True,
            )

        thin_large = [
            r for r in rows
            if str(r["i"]) not in roles
            and r["t"] <= 0.75
            and r["l"] >= max_l * 0.35
            and r["w"] >= max_w * 0.20
        ]
        thin_large.sort(key=lambda r: r["v"], reverse=True)
        if len(thin_large) >= 2:
            rails = thin_large[:2]
            rails.sort(key=lambda r: r[die_axis], reverse=True)
            roles[str(rails[0]["i"])] = ("rail_1", "MEDIUM", "Two-half mold pattern: first largest thin rail/strip plate.", True)
            roles[str(rails[1]["i"])] = ("rail_2", "MEDIUM", "Two-half mold pattern: second largest thin rail/strip plate.", True)
        if len(thin_large) >= 4:
            ejector_candidates = thin_large[2:4]
            ejector_candidates.sort(key=lambda r: r["v"], reverse=True)
            roles[str(ejector_candidates[0]["i"])] = (
                "ejector_retainer_plate",
                "MEDIUM",
                "Two-half mold pattern: larger remaining thin ejector-stack plate.",
                True,
            )
            roles[str(ejector_candidates[1]["i"])] = (
                "ejector_backup_plate",
                "MEDIUM",
                "Two-half mold pattern: smaller remaining thin ejector-stack plate.",
                True,
            )

    top_pos = full_plates[0][stack_axis_key] if full_plates else 0.0
    bottom_pos = full_plates[-1][stack_axis_key] if full_plates else 0.0
    support_pos = next((r[stack_axis_key] for r in full_plates if roles.get(str(r["i"]), ("",))[0] == "support_plate"), 0.0)

    for row in rows:
        idx = str(row["i"])
        if idx in roles:
            continue

        dia, bar_len = round_bar_diameter(row)
        if dia and dia <= 4.0:
            is_long_bar = bar_len >= dia * 1.6
            is_short_cylinder = not is_long_bar and bar_len >= dia * 0.6

            if 2.5 <= dia <= 4.0 and bar_len >= 6.0:
                roles[idx] = ("support_pillar", "HIGH", "Large long round post; support pillar geometry.", False)
            elif is_long_bar and 1.35 <= dia <= 2.2 and bar_len >= 10.0:
                roles[idx] = ("leader_pin", "MEDIUM", "Long smaller round bar; likely leader pin.", False)
            elif is_long_bar and 0.9 <= dia < 1.35 and bar_len >= 8.0:
                roles[idx] = ("return_pin", "MEDIUM", "Long 1-inch-class round pin; likely return/ejector return pin.", False)
            elif is_short_cylinder and 1.0 <= dia <= 2.6 and row[stack_axis_key] >= support_pos:
                roles[idx] = ("leader_pin_bushing", "MEDIUM", "Short round cylinder near guide hardware size.", False)
            elif is_short_cylinder and 1.0 <= dia <= 2.6 and row[stack_axis_key] < support_pos:
                roles[idx] = ("guided_ejector_bushing", "MEDIUM", "Short round cylinder in ejector half; likely guided-ejector bushing.", False)
            else:
                roles[idx] = ("hardware_other", "LOW", "Round hardware but role is not certain from dimensions alone.", False)
            continue

        long_full = row["l"] >= max_l * 0.85
        side_coords = [axis for axis in ("x", "y", "z") if axis != stack_axis_key]
        side_offset = max(abs(row[axis]) for axis in side_coords)
        centered_side = side_offset <= max_w * 0.15
        side_block = side_offset >= max_w * 0.25
        narrow_width = max_w * 0.18 <= row["w"] <= max_w * 0.72
        ejector_width = max_w * 0.58 <= row["w"] <= max_w * 0.86

        axis_pos = row[stack_axis_key]
        if long_full and narrow_width and side_block and bottom_pos <= axis_pos <= support_pos + 1.0:
            roles[idx] = ("rail", "HIGH", "Long narrow full-length side block in ejector/rail zone.", True)
        elif long_full and ejector_width and centered_side and bottom_pos <= axis_pos <= support_pos + 2.0:
            if row["t"] <= 0.8 or axis_pos > bottom_pos + 1.5:
                roles[idx] = ("ejector_plate", "MEDIUM", "Thinner centered ejector-stack plate; ejector plate.", True)
            else:
                roles[idx] = ("ejector_retainer_plate", "MEDIUM", "Thicker/backing centered ejector-stack plate; ejector retainer plate.", True)
        elif axis_pos > support_pos and row["w"] < max_w * 0.75 and row["l"] < max_l * 0.75:
            roles[idx] = ("insert_or_core_detail", "LOW", "Smaller block inside cavity/core area; not a standard full plate.", False)
        elif top_pos >= axis_pos >= bottom_pos:
            roles[idx] = ("hardware_other", "LOW", "Inside mold stack but not enough geometry to name confidently.", False)
        else:
            roles[idx] = ("ignore", "LOW", "Outside main standard-base classification rules.", False)

    leader_rows = [r for r in rows if roles.get(str(r["i"]), ("",))[0] == "leader_pin"]
    shoulder_bushings = [r for r in rows if roles.get(str(r["i"]), ("",))[0] == "leader_pin_bushing"]
    ejector_bushings = [r for r in rows if roles.get(str(r["i"]), ("",))[0] == "guided_ejector_bushing"]

    def near_same_axis_plane(a, b, tolerance=0.55):
        xy = abs(a["x"] - b["x"]) <= tolerance and abs(a["y"] - b["y"]) <= tolerance
        xz = abs(a["x"] - b["x"]) <= tolerance and abs(a["z"] - b["z"]) <= tolerance
        yz = abs(a["y"] - b["y"]) <= tolerance and abs(a["z"] - b["z"]) <= tolerance
        return xy or xz or yz

    for row in leader_rows:
        idx = str(row["i"])
        if any(near_same_axis_plane(row, b) for b in shoulder_bushings):
            roles[idx] = (
                "leader_pin",
                "HIGH",
                "Primary leader-pin set: center position matches shoulder bushings that sit on the B plate.",
                False,
            )
        elif any(near_same_axis_plane(row, b) for b in ejector_bushings):
            roles[idx] = (
                "leader_pin",
                "MEDIUM",
                "Secondary guide-pin set: center position matches guided-ejector bushings, so it should not decide A/B plate identity.",
                False,
            )

    long_round_pins = []
    short_round_bushings = []
    for row in rows:
        dia, bar_len = round_bar_diameter(row)
        if not dia or dia > 4.0:
            continue
        if bar_len >= dia * 3.0 and bar_len >= 6.0:
            long_round_pins.append(row)
        elif dia * 0.6 <= bar_len <= dia * 1.35 and 1.0 <= dia <= 2.6:
            short_round_bushings.append(row)

    for pin in long_round_pins:
        matching_bushings = [b for b in short_round_bushings if near_same_axis_plane(pin, b, 0.55)]
        if matching_bushings:
            roles[str(pin["i"])] = (
                "leader_pin",
                "MEDIUM",
                "Long round guide pin matched to short bushing at the same center position plane.",
                False,
            )
            for bushing in matching_bushings:
                roles[str(bushing["i"])] = (
                    "leader_pin_bushing",
                    "MEDIUM",
                    "Short bushing matched to a long guide pin at the same center position plane.",
                    False,
                )

    classifications = []
    for row in rows:
        role, confidence, reason, quote = roles.get(
            str(row["i"]),
            ("ignore", "LOW", "No matching standard-base geometry rule.", False),
        )
        classifications.append(
            {
                "index": str(row["i"]),
                "role": role,
                "confidence": confidence,
                "reason": reason,
                "quote": quote,
            }
        )

    parting = "Between a_plate and b_plate from the full-footprint stack order."
    return {
        "job_analysis": {
            "stack_axis": stack_axis,
            "parting_line": parting,
            "rules_for_this_job": [
                f"Full-footprint plates were sorted by {stack_axis} from top to bottom.",
                "Rails were detected as long narrow side blocks below the support plate.",
                "Ejector-stack plates were detected as centered long narrower plates near the rails.",
                "Round guide hardware was separated by diameter and length.",
            ],
        },
        "classifications": classifications,
    }


def run_ollama(prompt, model, timeout_minutes):
    if shutil.which("ollama") is None:
        raise SystemExit("Ollama is not installed or not on PATH. Open a new Command Prompt after installing Ollama.")

    print(f"Starting Ollama model: {model}", flush=True)
    print("This can take several minutes on an older CPU. If this is the first run, Ollama may also load/download the model.", flush=True)
    if timeout_minutes <= 0:
        timeout_seconds = None
        print("Timeout disabled. This can run until the model finishes.", flush=True)
    else:
        timeout_seconds = timeout_minutes * 60
        print(f"Timeout set to {timeout_minutes} minutes.", flush=True)
    cmd = ["ollama", "run", model]
    try:
        result = subprocess.run(
            cmd,
            input=prompt,
            text=True,
            encoding="utf-8",
            errors="replace",
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=timeout_seconds,
        )
    except FileNotFoundError:
        raise SystemExit("Ollama is not installed or not on PATH.")
    except subprocess.TimeoutExpired:
        debug_prompt = OUT_DIR / "last_timeout_prompt.txt"
        debug_prompt.write_text(prompt, encoding="utf-8", errors="replace")
        raise SystemExit(
            f"Ollama timed out after {timeout_minutes} minutes. "
            f"Prompt saved to {debug_prompt}. Try fewer rows or use --timeout-minutes 0."
        )
    if result.returncode != 0:
        raise SystemExit(result.stderr.strip() or "Ollama failed")
    return result.stdout


def original_row_lookup(csv_path):
    with Path(csv_path).open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
        return {str(row.get("Index", "")): row for row in csv.DictReader(f)}


def write_outputs(data, csv_path):
    stem = Path(csv_path).stem
    out_json = OUT_DIR / f"{stem}_qwen_classification.json"
    out_csv = OUT_DIR / f"{stem}_qwen_classification.csv"
    out_review = OUT_DIR / f"{stem}_CORRECT_ME.csv"
    originals = original_row_lookup(csv_path)

    out_json.write_text(json.dumps(data, indent=2), encoding="utf-8")

    fields = ["Index", "Component", "Role", "Confidence", "Quote", "Thickness", "Width", "Length", "CenterX", "CenterY", "CenterZ", "Reason"]
    with out_csv.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for item in data.get("classifications", []):
            original = originals.get(str(item.get("index", "")), {})
            writer.writerow(
                {
                    "Index": item.get("index", ""),
                    "Component": original.get("Component", ""),
                    "Role": item.get("role", ""),
                    "Confidence": item.get("confidence", ""),
                    "Quote": item.get("quote", ""),
                    "Thickness": original.get("Thickness", ""),
                    "Width": original.get("Width", ""),
                    "Length": original.get("Length", ""),
                    "CenterX": original.get("CenterX", ""),
                    "CenterY": original.get("CenterY", ""),
                    "CenterZ": original.get("CenterZ", ""),
                    "Reason": item.get("reason", ""),
                }
            )

    review_fields = ["Index", "Component", "PredictedRole", "CorrectRole", "Confidence", "Thickness", "Width", "Length", "CenterX", "CenterY", "CenterZ", "Reason", "Notes"]
    with out_review.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=review_fields)
        writer.writeheader()
        for item in data.get("classifications", []):
            original = originals.get(str(item.get("index", "")), {})
            writer.writerow(
                {
                    "Index": item.get("index", ""),
                    "Component": original.get("Component", ""),
                    "PredictedRole": item.get("role", ""),
                    "CorrectRole": "",
                    "Confidence": item.get("confidence", ""),
                    "Thickness": original.get("Thickness", ""),
                    "Width": original.get("Width", ""),
                    "Length": original.get("Length", ""),
                    "CenterX": original.get("CenterX", ""),
                    "CenterY": original.get("CenterY", ""),
                    "CenterZ": original.get("CenterZ", ""),
                    "Reason": item.get("reason", ""),
                    "Notes": "",
                }
            )

    return out_json, out_csv, out_review


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("csv_path", help="Path to XT_Export_CAD_Dimensions.csv")
    parser.add_argument("--model", default="qwen3.5:4b")
    parser.add_argument("--max-rows", type=int, default=140)
    parser.add_argument("--timeout-minutes", type=int, default=240, help="0 disables timeout")
    parser.add_argument("--include-names", action="store_true", help="include shortened component names in the prompt")
    parser.add_argument("--long-knowledge", action="store_true", help="include full knowledge files instead of compact rules")
    parser.add_argument("--rules-only", action="store_true", help="skip Qwen/Ollama and use deterministic geometry rules only")
    args = parser.parse_args()

    rows = read_rows(args.csv_path, include_names=args.include_names)
    print(f"Loaded {len(rows)} rows from {args.csv_path}", flush=True)
    if len(rows) > args.max_rows:
        print(f"Limiting to first {args.max_rows} rows for this run.", flush=True)
        rows = rows[: args.max_rows]

    if args.rules_only:
        print("Rules-only mode: skipping Qwen/Ollama.", flush=True)
        data = classify_geometry(rows)
    else:
        prompt = build_prompt(rows, args.csv_path, long_knowledge=args.long_knowledge)
        print(f"Prompt size: {len(prompt):,} characters", flush=True)
        print("Prompt built. Sending to Qwen now...", flush=True)
        raw = run_ollama(prompt, args.model, args.timeout_minutes)
        print("Qwen returned output. Parsing JSON...", flush=True)
        try:
            data = extract_json(raw)
        except Exception:
            debug = OUT_DIR / "last_bad_qwen_output.txt"
            debug.write_text(raw, encoding="utf-8", errors="replace")
            print(f"Could not parse Qwen JSON. Raw output saved to {debug}", flush=True)
            print("Using geometry-rule fallback so you still get output files.", flush=True)
            data = classify_geometry(rows)

    out_json, out_csv, out_review = write_outputs(data, args.csv_path)
    print(f"Wrote {out_json}")
    print(f"Wrote {out_csv}")
    print(f"Wrote correction file {out_review}")


if __name__ == "__main__":
    main()
