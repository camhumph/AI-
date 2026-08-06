from pathlib import Path
import datetime as dt
import shutil
import re


ROOT = Path(r"C:\CMS_AI\webapp")
TARGET = ROOT / "backend" / "app" / "sheet_pricing.py"


HELPER_CODE = r'''
# ============================================================
# CMS PATCH: BMS six-block display from Module6121 BOM match CSV
# ------------------------------------------------------------
# Source of truth:
#   XT_Export_BOM_Match_Report.csv
#
# Reads numbers directly from CSV columns:
#   CAD_Thickness, CAD_Width, CAD_Length
# with fallback to:
#   BOM_Thickness, BOM_Width, BOM_Length
# ============================================================

_CMS_BMS_ROLE_ORDER = [
    "TCP",
    "BCP",
    "ID HOLDER",
    "OD HOLDER",
    "ID POT BLOCK",
    "OD POT BLOCK",
]

_CMS_BMS_ROLE_ALIASES = {
    "TCP": "TCP",
    "TOP CLAMP": "TCP",
    "TOP CLAMP PLATE": "TCP",
    "TOP CLAMPING": "TCP",
    "TOP CLAMPING PLATE": "TCP",
    "TOP SMED": "TCP",
    "TOP SMED PLATE": "TCP",
    "ID SMED": "TCP",
    "ID SMED PLATE": "TCP",

    "BCP": "BCP",
    "BOTTOM CLAMP": "BCP",
    "BOT CLAMP": "BCP",
    "BOTTOM CLAMP PLATE": "BCP",
    "BOT CLAMP PLATE": "BCP",
    "BOTTOM CLAMPING": "BCP",
    "BOTTOM CLAMPING PLATE": "BCP",
    "BOT CLAMPING": "BCP",
    "BOT CLAMPING PLATE": "BCP",
    "BOTTOM SMED": "BCP",
    "BOT SMED": "BCP",
    "OD SMED": "BCP",
    "OD SMED PLATE": "BCP",

    "ID HOLDER": "ID HOLDER",
    "ID HOLDER BLOCK": "ID HOLDER",
    "TOP HOLDER": "ID HOLDER",
    "TOP HOLDER BLOCK": "ID HOLDER",
    "ID MOLD BASE": "ID HOLDER",
    "ID MOLDBASE": "ID HOLDER",
    "TOP MOLD BASE": "ID HOLDER",
    "TOP MOLDBASE": "ID HOLDER",

    "OD HOLDER": "OD HOLDER",
    "OD HOLDER BLOCK": "OD HOLDER",
    "BOTTOM HOLDER": "OD HOLDER",
    "BOT HOLDER": "OD HOLDER",
    "BOTTOM HOLDER BLOCK": "OD HOLDER",
    "BOT HOLDER BLOCK": "OD HOLDER",
    "OD MOLD BASE": "OD HOLDER",
    "OD MOLDBASE": "OD HOLDER",
    "BOTTOM MOLD BASE": "OD HOLDER",
    "BOT MOLD BASE": "OD HOLDER",

    "ID POT": "ID POT BLOCK",
    "ID POT BLOCK": "ID POT BLOCK",
    "TOP POT": "ID POT BLOCK",
    "TOP POT BLOCK": "ID POT BLOCK",
    "TCP POT": "ID POT BLOCK",
    "TCP POT BLOCK": "ID POT BLOCK",

    "OD POT": "OD POT BLOCK",
    "OD POT BLOCK": "OD POT BLOCK",
    "BOTTOM POT": "OD POT BLOCK",
    "BOT POT": "OD POT BLOCK",
    "BOTTOM POT BLOCK": "OD POT BLOCK",
    "BOT POT BLOCK": "OD POT BLOCK",
    "BCP POT": "OD POT BLOCK",
    "BCP POT BLOCK": "OD POT BLOCK",
}


def _cms_bms_norm(value):
    return " ".join(
        str(value or "")
        .replace("_", " ")
        .replace("-", " ")
        .replace(".", " ")
        .upper()
        .split()
    )


def _cms_bms_float(value, default=0.0):
    try:
        s = str(value or "").strip().replace("$", "").replace(",", "")
        if not s:
            return default
        return float(s)
    except Exception:
        return default


def _cms_bms_qty(value, default=1):
    try:
        s = str(value or "").strip()
        if not s:
            return default
        n = int(float(s))
        return n if n > 0 else default
    except Exception:
        return default


def _cms_bms_role(value):
    n = _cms_bms_norm(value)

    if n in _CMS_BMS_ROLE_ALIASES:
        return _CMS_BMS_ROLE_ALIASES[n]

    for key, role in _CMS_BMS_ROLE_ALIASES.items():
        if key in n:
            return role

    return None


def _cms_read_bms_six_from_bom_match(job_dir):
    """
    Read BMS six quoted base components from Module6121's
    XT_Export_BOM_Match_Report.csv.

    The web display dimensions come directly from:
      CAD_Thickness, CAD_Width, CAD_Length
    with fallback to:
      BOM_Thickness, BOM_Width, BOM_Length.
    """
    from pathlib import Path
    import csv

    job_dir = Path(job_dir)

    candidates = [
        job_dir / "XT_Export_BOM_Match_Report.csv",
        job_dir / "documents" / "XT_Export_BOM_Match_Report.csv",
    ]

    csv_path = None
    for p in candidates:
        if p.exists():
            csv_path = p
            break

    if csv_path is None:
        return []

    rows_by_role = {}

    with csv_path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        reader = csv.DictReader(f)

        for row in reader:
            role = _cms_bms_role(row.get("QuoteName", ""))
            if not role:
                continue

            thickness = _cms_bms_float(row.get("CAD_Thickness"))
            width = _cms_bms_float(row.get("CAD_Width"))
            length = _cms_bms_float(row.get("CAD_Length"))

            if thickness <= 0:
                thickness = _cms_bms_float(row.get("BOM_Thickness"))
            if width <= 0:
                width = _cms_bms_float(row.get("BOM_Width"))
            if length <= 0:
                length = _cms_bms_float(row.get("BOM_Length"))

            if thickness <= 0 or width <= 0 or length <= 0:
                continue

            qty = _cms_bms_qty(row.get("Qty"), 1)

            rows_by_role[role] = {
                "description": role,
                "name": role,
                "role": role,
                "resolved_name": role,

                "role_group_name": "Steel Plates / Mold Base",
                "group": "Steel Plates / Mold Base",
                "category": "Steel Plates / Mold Base",

                "qty": qty,
                "quantity": qty,

                "thickness": thickness,
                "width": width,
                "length": length,

                "Thickness": thickness,
                "Width": width,
                "Length": length,

                "hours": None,
                "price": 0.0,
                "total": 0.0,
                "quoted": True,
                "quote": True,
                "pricing_note": "sheet",

                "source": "XT_Export_BOM_Match_Report.csv",
                "price_source": "job_csv:XT_Export_BOM_Match_Report.csv",
                "csv_path": str(csv_path),
            }

    ordered = []
    for role in _CMS_BMS_ROLE_ORDER:
        if role in rows_by_role:
            ordered.append(rows_by_role[role])

    return ordered
'''


def read_file(path):
    data = path.read_bytes()

    if data.startswith(b"\xff\xfe") or data.startswith(b"\xfe\xff"):
        enc = "utf-16"
    elif data.startswith(b"\xef\xbb\xbf"):
        enc = "utf-8-sig"
    else:
        sample = data[:4000]
        null_ratio = sample.count(b"\x00") / max(len(sample), 1)
        if null_ratio > 0.20:
            enc = "utf-16-le"
        else:
            enc = "cp1252"
            for candidate in ("utf-8-sig", "utf-8", "cp1252"):
                try:
                    data.decode(candidate)
                    enc = candidate
                    break
                except UnicodeDecodeError:
                    pass

    text = data.decode(enc, errors="replace")

    crlf = text.count("\r\n")
    lf = text.count("\n") - crlf
    newline = "\r\n" if crlf >= lf else "\n"

    text = text.replace("\r\n", "\n").replace("\r", "\n")
    return text, enc, newline


def write_file(path, text, enc, newline):
    path.write_text(text.replace("\n", newline), encoding=enc, newline="")


def get_indent(line):
    return line[:len(line) - len(line.lstrip(" \t"))]


def find_previous_def(lines, start_index):
    for i in range(start_index, -1, -1):
        stripped = lines[i].lstrip()
        if stripped.startswith("def "):
            return i
    return None


def collect_def_text(lines, def_index):
    collected = []
    for i in range(def_index, min(def_index + 20, len(lines))):
        collected.append(lines[i].strip())
        if lines[i].rstrip().endswith(":"):
            break
    return " ".join(collected)


def get_function_name(def_text):
    s = def_text.strip()
    if not s.startswith("def "):
        return ""

    s = s[4:]
    pos = s.find("(")
    if pos < 0:
        return ""

    return s[:pos].strip()


def get_first_param(def_text):
    left = def_text.find("(")
    right = def_text.find(")", left + 1)

    if left < 0 or right < 0:
        return "job_dir"

    inside = def_text[left + 1:right]
    parts = [p.strip() for p in inside.split(",") if p.strip()]

    for p in parts:
        name = p.split(":", 1)[0].split("=", 1)[0].strip()
        if name not in ("self", "cls"):
            return name

    return "job_dir"


def find_steel_docstring_line(lines):
    needle = "Steel plates from quote workbook #2 block, else J000 steel sheet."
    for i, line in enumerate(lines):
        if needle in line:
            return i
    return None


def insert_helper_after_imports(text):
    if "_cms_read_bms_six_from_bom_match" in text:
        return text, "BMS CSV helper already present"

    lines = text.splitlines(True)

    last_import = -1
    for i, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith("import ") or stripped.startswith("from "):
            last_import = i

    helper = HELPER_CODE.strip() + "\n\n"

    if last_import >= 0:
        lines.insert(last_import + 1, "\n" + helper)
        return "".join(lines), "inserted BMS CSV helper after imports"

    return helper + text, "inserted BMS CSV helper at top"


def patch_steel_function(text):
    lines = text.splitlines(True)

    doc_i = find_steel_docstring_line(lines)
    if doc_i is None:
        raise RuntimeError("Could not find steel plate function docstring in sheet_pricing.py")

    # If already patched near docstring, skip.
    lookahead = "".join(lines[doc_i:doc_i + 20])
    if "_cms_bms_rows = _cms_read_bms_six_from_bom_match" in lookahead:
        return text, "steel function already patched"

    def_i = find_previous_def(lines, doc_i)
    if def_i is None:
        raise RuntimeError("Could not find def line before steel plate docstring")

    def_text = collect_def_text(lines, def_i)
    func_name = get_function_name(def_text)
    first_param = get_first_param(def_text)

    indent = get_indent(lines[doc_i])

    injection = (
        "\n"
        + indent + "# CMS PATCH: BMS/pot-block jobs display all six quoted blocks\n"
        + indent + "# from XT_Export_BOM_Match_Report.csv. Numbers are read\n"
        + indent + "# from CAD_Thickness, CAD_Width, CAD_Length first.\n"
        + indent + "_cms_bms_rows = _cms_read_bms_six_from_bom_match(" + first_param + ")\n"
        + indent + "if len(_cms_bms_rows) >= 4:\n"
        + indent + "    return _cms_bms_rows\n"
    )

    # Insert after docstring line.
    lines.insert(doc_i + 1, injection)

    return "".join(lines), "patched " + func_name + " to return BMS CSV rows first"


def patch(text):
    log = []

    text, msg = insert_helper_after_imports(text)
    log.append(msg)

    text, msg = patch_steel_function(text)
    log.append(msg)

    return text, log


def main():
    if not TARGET.exists():
        print("ERROR: file not found:", TARGET)
        return 1

    text, enc, newline = read_file(TARGET)

    try:
        patched, log = patch(text)
    except Exception as e:
        print("ERROR:", e)
        return 1

    if patched == text:
        print("No changes needed.")
        for item in log:
            print(" -", item)
        return 0

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup = TARGET.with_name(TARGET.name + ".bak_" + stamp)

    shutil.copy2(TARGET, backup)
    write_file(TARGET, patched, enc, newline)

    print("Patched:", TARGET)
    print("Backup: ", backup)
    print("Encoding preserved:", enc)
    print("")
    print("Patch log:")
    for item in log:
        print(" -", item)

    print("")
    print("Restart the web app:")
    print(r"cd /d C:\CMS_AI\webapp")
    print("START_CMS_QUOTING_APP.bat")
    print("")
    print("Then hard-refresh browser with Ctrl+F5.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())