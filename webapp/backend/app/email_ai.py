"""Local-LLM reading of an incoming RFQ email plus its attachments.

WHY THIS EXISTS
---------------
extract_quote_info() in email_service.py is a regex scraper. It pulls four
things -- cust_job, c_number, ship_date, similar_to -- and is blind to
everything the customer actually SAID. On the Hewitt RFQ for 25-424 the body
reads:

    "Please quote this custom mold base. File is .stp. Include all components
     that are present in the design. Please see the attached bill of materials
     for your reference on plate material, quantities, component part #'s, etc.
     I'm looking for the best possible delivery."

Every clause there is a quoting instruction -- custom base (not a catalog
base), price the hardware too (not just steel), the BOM is authoritative for
material/qty/part numbers, and delivery is being competed on. The old scraper
returned ship_date="" and similar_to="" for that email and threw the rest away,
so a human had to read it anyway before the macro could be trusted.

WHAT THIS MODULE DOES *NOT* DO
------------------------------
It does not re-implement BOM parsing for pricing. Module6121.bas already reads
these BOMs properly -- FindBomHeaderLikeInArrayRow skips the title block,
ParseInchDimsFromText handles "1.375 x 9-7/8 x 11-7/8", and NormalizeSteelType
maps Hewitt's "#7 steel" to 420SS (Module6121.bas:15650). That code stays the
pricing authority. This module runs BEFORE SolidWorks opens, to answer a
different question: "what is this email asking for, and can we quote it?"

THE SPLIT THAT MAKES THIS SAFE
------------------------------
Python computes the FACTS. The LLM only INTERPRETS.

    facts  (deterministic, this file)   quantities, dimensions, steel grades,
                                       vendor part numbers, BOM<->CAD
                                       reconciliation, file inventory
    brief  (qwen via Ollama)           request type, scope, delivery urgency,
                                       special instructions, open questions

A local 9B model asked to count plates will sometimes say six when there are
seven, and a wrong quantity is money. So no number in `facts` is ever sourced
from the model, and merge_brief() drops any count the model contradicts,
recording the disagreement instead of silently trusting either side. If Ollama
is not running the facts still come back in full, with llm.ok False -- the
quote is never blocked on the model being up.

WHAT IT CAUGHT ON THE FIRST REAL JOB
------------------------------------
25-424's BOM lists 8 fastener lines; the .stp contains 7. The missing one is
`bhcs--8-32 x .37` ("Hewitt ID tag"), and it is the one BOM row whose include
column is blank while all 22 others carry an X. The customer wrote "include all
components that are present in the design", so that line should be excluded --
a reconciliation the regex scraper could not have surfaced because it never
opened either attachment.

USAGE
-----
    from . import email_ai
    brief = email_ai.analyze(subject, body, attachment_paths, from_addr=...)

    # standalone, against files on disk (no mail server needed):
    python -m webapp.backend.app.email_ai --body-file rfq.txt \\
        "C:/Users/lenovo/Downloads/25-424 BOM base.xls" \\
        "C:/Users/lenovo/Downloads/25-424--ba-quote.stp"
"""
from __future__ import annotations

import argparse
import json
import os
import re
import urllib.error
import urllib.request
import zipfile
from pathlib import Path

# --- Ollama ----------------------------------------------------------------

OLLAMA_URL = os.environ.get("CMS_OLLAMA_URL", "http://localhost:11434")

# SMALLEST MODEL FIRST, ON PURPOSE.
#
# This box runs Ollama on CPU at about 3.8 tokens/sec for qwen3.5:9b. Measured on
# the 25-424 RFQ: the 9b never finished a full brief inside 560 s, and qwen3:8b
# took 75 s on even a trimmed prompt. qwen3:1.7b answered the same prompt in 25 s
# (including model load) and got the company, job number, scope sentence and both
# special instructions right -- because after gather_facts() there is very little
# left to think about. The model reads a few sentences of English; the plate
# count, dimensions, grades and BOM<->CAD reconciliation were already computed.
#
# So size buys nothing here and costs minutes. Pin a bigger tag with
# CMS_EMAIL_AI_MODEL, or pass --deep for the full-digest prompt, if you ever want
# the heavier read.
MODEL_PREFERENCE = ("qwen3:1.7b", "qwen3.5:4b", "qwen3:8b", "qwen3.5:9b")

# Ollama defaults num_ctx to 4096 whatever the model supports, which would
# silently truncate a long digest mid-row and make the model answer about half a
# job. Sized per prompt in ask_model() instead of pinned high, because an
# oversized context costs RAM and load time for no gain on the compact prompt.
NUM_CTX_MIN = 4096
NUM_CTX_MAX = int(os.environ.get("CMS_EMAIL_AI_NUM_CTX", "24576"))

# Hard ceiling on generated tokens. The brief is a fixed shape; anything past
# this is the model rambling, and at 4 tok/s rambling is measured in minutes.
NUM_PREDICT = int(os.environ.get("CMS_EMAIL_AI_NUM_PREDICT", "600"))

# Snappy by default. The facts layer is already complete when this starts, so a
# timeout costs interpretation, never data -- merge_brief() falls back cleanly.
LLM_TIMEOUT_S = int(os.environ.get("CMS_EMAIL_AI_TIMEOUT", "120"))

# How long Ollama holds the weights in RAM after a call.
KEEP_ALIVE = os.environ.get("CMS_EMAIL_AI_KEEP_ALIVE", "30m")

BRIEF_FILENAME = "email_ai_brief.json"

# Cap on how much of a CAD file we will walk. STEP assemblies are usually a few
# MB, but a full tool with cavity detail can be hundreds; past this we stop and
# say so rather than stalling the inbox poller.
STEP_SCAN_BYTE_LIMIT = int(os.environ.get("CMS_EMAIL_AI_STEP_LIMIT", str(240 * 1024 * 1024)))

CAD_EXTS = {".stp", ".step", ".x_t", ".x_b", ".xmt_txt", ".sldasm", ".sldprt", ".igs", ".iges", ".3dm", ".prt"}
BOM_EXTS = {".xls", ".xlsx", ".xlsm", ".xlt", ".xltx", ".csv", ".tsv"}
DOC_EXTS = {".pdf", ".doc", ".docx", ".txt", ".rtf"}


# ===========================================================================
# Number / dimension parsing
# ===========================================================================

# "9-7/8"  "1-1/4"  "7/8"  "1.375"  ".50"
_FRACTION_RE = re.compile(r"(?<![\d./])(\d+)\s*-\s*(\d+)\s*/\s*(\d+)(?![\d/])")
_BARE_FRACTION_RE = re.compile(r"(?<![\d./])(\d+)\s*/\s*(\d+)(?![\d/])")


def frac_to_float(text: str) -> float | None:
    """Parse one shop dimension token. "9-7/8" -> 9.875, ".50" -> 0.5."""
    s = (text or "").strip().strip('"').replace("\u2033", "").strip()
    if not s:
        return None
    m = _FRACTION_RE.fullmatch(s)
    if m:
        whole, num, den = (int(g) for g in m.groups())
        return whole + num / den if den else float(whole)
    m = _BARE_FRACTION_RE.fullmatch(s)
    if m:
        num, den = int(m.group(1)), int(m.group(2))
        return num / den if den else None
    try:
        return float(s)
    except ValueError:
        return None


def _dim_tokens(text: str) -> list[float]:
    """Every number in an 'A x B x C' run, fractions resolved, in file order."""
    if not text:
        return []
    # Normalise the separators the shop uses interchangeably.
    s = text.replace("\u00d7", "x").replace("X", "x")
    # Split on 'x' only when it sits between numbers, so "6-7/8 x 11-7/8" splits
    # but "Ejector" (which contains no standalone x) is untouched.
    parts = re.split(r"(?<=[\d\"/])\s*x\s*(?=[\d.])", s)
    out: list[float] = []
    for part in parts:
        m = re.search(r"\d+\s*-\s*\d+\s*/\s*\d+|\d+\s*/\s*\d+|\d*\.\d+|\d+", part)
        if not m:
            continue
        val = frac_to_float(m.group(0))
        if val is not None:
            out.append(val)
    return out


def parse_inch_dims(text: str) -> dict | None:
    """Pull a plate or round size out of a BOM description.

    "Top clamp plate, 1.375 x 9-7/8 x 11-7/8" -> thickness/width/length
    "Leader pin, 1.000 x 3-1/4"               -> diameter/length

    Returns None when the description carries no size, which is normal for
    section headers and set-sold rows ("Sold as a set with 8820").
    """
    # Only consider the part of the string that actually contains an 'A x B' run,
    # so a stray number in the name ("#10-28") cannot become a dimension.
    m = re.search(r"[\d./\-]+\s*(?:x|\u00d7)\s*[\d./\-\s\"x\u00d7]+", text or "", re.I)
    if not m:
        return None
    nums = _dim_tokens(m.group(0))
    nums = [n for n in nums if n > 0]
    if len(nums) >= 3:
        t, w, l = nums[0], nums[1], nums[2]
        return {"kind": "plate", "thickness": t, "width": w, "length": l,
                "sorted": sorted([t, w, l], reverse=True)}
    if len(nums) == 2:
        return {"kind": "round", "diameter": nums[0], "length": nums[1]}
    return None


# ===========================================================================
# Steel grades
# ===========================================================================

# Deliberately a line-for-line mirror of NormalizeSteelType in Module6121.bas
# (~15635). Order matters: an explicit grade beats a "#n" code, because a BOM
# that says "4140 (#2)" must not resolve twice. Keep the two in sync -- if the
# macro learns a grade and this does not, the brief will disagree with the
# quote the macro then builds.
_STEEL_RULES = (
    ("PYROPEL", "Pyropel"),
    ("4140", "4140"),
    ("P20", "P20"), ("P-20", "P20"),
    ("A36", "A36"), ("A-36", "A36"), ("1045", "A36"), ("1030", "A36"),
    ("1020", "A36"), ("HOT ROLLED", "A36"), ("COLD ROLLED", "A36"),
    ("420", "420SS"), ("S136", "420SS"),
    ("H13", "H13"), ("H-13", "H13"),
    ("6061", "6061"), ("ALUM", "6061"),
    ("A-2", "A2"),
    ("O-1", "O1"), ("0-1", "O1"), ("FLAT GROUND", "O1"),
)

# DME grade codes ("#7 steel"), as they appear on customer BOMs and shop sheets.
#
# These need a guard the macro's plain InStr does not have. This module scans
# EVERY cell of a BOM row for a material token, so a bare InStr(u,"#1") also
# fires on the thread size in "Flat head screw, #10-28 x 1/2" and would tag that
# row A-36. Requiring the digit not to be followed by another digit -- or by
# "-<digit>", which is how thread sizes are written -- keeps "#7 steel" and
# rejects "#10-28".
_DME_CODE_RE = tuple(
    (re.compile(rf"#{n}(?!\s*-?\d)"), grade)
    for n, grade in (("7", "420SS"), ("5", "H13"), ("3", "P20"),
                     ("2", "4140"), ("1", "A36"))
)

# Grades the quote workbook has rows for (plate_grades.VALID_GRADES plus the two
# the macro normalises but the override picker does not offer).
PRICEABLE_GRADES = {"A36", "4140", "P20", "420SS", "6061", "A2", "O1", "H13", "Pyropel"}

GRADE_LABELS = {
    "A36": "#1 A-36", "4140": "#2 4140", "P20": "#3 P20", "H13": "#5 H-13",
    "420SS": "#7 420-SS", "6061": "6061 ALM", "A2": "A-2", "O1": "O-1",
    "Pyropel": "Pyropel",
}


def normalize_steel(mat: str) -> str:
    """Customer material text -> CMS grade key. "" when nothing matches."""
    u = (mat or "").strip().upper()
    if not u:
        return ""
    for token, grade in _STEEL_RULES:
        if token in u:
            return grade
    if u == "A2":
        return "A2"
    if u == "O1":
        return "O1"
    for pattern, grade in _DME_CODE_RE:
        if pattern.search(u):
            return grade
    return ""


# ===========================================================================
# Attachment readers
# ===========================================================================

def _sheet_rows_xls(path: Path) -> dict:
    import xlrd  # legacy .xls only; xlrd 2.x dropped .xlsx on purpose

    wb = xlrd.open_workbook(str(path))
    sheets = {}
    for name in wb.sheet_names():
        sh = wb.sheet_by_name(name)
        rows = []
        for r in range(sh.nrows):
            rows.append([_cell_text(sh.cell_value(r, c)) for c in range(sh.ncols)])
        sheets[name] = rows
    return sheets


def _sheet_rows_xlsx(path: Path) -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(str(path), data_only=True, read_only=True)
    sheets = {}
    for name in wb.sheetnames:
        rows = []
        for row in wb[name].iter_rows(values_only=True):
            rows.append([_cell_text(v) for v in row])
        sheets[name] = rows
    wb.close()
    return sheets


def _sheet_rows_csv(path: Path) -> dict:
    import csv

    delim = "\t" if path.suffix.lower() == ".tsv" else ","
    with path.open("r", encoding="utf-8-sig", errors="replace", newline="") as f:
        rows = [[_cell_text(v) for v in row] for row in csv.reader(f, delimiter=delim)]
    return {path.stem: rows}


def _cell_text(value) -> str:
    """Excel cell -> text, without the trailing .0 on whole numbers."""
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    return str(value).strip()


def read_spreadsheet(path: Path) -> dict:
    ext = path.suffix.lower()
    try:
        if ext == ".xls":
            return _sheet_rows_xls(path)
        if ext in (".xlsx", ".xlsm", ".xltx", ".xlt"):
            return _sheet_rows_xlsx(path)
        if ext in (".csv", ".tsv"):
            return _sheet_rows_csv(path)
    except Exception as e:
        return {"__error__": [[f"{type(e).__name__}: {e}"]]}
    return {}


def read_pdf_text(path: Path, max_chars: int = 20000) -> str:
    try:
        from pypdf import PdfReader
    except ImportError:
        return ""
    try:
        reader = PdfReader(str(path))
        chunks = []
        total = 0
        for page in reader.pages:
            txt = page.extract_text() or ""
            chunks.append(txt)
            total += len(txt)
            if total >= max_chars:
                break
        return "\n".join(chunks)[:max_chars]
    except Exception:
        return ""


# --- STEP ------------------------------------------------------------------

# Only these entity types are kept while streaming; everything else (the
# millions of CARTESIAN_POINTs) is discarded as it goes, so memory stays flat
# regardless of file size.
_STEP_KEEP = (
    "PRODUCT",
    "PRODUCT_DEFINITION",
    "PRODUCT_DEFINITION_FORMATION",
    "PRODUCT_DEFINITION_FORMATION_WITH_SPECIFIED_SOURCE",
    "NEXT_ASSEMBLY_USAGE_OCCURRENCE",
    "MANIFOLD_SOLID_BREP",
)
_STEP_ENTITY_RE = re.compile(r"^#(\d+)\s*=\s*([A-Z_0-9]+)\s*\((.*)\)\s*$", re.S)
_STEP_REF_RE = re.compile(r"#(\d+)")
_STEP_STR_RE = re.compile(r"'((?:[^']|'')*)'")


def _step_statements(path: Path):
    """Yield one STEP statement at a time. Statements can wrap lines."""
    buf: list[str] = []
    read = 0
    truncated = False
    with path.open("r", encoding="utf-8", errors="replace") as f:
        for line in f:
            read += len(line)
            if read > STEP_SCAN_BYTE_LIMIT:
                truncated = True
                break
            s = line.strip()
            if not s:
                continue
            buf.append(s)
            if s.endswith(";"):
                yield " ".join(buf).rstrip(";"), False
                buf = []
    if buf:
        yield " ".join(buf).rstrip(";"), False
    if truncated:
        yield "", True


def read_step(path: Path) -> dict:
    """Product names and per-product instance counts from a STEP assembly.

    The instance count is what makes this worth doing: resolving each
    NEXT_ASSEMBLY_USAGE_OCCURRENCE back to its PRODUCT gives how many of each
    part the CAD actually contains, which is directly comparable to the BOM's
    QTY column. That comparison is the whole point of the customer's "include
    all components that are present in the design".
    """
    info: dict = {
        "file": path.name,
        "format": "STEP",
        "originating_system": "",
        "timestamp": "",
        "schema": "",
        "products": [],
        "instance_counts": {},
        "assemblies": [],
        "solid_count": 0,
        "truncated": False,
        "error": "",
    }
    products: dict[str, str] = {}          # #ref -> product name
    formation_to_product: dict[str, str] = {}
    definition_to_formation: dict[str, str] = {}
    nauo_children: list[str] = []
    nauo_parents: list[str] = []
    header_lines: list[str] = []
    in_header = False

    try:
        for stmt, truncated in _step_statements(path):
            if truncated:
                info["truncated"] = True
                break
            upper_head = stmt[:40].upper()
            if upper_head.startswith("HEADER"):
                in_header = True
                continue
            if upper_head.startswith("ENDSEC"):
                in_header = False
                continue
            if in_header:
                header_lines.append(stmt)
                continue

            m = _STEP_ENTITY_RE.match(stmt)
            if not m:
                continue
            ref, etype, args = m.group(1), m.group(2), m.group(3)
            if etype not in _STEP_KEEP:
                continue

            if etype == "MANIFOLD_SOLID_BREP":
                info["solid_count"] += 1
            elif etype == "PRODUCT":
                names = _STEP_STR_RE.findall(args)
                if names:
                    products[ref] = names[0].replace("''", "'")
            elif etype in ("PRODUCT_DEFINITION_FORMATION",
                           "PRODUCT_DEFINITION_FORMATION_WITH_SPECIFIED_SOURCE"):
                refs = _STEP_REF_RE.findall(args)
                if refs:
                    formation_to_product[ref] = refs[0]
            elif etype == "PRODUCT_DEFINITION":
                refs = _STEP_REF_RE.findall(args)
                if refs:
                    definition_to_formation[ref] = refs[0]
            elif etype == "NEXT_ASSEMBLY_USAGE_OCCURRENCE":
                # (id, name, description, relating_pd, related_pd, ...) -- the
                # SECOND product_definition reference is the child.
                refs = _STEP_REF_RE.findall(args)
                if len(refs) >= 2:
                    nauo_parents.append(refs[0])
                    nauo_children.append(refs[1])
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"

    header_blob = " ".join(header_lines)
    m = re.search(r"originating_system\s*\*?/?\s*'([^']*)'", header_blob, re.I)
    if not m:
        # ST-Developer writes the value after a /* comment */ marker.
        strs = _STEP_STR_RE.findall(header_blob)
        m = None
        for s in strs:
            if any(v in s.upper() for v in ("NX", "SOLIDWORKS", "CATIA", "CREO", "INVENTOR", "PRO/E")):
                info["originating_system"] = s
                break
    else:
        info["originating_system"] = m.group(1)
    m = re.search(r"(\d{4}-\d{2}-\d{2}T[\d:+\-]+)", header_blob)
    if m:
        info["timestamp"] = m.group(1)
    m = re.search(r"(AUTOMOTIVE_DESIGN|CONFIG_CONTROL_DESIGN|AP\d+)", header_blob, re.I)
    if m:
        info["schema"] = m.group(1)

    def _product_name(pd_ref: str) -> str | None:
        formation = definition_to_formation.get(pd_ref)
        product = formation_to_product.get(formation) if formation else None
        return products.get(product) if product else None

    counts: dict[str, int] = {}
    for child_pd in nauo_children:
        name = _product_name(child_pd)
        if name:
            counts[name] = counts.get(name, 0) + 1

    # A product that appears as the PARENT of an occurrence is an assembly node,
    # not a part. Deriving that from the tree beats guessing from the name: this
    # file's sub-assemblies are "25-424--bm-quote" and "25-424--bs-quote"
    # (movable and stationary halves), which no name pattern reliably catches,
    # and each is placed once so an instance-count test does not catch them
    # either. Without this they were reported as CAD parts missing from the BOM.
    assemblies = set()
    for parent_pd in nauo_parents:
        name = _product_name(parent_pd)
        if name:
            assemblies.add(name)

    info["instance_counts"] = counts
    info["assemblies"] = sorted(assemblies)
    info["products"] = sorted(products.values())
    return info


def read_parasolid(path: Path) -> dict:
    """Parasolid x_t: no assembly tree to walk, so just the header + node count."""
    info = {"file": path.name, "format": "Parasolid", "products": [],
            "instance_counts": {}, "solid_count": 0, "schema": "", "error": "",
            "originating_system": "", "timestamp": "", "truncated": False}
    try:
        head = path.open("r", encoding="utf-8", errors="replace").read(4096)
        m = re.search(r"\*\*ABCDEFGHIJKLMNOPQRSTUVWXYZ.*?TRANSMIT FILE created by modeller version ([\d]+)", head, re.S)
        if m:
            info["originating_system"] = f"Parasolid modeller {m.group(1)}"
    except Exception as e:
        info["error"] = f"{type(e).__name__}: {e}"
    return info


def read_cad(path: Path) -> dict:
    ext = path.suffix.lower()
    if ext in (".stp", ".step"):
        return read_step(path)
    if ext in (".x_t", ".xmt_txt"):
        return read_parasolid(path)
    return {"file": path.name, "format": ext.lstrip(".").upper() or "unknown",
            "products": [], "instance_counts": {}, "solid_count": 0,
            "schema": "", "originating_system": "", "timestamp": "",
            "truncated": False, "error": "no reader for this format"}


def read_zip(path: Path) -> dict:
    try:
        with zipfile.ZipFile(path) as z:
            members = [n for n in z.namelist() if not n.endswith("/")]
        return {"file": path.name, "members": members[:400], "member_count": len(members)}
    except Exception as e:
        return {"file": path.name, "members": [], "member_count": 0,
                "error": f"{type(e).__name__}: {e}"}


# ===========================================================================
# BOM interpretation
# ===========================================================================

_HEADER_DESC_TOKENS = ("DESCRIPTION", "PART NAME", "PART NO", "DETAIL NAME",
                       "COMPONENT", "ITEM DESCRIPTION", "NAME")
_HEADER_QTY_TOKENS = ("QTY", "QUANTITY", "NO. REQ", "NO REQ", "REQ'D", "REQD")
_HEADER_MAT_TOKENS = ("MATERIAL", "MAT'L", "MATL", "STEEL", "GRADE")
_HEADER_VENDOR_TOKENS = ("VENDOR", "SUPPLIER", "MANUFACTURER", "MANUF", "SOURCE")
_HEADER_COMMENT_TOKENS = ("COMMENT", "NOTE", "REMARK")
# "DET NO." is the Tempcraft/BMS spelling and does not contain "DETAIL NO", so it
# needs its own token -- without it those BOMs parsed with no detail column at
# all, every row got detail "", and the CAD join had nothing to match on.
_HEADER_DETAIL_TOKENS = ("DETAIL NAME", "DETAIL NO", "DET NO", "DETAIL",
                         "PART NO", "PART #", "ITEM NO")
# Some BOMs put the size in three numeric columns instead of in the description.
# Tempcraft: "Lth. (in.) | Wth./O.D. (in.) | Hgt./I.D. (in.)".
_HEADER_LEN_TOKENS = ("LTH", "LENGTH", "LGTH")
_HEADER_WID_TOKENS = ("WTH", "WIDTH", "O.D.", "OD (")
_HEADER_HGT_TOKENS = ("HGT", "HEIGHT", "THICKNESS", "THK", "I.D.")
# An explicit kind column beats every heuristic in _classify_row.
_HEADER_TYPE_TOKENS = ("TYPE", "ITEM SOURCE", "SOURCE")

_FASTENER_PREFIXES = ("SHCS", "BHCS", "FHCS", "FHS", "SCS", "SHSS", "SSS", "HHCS",
                      "LHCS", "SET SCREW", "CAP SCREW", "SCREW", "DOWEL", "WASHER",
                      "O-RING", "ORING", "PIPE PLUG", "PLUG", "NPT")


def _find_header_row(rows: list[list[str]]) -> dict | None:
    """Locate the real BOM header row and map its columns.

    Same trap Module6121.bas documents at ~13276: a Hewitt title block contains
    a lone "Description"-ish cell, so the FIRST description-looking row is often
    not the header. Require a description column AND at least one of
    qty/material/vendor beside it before accepting a row.
    """
    weak = None
    for r, row in enumerate(rows[:40]):
        cells = [(c, (v or "").strip().upper()) for c, v in enumerate(row)]
        desc_col = qty_col = mat_col = vendor_col = comment_col = detail_col = None
        len_col = wid_col = hgt_col = type_col = None
        for c, u in cells:
            if not u:
                continue
            # Header cells in these workbooks carry embedded newlines
            # ("Mat'l Spec\nor\nMfg. Part"), so match on a flattened copy.
            u = re.sub(r"\s+", " ", u)
            if len_col is None and any(t in u for t in _HEADER_LEN_TOKENS):
                len_col = c
            if wid_col is None and any(t in u for t in _HEADER_WID_TOKENS):
                wid_col = c
            if hgt_col is None and any(t in u for t in _HEADER_HGT_TOKENS):
                hgt_col = c
            if type_col is None and u in _HEADER_TYPE_TOKENS:
                type_col = c
            # Two columns can both look description-ish. Hewitt's sheet has
            # "Detail Name/No." at column 4 and the real "Description" at column
            # 7, and taking the first match put the part number where the
            # description belonged -- which left every plate row with no size and
            # no material, so all nine plates dropped out of the brief and
            # resurfaced as "in CAD but not on BOM". A column named exactly
            # Description always wins, whichever order they appear in.
            if any(t in u for t in _HEADER_DESC_TOKENS):
                if "DESCRIPTION" in u:
                    desc_col = c
                elif desc_col is None:
                    desc_col = c
            if qty_col is None and any(t in u for t in _HEADER_QTY_TOKENS):
                qty_col = c
            if mat_col is None and any(t in u for t in _HEADER_MAT_TOKENS):
                mat_col = c
            if vendor_col is None and any(t in u for t in _HEADER_VENDOR_TOKENS):
                vendor_col = c
            if comment_col is None and any(t in u for t in _HEADER_COMMENT_TOKENS):
                comment_col = c
            if detail_col is None and any(t in u for t in _HEADER_DETAIL_TOKENS):
                detail_col = c
        if desc_col is None:
            continue
        strong = sum(x is not None for x in (qty_col, mat_col, vendor_col)) >= 1
        found = {"row": r, "desc": desc_col, "qty": qty_col, "material": mat_col,
                 "vendor": vendor_col, "comment": comment_col, "detail": detail_col,
                 "len": len_col, "wid": wid_col, "hgt": hgt_col, "type": type_col}
        if strong:
            return found
        if weak is None:
            weak = found
    return weak


def _dims_from_columns(row: list, hdr: dict) -> dict | None:
    """Size from dedicated Lth./Wth./Hgt. columns, when the BOM has them.

    Returns the same shape parse_inch_dims does. The three values are sorted into
    thickness/width/length rather than trusted in column order: the headers say
    "Lth./Wth./Hgt." but a holder block is routinely listed 6.000 x 6.875 x
    13.875, where the "length" column holds the smallest number. Sorting matches
    what Module6121's SortThreeDimensions does with the same files.
    """
    vals = []
    for key in ("hgt", "wid", "len"):
        idx = hdr.get(key)
        if idx is None or idx >= len(row):
            return None
        v = frac_to_float((row[idx] or "").strip())
        if v is None or v <= 0:
            return None
        vals.append(v)
    if len(vals) != 3:
        return None
    ordered = sorted(vals, reverse=True)
    return {"kind": "plate", "thickness": ordered[2], "width": ordered[1],
            "length": ordered[0], "sorted": ordered, "source": "columns"}


def _row_qty(text: str) -> int:
    m = re.search(r"\d+", text or "")
    return int(m.group(0)) if m else 0


def _classify_row(detail: str, desc: str, vendor: str, comment: str,
                  dims: dict | None, grade: str) -> str:
    """plate | purchased | fastener | note."""
    blob = f"{detail} {desc}".upper()
    if vendor.strip():
        return "purchased"
    du = detail.strip().upper()
    if any(du.startswith(p) or blob.startswith(p) for p in _FASTENER_PREFIXES):
        return "fastener"
    if re.search(r"\b(SCREW|DOWEL PIN|WASHER|O-RING)\b", blob):
        return "fastener"
    if grade and dims and dims.get("kind") == "plate":
        return "plate"
    # A three-dimension size with no vendor is steel we have to buy and cut even
    # when the material cell is blank -- flagged later as grade-unknown.
    if dims and dims.get("kind") == "plate":
        return "plate"
    if dims and dims.get("kind") == "round":
        return "purchased"
    return "note"


def interpret_bom(sheets: dict, source: str) -> dict:
    """Turn spreadsheet rows into plate / purchased / fastener line items."""
    result = {
        "source": source,
        "sheet": "",
        "header_row": None,
        "title": "",
        "plates": [],
        "purchased": [],
        "fasteners": [],
        "notes": [],
        "unparsed_rows": 0,
        "error": "",
    }
    if "__error__" in sheets:
        result["error"] = sheets["__error__"][0][0]
        return result

    # PICK THE SHEET THAT HAS A BOM HEADER, NOT THE BIGGEST ONE.
    #
    # This used to take whichever sheet had the most non-empty rows, which is a
    # proxy for "the real one" only when the workbook has nothing but the BOM and
    # some blank template tabs. The Tempcraft/BMS workbooks break it badly: they
    # carry "Oracle Data" / "Material Data" / "Outsource Data" lookup tables
    # alongside the "BOM" sheet, and those are thousands of rows long. So
    # 861000100 Base BOM.xlsm parsed a price lookup table and reported 4421
    # purchased components and zero plates -- the same for C18496 (5698),
    # C18508 (3249) and C18509 (3259).
    #
    # Score by evidence instead: a findable header row is worth far more than
    # size, and a sheet actually named BOM outranks everything.
    def _sheet_score(name: str, rows: list) -> tuple:
        nonempty = sum(1 for r in rows if any((v or "").strip() for v in r))
        if not nonempty:
            return (0, 0, 0)
        named = 2 if re.fullmatch(r"\s*bom\s*", name, re.I) else (
            1 if "bom" in name.lower() else 0)
        hdr = _find_header_row(rows)
        strength = 0
        if hdr:
            strength = 1 + sum(
                1 for k in ("qty", "material", "vendor", "detail", "len", "type")
                if hdr.get(k) is not None
            )
        return (named, strength, min(nonempty, 500))

    best_name, best_rows, best_score = "", [], (-1, -1, -1)
    for name, rows in sheets.items():
        score = _sheet_score(name, rows)
        if score > best_score:
            best_name, best_rows, best_score = name, rows, score
    result["sheet"] = best_name
    if not best_rows:
        result["error"] = "sheet is empty"
        return result

    # Anything above the header row is the title block; keep it, it usually holds
    # the customer's own job number ("BILL OF MATERIALS BASE  25-424").
    hdr = _find_header_row(best_rows)
    if not hdr:
        result["error"] = "no BOM header row found"
        result["title"] = " | ".join(v for v in best_rows[0] if v) if best_rows else ""
        return result
    result["header_row"] = hdr["row"]
    result["title"] = " ".join(
        v.strip() for row in best_rows[: hdr["row"]] for v in row if (v or "").strip()
    )[:300]

    section = ""
    for r in range(hdr["row"] + 1, len(best_rows)):
        row = best_rows[r]

        def col(key: str) -> str:
            idx = hdr.get(key)
            if idx is None or idx >= len(row):
                return ""
            return (row[idx] or "").strip()

        desc = col("desc")
        detail = col("detail")
        vendor = col("vendor")
        comment = col("comment")
        qty_text = col("qty")

        # A row with a description but nothing else, in caps, is a section banner
        # ("FASTENERS"). It changes how following rows are read.
        nonempty = [v for v in row if (v or "").strip()]
        if desc and len(nonempty) <= 2 and desc.upper() == desc and not qty_text:
            section = desc.upper()
            result["notes"].append({"row": r, "text": desc, "kind": "section"})
            continue
        if not desc and not detail:
            continue

        # Material may sit in a labelled column, or loose in any cell on the row
        # ("#7 steel" under "Addt'l Comments" -- Module6121.bas:13431).
        mat_text = col("material")
        grade = normalize_steel(mat_text)
        raw_material = mat_text
        if not grade:
            for cell in row:
                g = normalize_steel(cell)
                if g:
                    grade, raw_material = g, (cell or "").strip()
                    break

        # Size from dedicated columns first, then from the description text.
        # Tempcraft BOMs put it in Lth./Wth./Hgt. columns and leave the
        # description a bare name ("Top Holder Block"), so text-only parsing
        # found no dims, every steel row fell through to "note", and those jobs
        # reported zero plates.
        dims = _dims_from_columns(row, hdr) or parse_inch_dims(desc) or parse_inch_dims(detail)
        qty = _row_qty(qty_text)

        # An explicit TYPE column is authoritative -- Material / Purchase /
        # Outsource is the customer telling us the kind directly, which beats
        # inferring it from whether a vendor cell happens to be filled in.
        type_text = col("type").upper()
        kind = ""
        if type_text:
            if type_text.startswith("MATERIAL"):
                kind = "plate"
            elif type_text.startswith("PURCHASE"):
                kind = "purchased"
            elif type_text.startswith("OUTSOURCE"):
                kind = "purchased"
        if not kind:
            kind = _classify_row(detail, desc, vendor, comment, dims, grade)
        # A FASTENERS banner makes the rows under it fasteners even when their
        # names carry no shcs/bhcs prefix -- but only the rows that are actually
        # line items. Gating on a quantity keeps trailing prose out ("Included
        # w/ custom moldbase" is the last row of this BOM and is a note, not a
        # part to buy).
        if section.startswith("FASTENER") and kind == "note" and qty > 0:
            kind = "fastener"

        # The include column: Hewitt marks every line to be supplied with an X in
        # an unlabelled column left of the detail name. Blank means "not ours".
        # Bound the scan at the detail column so a description reading "X" or a
        # quantity cannot be mistaken for the mark.
        mark_limit = hdr["detail"] if hdr.get("detail") is not None else hdr["desc"]
        include_flag = None
        for cell in row[:mark_limit]:
            if (cell or "").strip().upper() == "X":
                include_flag = True
                break
        if include_flag is None and (detail or desc):
            include_flag = False

        item = {
            "row": r,
            "detail": detail,
            "description": desc,
            "qty": qty,
            "qty_text": qty_text,
            "vendor_part": vendor,
            "comment": comment,
            "material_raw": raw_material,
            "grade": grade,
            "grade_label": GRADE_LABELS.get(grade, ""),
            "marked_include": include_flag,
            "section": section,
        }
        if dims:
            item["dims"] = dims

        if kind == "plate":
            result["plates"].append(item)
        elif kind == "purchased":
            vend, pn = _split_vendor_part(vendor)
            item["vendor"] = vend
            item["part_number"] = pn
            result["purchased"].append(item)
        elif kind == "fastener":
            result["fasteners"].append(item)
        else:
            result["notes"].append({"row": r, "text": (desc or detail)[:160], "kind": "row"})
    return result


def _split_vendor_part(text: str) -> tuple[str, str]:
    """"Progressive, LP100L3.25" -> ("Progressive", "LP100L3.25").

    The third return case matters: estimators write instructions in the vendor
    column too. 8821-a00 on the 25-424 BOM reads "Sold as a set with 8820 ^^^^^",
    which is a note about the line above, and treating it as a vendor name put a
    fake supplier in the brief's vendor tally.
    """
    s = (text or "").strip()
    if not s:
        return "", ""
    if "," in s:
        vend, _, pn = s.partition(",")
        return vend.strip(), pn.strip().lstrip("#").strip()
    m = re.match(r"([A-Za-z][A-Za-z\-\s/&.]{2,})\s+([A-Z0-9][A-Z0-9.\-/]{2,})$", s)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    # No part number and reads like a sentence -> it is a note, not a vendor.
    if len(s.split()) > 3 or re.search(r"\^{2,}|\bsee\b|\bsold\b|\bincluded\b", s, re.I):
        return "", ""
    return s, ""


# ===========================================================================
# BOM <-> CAD reconciliation
# ===========================================================================

def _detail_key(text: str) -> str:
    """Match BOM detail names to STEP product names across prefixes.

    STEP calls the part "25-424--1000-a00"; the BOM calls it "1000-a00". Strip
    any job prefix and normalise separators so they collide on one key.
    """
    s = (text or "").strip().upper().replace(" ", "")
    s = re.sub(r"[\u2013\u2014]", "-", s)
    if "--" in s:
        s = s.rsplit("--", 1)[-1]
    # Strip SolidWorks copy suffixes: re-importing a job leaves "_1", "_1_2" on
    # the part file, so C17880's CAD carries "24-258--1200-a00_1-1" for BOM row
    # "1200-a00". Without this the keys were "1200-A00_1" vs "1200-A00" and a job
    # whose CAD *did* name every plate matched nothing at all.
    s = re.sub(r"(?:_\d+)+$", "", s)
    s = s.strip("-_")
    return s


def reconcile(bom: dict, cad_files: list[dict]) -> dict:
    """Compare what the BOM lists against what the CAD actually contains."""
    out = {
        "compared": False,
        "cad_file": "",
        "matched": [],
        "bom_only": [],
        "cad_only": [],
        "qty_mismatch": [],
        "note": "",
    }
    cad = next((c for c in cad_files if c.get("products")), None)
    if not cad:
        out["note"] = "No CAD part list available to compare against."
        return out
    out["compared"] = True
    out["cad_file"] = cad.get("file", "")

    cad_by_key: dict[str, dict] = {}
    for name in cad.get("products", []):
        cad_by_key[_detail_key(name)] = {
            "product": name,
            "instances": cad.get("instance_counts", {}).get(name, 0),
        }

    bom_items = (
        [("plate", i) for i in bom.get("plates", [])]
        + [("purchased", i) for i in bom.get("purchased", [])]
        + [("fastener", i) for i in bom.get("fasteners", [])]
    )
    seen_keys: set[str] = set()
    for kind, item in bom_items:
        key = _detail_key(item.get("detail") or "")
        if not key:
            continue
        hit = cad_by_key.get(key)
        if hit:
            seen_keys.add(key)
            rec = {
                "kind": kind,
                "detail": item.get("detail"),
                "product": hit["product"],
                "bom_qty": item.get("qty", 0),
                "cad_instances": hit["instances"],
            }
            out["matched"].append(rec)
            # An instance count of 0 means the product exists but this file's
            # assembly tree never places it -- not a real quantity conflict.
            if hit["instances"] and item.get("qty") and hit["instances"] != item["qty"]:
                out["qty_mismatch"].append(rec)
        else:
            out["bom_only"].append({
                "kind": kind,
                "detail": item.get("detail"),
                "description": item.get("description"),
                "qty": item.get("qty", 0),
                "marked_include": item.get("marked_include"),
                "comment": item.get("comment", ""),
            })

    assemblies = set(cad.get("assemblies", []))
    for key, hit in cad_by_key.items():
        if key in seen_keys:
            continue
        # Assembly nodes are containers, not BOM lines.
        if hit["product"] in assemblies:
            continue
        out["cad_only"].append({"product": hit["product"], "instances": hit["instances"]})
    out["cad_assemblies"] = sorted(assemblies)
    return out


# ===========================================================================
# Facts
# ===========================================================================

# Delivery urgency, computed rather than asked.
#
# The model got this wrong on the very first email: "I'm looking for the best
# possible delivery" came back as urgency "normal", which is the opposite of what
# the customer meant -- that sentence is how a buyer says delivery is being
# competed on. Phrase matching is exact and free, so it decides, and the model's
# answer is only a fallback for wording not listed here.
_EXPEDITE_PHRASES = (
    "best possible delivery", "best delivery", "shortest possible", "soonest",
    "as soon as possible", "asap", "a.s.a.p", "expedite", "expedited", "rush",
    "quick turn", "quickest", "urgent", "hot job", "need it fast", "lead time is critical",
)
_DATE_RE = re.compile(
    r"\b(?:\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?"
    r"|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*\.?\s+\d{1,2}(?:\s*,?\s*\d{4})?)\b",
    re.I,
)
_DUE_LABEL_RE = re.compile(
    r"(?:ship\s*date|need(?:ed)?\s*by|due\s*(?:date)?|delivery\s*(?:date)?|required\s*by|want\s*it\s*by)"
    r"\s*[:#]?\s*([^\n\r.;]{2,60})",
    re.I,
)


def delivery_signals(text: str) -> dict:
    """Urgency and the customer's own delivery wording, from the email text."""
    out = {"urgency": "unstated", "requested": "", "matched_phrase": ""}
    blob = (text or "").strip()
    if not blob:
        return out
    low = blob.lower()

    for phrase in _EXPEDITE_PHRASES:
        if phrase in low:
            out["urgency"] = "expedite"
            out["matched_phrase"] = phrase
            # Quote the customer's clause, not just the keyword.
            m = re.search(r"[^.!?\n]*" + re.escape(phrase) + r"[^.!?\n]*", blob, re.I)
            out["requested"] = (m.group(0).strip() if m else phrase)[:120]
            break

    m = _DUE_LABEL_RE.search(blob)
    if m:
        label_val = m.group(1).strip()
        if not out["requested"]:
            out["requested"] = label_val[:120]
        if out["urgency"] == "unstated":
            out["urgency"] = "normal"
    elif out["urgency"] == "unstated":
        m = _DATE_RE.search(blob)
        if m:
            out["requested"] = m.group(0)
            out["urgency"] = "normal"
    return out


def gather_facts(subject: str, body: str, attachment_paths: list[Path],
                 from_addr: str = "") -> dict:
    """Everything we can establish without a model. Authoritative."""
    facts: dict = {
        "email": {
            "from": from_addr,
            "subject": subject or "",
            "body": (body or "").strip(),
            "body_chars": len(body or ""),
        },
        "attachments": [],
        "cad": [],
        "boms": [],
        "documents": [],
        "archives": [],
        "unreadable": [],
    }

    for path in attachment_paths:
        path = Path(path)
        ext = path.suffix.lower()
        entry = {
            "name": path.name,
            "ext": ext,
            "bytes": path.stat().st_size if path.exists() else 0,
            "exists": path.exists(),
            "role": "other",
        }
        if not path.exists():
            entry["role"] = "missing"
            facts["attachments"].append(entry)
            facts["unreadable"].append({"name": path.name, "why": "file not found"})
            continue

        if ext in CAD_EXTS:
            entry["role"] = "cad"
            facts["cad"].append(read_cad(path))
        elif ext in BOM_EXTS:
            entry["role"] = "bom"
            facts["boms"].append(interpret_bom(read_spreadsheet(path), path.name))
        elif ext == ".zip":
            entry["role"] = "archive"
            facts["archives"].append(read_zip(path))
        elif ext in DOC_EXTS:
            entry["role"] = "document"
            text = read_pdf_text(path) if ext == ".pdf" else ""
            if not text and ext == ".txt":
                try:
                    text = path.read_text(encoding="utf-8", errors="replace")[:20000]
                except Exception:
                    text = ""
            if text:
                facts["documents"].append({"name": path.name, "text": text})
            else:
                facts["unreadable"].append({"name": path.name, "why": f"no text extracted from {ext}"})
        facts["attachments"].append(entry)

    # Reconcile the first real BOM against the CAD.
    bom = next((b for b in facts["boms"] if not b.get("error")), None)
    facts["reconciliation"] = reconcile(bom, facts["cad"]) if bom else {
        "compared": False, "note": "No readable BOM attached.",
        "matched": [], "bom_only": [], "cad_only": [], "qty_mismatch": [], "cad_file": "",
    }
    facts["delivery_signals"] = delivery_signals(
        f"{subject or ''}\n{body or ''}"
    )
    facts["totals"] = _totals(bom, facts)
    facts["flags"] = _flags(bom, facts)
    return facts


def _totals(bom: dict | None, facts: dict) -> dict:
    plates = bom.get("plates", []) if bom else []
    purchased = bom.get("purchased", []) if bom else []
    fasteners = bom.get("fasteners", []) if bom else []
    grades: dict[str, int] = {}
    for p in plates:
        g = p.get("grade") or "UNKNOWN"
        grades[g] = grades.get(g, 0) + max(1, p.get("qty") or 1)
    vendors: dict[str, int] = {}
    for p in purchased:
        v = p.get("vendor") or "unspecified"
        vendors[v] = vendors.get(v, 0) + 1
    return {
        "plate_lines": len(plates),
        "plate_pieces": sum(max(1, p.get("qty") or 1) for p in plates),
        "purchased_lines": len(purchased),
        "purchased_pieces": sum(max(1, p.get("qty") or 1) for p in purchased),
        "fastener_lines": len(fasteners),
        "fastener_pieces": sum(max(1, p.get("qty") or 1) for p in fasteners),
        "grades": grades,
        "vendors": vendors,
        "cad_products": sum(len(c.get("products", [])) for c in facts["cad"]),
        "cad_instances": sum(sum(c.get("instance_counts", {}).values()) for c in facts["cad"]),
    }


def _flags(bom: dict | None, facts: dict) -> list[dict]:
    """Things a human has to look at before this quote is trustworthy."""
    flags: list[dict] = []
    rec = facts.get("reconciliation", {})

    if not facts["cad"]:
        flags.append({"level": "blocker", "code": "no_cad",
                      "message": "No CAD file attached -- the macro has nothing to open."})
    if bom is None:
        flags.append({"level": "warn", "code": "no_bom",
                      "message": "No readable BOM attached; plate material and quantities "
                                 "will have to come from the CAD alone."})
    elif bom.get("error"):
        flags.append({"level": "warn", "code": "bom_unreadable",
                      "message": f"BOM '{bom.get('source')}' could not be read: {bom['error']}"})

    if bom:
        no_grade = [p["detail"] or p["description"] for p in bom.get("plates", []) if not p.get("grade")]
        if no_grade:
            flags.append({"level": "warn", "code": "grade_unknown",
                          "message": f"{len(no_grade)} plate line(s) have no recognisable steel grade: "
                                     + ", ".join(no_grade[:6]),
                          "items": no_grade})
        no_dims = [p["detail"] or p["description"] for p in bom.get("plates", []) if not p.get("dims")]
        if no_dims:
            flags.append({"level": "warn", "code": "dims_missing",
                          "message": f"{len(no_dims)} plate line(s) have no parsable size: "
                                     + ", ".join(no_dims[:6]),
                          "items": no_dims})
        unpriced_vendor = [p["detail"] for p in bom.get("purchased", []) if not p.get("part_number")]
        if unpriced_vendor:
            flags.append({"level": "info", "code": "no_part_number",
                          "message": f"{len(unpriced_vendor)} purchased line(s) carry no vendor part "
                                     "number, so they cannot be priced from the price list: "
                                     + ", ".join(unpriced_vendor[:6]),
                          "items": unpriced_vendor})

    if rec.get("compared"):
        if rec.get("bom_only"):
            detail = ", ".join(
                f"{i['detail']}" + (" (include column blank)" if i.get("marked_include") is False else "")
                for i in rec["bom_only"][:6]
            )
            flags.append({"level": "warn", "code": "bom_only",
                          "message": f"{len(rec['bom_only'])} BOM line(s) have no matching part in "
                                     f"{rec.get('cad_file')}: {detail}",
                          "items": [i["detail"] for i in rec["bom_only"]]})
        if rec.get("cad_only"):
            detail = ", ".join(i["product"] for i in rec["cad_only"][:6])
            flags.append({"level": "warn", "code": "cad_only",
                          "message": f"{len(rec['cad_only'])} CAD part(s) are not on the BOM: {detail}",
                          "items": [i["product"] for i in rec["cad_only"]]})
        if rec.get("qty_mismatch"):
            detail = ", ".join(
                f"{i['detail']} BOM {i['bom_qty']} vs CAD {i['cad_instances']}"
                for i in rec["qty_mismatch"][:6]
            )
            flags.append({"level": "warn", "code": "qty_mismatch",
                          "message": f"{len(rec['qty_mismatch'])} line(s) disagree on quantity: {detail}",
                          "items": [i["detail"] for i in rec["qty_mismatch"]]})

    for c in facts["cad"]:
        if c.get("truncated"):
            flags.append({"level": "warn", "code": "cad_truncated",
                          "message": f"{c['file']} is larger than the scan limit; its part list is partial."})
        if c.get("error"):
            flags.append({"level": "warn", "code": "cad_error",
                          "message": f"{c['file']}: {c['error']}"})
    for u in facts["unreadable"]:
        flags.append({"level": "info", "code": "attachment_unreadable",
                      "message": f"{u['name']}: {u['why']}"})
    return flags


# ===========================================================================
# The model
# ===========================================================================

# Every enum value is spelled out because a small model given a free-text field
# writes whatever it likes -- qwen3:1.7b left request_type empty and qwen3:8b
# answered "RFQ", neither of which merge_brief() accepts.
BRIEF_SCHEMA_HINT = """{
  "request_type": one of "new_quote","requote","revision","purchase_order","question","not_a_quote",
  "customer_company": <string>,
  "customer_contact": <string>,
  "job_reference": <string>,
  "what_is_being_quoted": <string>,
  "base_type": one of "custom","standard_catalog","insert_only","unclear",
  "requested_delivery": <string>,
  "urgency": one of "expedite","normal","unstated",
  "quote_delivery_with_price": true or false,
  "include_purchased_components": true or false,
  "include_fasteners": true or false,
  "bom_is_authoritative": true or false,
  "special_instructions": [<string>, ...],
  "open_questions": [<string>, ...]
}

Field notes (do NOT copy this text into the values -- use "" when unknown):
  job_reference        the customer's own job or RFQ number
  what_is_being_quoted one short sentence, shop language
  requested_delivery   the delivery date or phrase the customer used
  special_instructions up to 3 SHORT lines in the customer's own words
  open_questions       up to 2, only what blocks quoting, [] if none

Be brief. Every extra word costs seconds on this machine."""


def _ollama_models() -> list[str]:
    try:
        with urllib.request.urlopen(f"{OLLAMA_URL}/api/tags", timeout=6) as r:
            data = json.loads(r.read().decode("utf-8", "replace"))
        return [m.get("name", "") for m in data.get("models", [])]
    except Exception:
        return []


def pick_model(preferred: str = "") -> str:
    """Pick a model that is actually pulled on this machine."""
    if preferred:
        return preferred
    env = os.environ.get("CMS_EMAIL_AI_MODEL", "").strip()
    if env:
        return env
    available = _ollama_models()
    for name in MODEL_PREFERENCE:
        if name in available:
            return name
    # Any qwen beats nothing; otherwise first installed model.
    for name in available:
        if "qwen" in name.lower():
            return name
    return available[0] if available else MODEL_PREFERENCE[0]


def extract_json(text: str) -> dict:
    """Same defensive parse geometry_classifier/qwen_classify_xt_csv.py uses.

    Qwen emits terminal control codes and <think> blocks even through the HTTP
    API when thinking is on, and wraps JSON in fences about a third of the time.
    """
    text = re.sub(r"\x1b\[[0-9;?]*[A-Za-z]", "", text or "")
    text = text.replace("\b", "")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)
    text = re.sub(r"<think>.*?</think>", "", text, flags=re.S | re.I)
    text = re.sub(r"```(?:json)?", "", text, flags=re.I)
    text = re.sub(r"(?i)done thinking\.\s*", "", text).strip()
    if text.startswith("{"):
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass
    # Greedy to the last brace: the model sometimes narrates after the object.
    m = re.search(r"\{.*\}", text, flags=re.S)
    if not m:
        raise ValueError("no JSON object in model output")
    return json.loads(m.group(0))


def _post_ollama(payload: dict, timeout: int) -> dict:
    req = urllib.request.Request(
        f"{OLLAMA_URL}/api/generate",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=timeout) as r:
        return json.loads(r.read().decode("utf-8", "replace"))


def ask_model(prompt: str, model: str = "", timeout: int = LLM_TIMEOUT_S) -> dict:
    """Run the prompt through Ollama. Never raises; reports failure in-band."""
    model = pick_model(model)
    out = {"ok": False, "model": model, "raw": "", "error": "", "seconds": 0.0,
           "transport": "http"}
    # ~4 chars/token, plus headroom for the answer, rounded to a power of two and
    # clamped. Sizing to the prompt keeps the compact path off a 24k context it
    # does not need.
    want = max(NUM_CTX_MIN, min(NUM_CTX_MAX, ((len(prompt) // 4) + NUM_PREDICT + 512)))
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        # Thinking burns minutes on a CPU box and the answer is a fixed schema,
        # so turn it off where the server honours the flag.
        "think": False,
        # Keep the weights resident between emails. Loading qwen3:1.7b costs
        # ~15-20 s of every cold call, which is most of the wait when the answer
        # itself takes ~30 s. An inbox with five RFQs pays that once, not five
        # times. Ollama's default is 5 minutes, which a slow batch can fall off.
        "keep_alive": KEEP_ALIVE,
        "options": {"temperature": 0, "top_p": 0.9, "num_ctx": want,
                    "num_predict": NUM_PREDICT},
    }
    try:
        data = _post_ollama(payload, timeout)
    except urllib.error.HTTPError as e:
        # Older Ollama builds 400 on an unknown "think" key; retry without it.
        body = ""
        try:
            body = e.read().decode("utf-8", "replace")[:300]
        except Exception:
            pass
        if e.code == 400:
            payload.pop("think", None)
            try:
                data = _post_ollama(payload, timeout)
            except Exception as e2:
                out["error"] = f"Ollama HTTP {e.code}: {body or e2}"
                return out
        else:
            out["error"] = f"Ollama HTTP {e.code}: {body}"
            return out
    except urllib.error.URLError as e:
        out["error"] = (f"Ollama not reachable at {OLLAMA_URL} ({e.reason}). "
                        "Start it with 'ollama serve'.")
        return out
    except Exception as e:
        out["error"] = f"{type(e).__name__}: {e}"
        return out

    out["raw"] = data.get("response", "")
    out["seconds"] = round((data.get("total_duration") or 0) / 1e9, 1)
    try:
        out["parsed"] = extract_json(out["raw"])
        out["ok"] = True
    except Exception as e:
        out["error"] = f"model output was not JSON: {e}"
    return out


SHOP_CONTEXT = """You read incoming RFQ emails for Custom Mold Services, a shop that builds custom
injection mold bases. An estimator acts on your summary.

A mold base is a stack of steel plates (top clamp, A/B or retainer plates,
support plate, ejector and ejector retainer plates, bottom clamp, rails) plus
purchased hardware (leader pins, bushings, return pins, support pillars, stop
discs, side locks, safety straps) and fasteners. "#1/#2/#3/#5/#7" are DME grade
codes for A-36/4140/P20/H-13/420-SS. "Custom" means cut from plate to the
customer's sizes, not a catalog base."""


def build_prompt(facts: dict) -> str:
    """The fast prompt: interpret the customer's words, nothing else.

    Deliberately does NOT include the BOM rows or the CAD part list. Those are
    already resolved in `facts` -- feeding them to the model added ~1400 tokens
    of prompt (minutes of CPU at this box's speed) and asked a 2B-parameter model
    to re-derive numbers that were already exact. What is left here is the only
    part a model is actually needed for: a few sentences of English.
    """
    email = facts["email"]
    tot = facts.get("totals", {})
    rec = facts.get("reconciliation", {})

    inventory = ", ".join(
        f"{a['name']} ({a['role']})" for a in facts.get("attachments", [])
    ) or "none"

    # One line of computed context so the model's sentence is not vague, plus the
    # discrepancies, which are the only computed facts it needs to reason about.
    computed = (
        f"{tot.get('plate_pieces', 0)} steel plates, "
        f"{tot.get('purchased_pieces', 0)} purchased components, "
        f"{tot.get('fastener_pieces', 0)} fasteners, "
        f"{tot.get('cad_products', 0)} distinct CAD parts"
    )
    issues = []
    for i in rec.get("bom_only", []):
        issues.append(f"BOM line {i.get('detail')} is not in the CAD"
                      + (" and its include column is blank" if i.get("marked_include") is False else "")
                      + (f" ({i['comment']})" if i.get("comment") else ""))
    for i in rec.get("cad_only", []):
        issues.append(f"CAD part {i.get('product')} is not on the BOM")
    for i in rec.get("qty_mismatch", []):
        issues.append(f"{i.get('detail')}: BOM says {i.get('bom_qty')}, CAD has {i.get('cad_instances')}")
    for f in facts.get("flags", []):
        if f.get("level") == "blocker":
            issues.append(f["message"])

    return f"""{SHOP_CONTEXT}

Read the email below and return ONE JSON object, no prose, exactly this shape:

{BRIEF_SCHEMA_HINT}

Rules:
- Base your answer on what the customer WROTE. Do not invent numbers.
- Quote the customer's own words in special_instructions.
- open_questions is only for what genuinely blocks quoting. If the attachments
  answer it, it is not an open question. Usually [].

=== EMAIL ===
From: {email.get('from') or '(unknown)'}
Subject: {email.get('subject') or '(none)'}
Attachments: {inventory}

{email.get('body', '')[:3000] or '(empty body)'}
=== END EMAIL ===

Already established from the attached files (do not recount, do not restate):
{computed}
{("Discrepancies found: " + "; ".join(issues)) if issues else "No discrepancies between the BOM and the CAD."}

JSON:"""


def build_prompt_deep(facts: dict) -> str:
    """The full digest: every BOM row and CAD part, for a big model on a fast box.

    Kept because it is the right prompt when the hardware can afford it -- it
    lets the model reason about individual lines. Reach for it with --deep and a
    pinned model; it is far too slow to be the default here.
    """
    email = facts["email"]
    bom = next((b for b in facts["boms"] if not b.get("error")), None)
    rec = facts.get("reconciliation", {})
    tot = facts.get("totals", {})

    lines: list[str] = []
    lines.append("=== EMAIL ===")
    lines.append(f"From: {email.get('from') or '(unknown)'}")
    lines.append(f"Subject: {email.get('subject')}")
    lines.append("Body:")
    lines.append(email.get("body", "")[:6000] or "(empty)")

    lines.append("")
    lines.append("=== ATTACHMENTS ===")
    for a in facts["attachments"]:
        kb = a["bytes"] // 1024
        lines.append(f"- {a['name']} [{a['role']}, {kb} KB]")

    for c in facts["cad"]:
        lines.append("")
        lines.append(f"=== CAD: {c['file']} ({c.get('format')}) ===")
        if c.get("originating_system"):
            lines.append(f"Exported by: {c['originating_system']}")
        counts = c.get("instance_counts", {})
        lines.append(f"Distinct parts: {len(c.get('products', []))}; "
                     f"placed instances: {sum(counts.values())}; solids: {c.get('solid_count')}")
        for name in c.get("products", [])[:120]:
            n = counts.get(name, 0)
            lines.append(f"  {name}" + (f"  x{n}" if n else "  (assembly node)"))

    if bom:
        lines.append("")
        lines.append(f"=== BOM: {bom['source']} / sheet '{bom['sheet']}' ===")
        if bom.get("title"):
            lines.append(f"Title block: {bom['title'][:200]}")
        lines.append(f"-- STEEL PLATES ({len(bom['plates'])} lines) --")
        for p in bom["plates"]:
            d = p.get("dims") or {}
            size = (f"{d.get('thickness')} x {d.get('width')} x {d.get('length')}"
                    if d.get("kind") == "plate" else "size?")
            lines.append(
                f"  {p['detail'] or '(no detail no.)'} | qty {p['qty']} | {size} | "
                f"material '{p['material_raw']}' -> {p['grade'] or 'UNRECOGNISED'} | {p['description']}"
            )
        lines.append(f"-- PURCHASED COMPONENTS ({len(bom['purchased'])} lines) --")
        for p in bom["purchased"]:
            lines.append(
                f"  {p['detail'] or '(no detail no.)'} | qty {p['qty']} | "
                f"{p.get('vendor') or 'vendor?'} {p.get('part_number') or 'part#?'} | {p['description']}"
                + (f" | {p['comment']}" if p.get("comment") else "")
            )
        lines.append(f"-- FASTENERS ({len(bom['fasteners'])} lines) --")
        for p in bom["fasteners"]:
            mark = "" if p.get("marked_include") else "  [include column BLANK]"
            lines.append(f"  {p['detail']} | qty {p['qty']} | {p['description']}"
                         + (f" | {p['comment']}" if p.get("comment") else "") + mark)
        if bom.get("notes"):
            lines.append("-- OTHER ROWS --")
            for n in bom["notes"][:20]:
                lines.append(f"  {n['text']}")

    for d in facts.get("documents", []):
        lines.append("")
        lines.append(f"=== DOCUMENT: {d['name']} ===")
        lines.append(d["text"][:4000])

    lines.append("")
    lines.append("=== RECONCILIATION (computed, trust these) ===")
    if rec.get("compared"):
        lines.append(f"BOM lines matched to CAD parts: {len(rec.get('matched', []))}")
        lines.append(f"On BOM but not in CAD: "
                     + (", ".join(f"{i['detail']} (qty {i['qty']}"
                                  + (", include column blank" if i.get("marked_include") is False else "")
                                  + f", {i.get('comment') or 'no comment'})"
                                  for i in rec.get("bom_only", [])) or "none"))
        lines.append(f"In CAD but not on BOM: "
                     + (", ".join(f"{i['product']} x{i['instances']}"
                                  for i in rec.get("cad_only", [])) or "none"))
        lines.append(f"Quantity disagreements: "
                     + (", ".join(f"{i['detail']} BOM {i['bom_qty']} vs CAD {i['cad_instances']}"
                                  for i in rec.get("qty_mismatch", [])) or "none"))
    else:
        lines.append(rec.get("note", "not compared"))

    lines.append("")
    lines.append("=== COUNTS (computed, trust these) ===")
    lines.append(json.dumps(tot, separators=(",", ":")))
    if facts.get("flags"):
        lines.append("")
        lines.append("=== FLAGS ALREADY RAISED (do not repeat verbatim) ===")
        for f in facts["flags"]:
            lines.append(f"  [{f['level']}] {f['message']}")

    digest = "\n".join(lines)

    return f"""{SHOP_CONTEXT}

Your job is INTERPRETATION, not arithmetic. Every count, dimension, quantity,
grade and part number in the digest was computed from the actual files and is
correct. Do not restate, recount, or "correct" them. Read what the customer
ASKED FOR and say what that means for how we quote.

Rules:
- Quote the customer's own words in special_instructions where it matters.
- open_questions is only for things that genuinely block quoting. If the files
  answer a question, it is not an open question.
- If a BOM line is absent from the CAD and the customer said to include what is
  in the design, say which one wins -- do not leave it ambiguous.
- Output ONE JSON object, no prose around it, matching exactly this shape:

{BRIEF_SCHEMA_HINT}

--- DIGEST ---
{digest}
--- END DIGEST ---

JSON:"""


# ===========================================================================
# Merge: facts win over the model
# ===========================================================================

_ALLOWED_REQUEST_TYPES = {"new_quote", "requote", "revision", "purchase_order",
                          "question", "not_a_quote"}
_ALLOWED_ACTIONS = {"quote_now", "quote_with_assumptions", "ask_customer_first",
                    "route_to_human"}


_ALLOWED_BASE_TYPES = {"custom", "standard_catalog", "insert_only", "unclear"}
_ALLOWED_URGENCY = {"expedite", "normal", "unstated"}


def _pick_bool(*candidates) -> bool:
    """First real boolean in the list. Small models answer "yes"/"true" as text."""
    for c in candidates:
        if isinstance(c, bool):
            return c
        if isinstance(c, str) and c.strip().lower() in ("true", "yes", "y", "1"):
            return True
        if isinstance(c, str) and c.strip().lower() in ("false", "no", "n", "0"):
            return False
    return False


# Field descriptions the model likes to copy back verbatim instead of filling in.
# qwen3:1.7b returned "as the customer wrote it" as the requested delivery date on
# the first run, which would have gone onto a brief as if the customer had said it.
_SCHEMA_ECHO = (
    "as the customer wrote it", "exactly as written", "exactly as the customer",
    "or \"\"", "one short sentence", "shop language", "date or phrase",
    "max 4 short lines", "quote the customer's words", "person's name",
    "company name", "the customer's own", "if none", "2 sentences",
)


def _clean_model_text(value) -> str:
    """Drop a value that is really the schema's own instruction text."""
    s = str(value or "").strip()
    if not s:
        return ""
    low = s.lower()
    for echo in _SCHEMA_ECHO:
        if echo in low:
            return ""
    return s


def _one_of(value, allowed: set, default: str) -> str:
    v = str(value or "").strip().lower().replace(" ", "_").replace("-", "_")
    return v if v in allowed else default


def _as_list(value, limit: int = 12) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        value = [value]
    out = []
    for v in value:
        s = _clean_model_text(v)
        if s and s.lower() not in ("none", "n/a", "null"):
            out.append(s[:400])
    return out[:limit]


def merge_brief(facts: dict, llm: dict) -> dict:
    """Assemble the brief. Deterministic facts override the model everywhere."""
    parsed = llm.get("parsed") or {}
    tot = facts.get("totals", {})
    rec = facts.get("reconciliation", {})
    flags = list(facts.get("flags", []))

    request_type = str(parsed.get("request_type", "")).strip().lower().replace(" ", "_")
    if request_type not in _ALLOWED_REQUEST_TYPES:
        # Small models answer "rfq" or "quote" here however hard the enum is
        # spelled out; map the near-misses rather than discarding a good read.
        alias = {"rfq": "new_quote", "quote": "new_quote", "quote_request": "new_quote",
                 "new": "new_quote", "po": "purchase_order", "order": "purchase_order",
                 "re_quote": "requote", "inquiry": "question"}.get(request_type)
        request_type = alias or ("new_quote" if llm.get("ok") else "unknown")

    # recommended_action is no longer asked of the model: it is a pure function of
    # the flags, which are computed. Deriving it here removes a field the model
    # got wrong often enough to need overriding anyway.
    action = ""

    blockers = [f for f in flags if f["level"] == "blocker"]
    warns = [f for f in flags if f["level"] == "warn"]
    if blockers:
        action = "route_to_human"
    elif warns:
        action = "quote_with_assumptions"
    else:
        action = "quote_now"

    # The compact schema is flat, so the model cannot omit a nested object and
    # take three fields down with it. Accept the nested shape too, for briefs
    # written by the --deep prompt or an older cached file.
    nested_cust = parsed.get("customer") if isinstance(parsed.get("customer"), dict) else {}
    customer = {
        "company": _clean_model_text(parsed.get("customer_company")
                                     or nested_cust.get("company")),
        "contact": _clean_model_text(parsed.get("customer_contact")
                                     or nested_cust.get("contact")),
        "email": nested_cust.get("email", ""),
        "phone": nested_cust.get("phone", ""),
    }
    if not customer["email"] and facts["email"].get("from"):
        m = re.search(r"[\w.+-]+@[\w-]+\.[\w.-]+", facts["email"]["from"])
        if m:
            customer["email"] = m.group(0)
    # The model sometimes fills contact with the address, which then prints twice.
    if "@" in customer["contact"]:
        if not customer["email"]:
            customer["email"] = customer["contact"]
        customer["contact"] = ""

    nested_scope = parsed.get("scope") if isinstance(parsed.get("scope"), dict) else {}
    scope = {
        # Whether the BOM HAS these lines is a fact; whether the customer wants
        # them priced is the model's call. Fall back to presence when it is silent.
        "steel_plates": tot.get("plate_lines", 0) > 0,
        "purchased_components": _pick_bool(
            parsed.get("include_purchased_components"),
            nested_scope.get("purchased_components"),
            tot.get("purchased_lines", 0) > 0),
        "fasteners": _pick_bool(
            parsed.get("include_fasteners"),
            nested_scope.get("fasteners"),
            tot.get("fastener_lines", 0) > 0),
        "notes": str(nested_scope.get("notes", ""))[:400],
    }

    # Computed delivery wins; the model only fills gaps. See delivery_signals.
    sig = facts.get("delivery_signals") or {}
    nested_del = parsed.get("delivery") if isinstance(parsed.get("delivery"), dict) else {}
    model_requested = _clean_model_text(
        parsed.get("requested_delivery") or nested_del.get("requested", ""))
    model_urgency = _one_of(parsed.get("urgency") or nested_del.get("urgency"),
                            _ALLOWED_URGENCY, "unstated")
    delivery = {
        "requested": sig.get("requested") or model_requested,
        "urgency": sig.get("urgency") if sig.get("urgency") != "unstated" else model_urgency,
        "quote_delivery_with_price": _pick_bool(
            parsed.get("quote_delivery_with_price"),
            nested_del.get("quote_delivery_with_price"),
            sig.get("urgency") == "expedite"),
    }
    if sig.get("urgency") == "expedite" and model_urgency not in ("expedite", "unstated"):
        flags.append({
            "level": "info", "code": "urgency_corrected",
            "message": f"Model called delivery '{model_urgency}', but the email says "
                       f"\"{sig.get('matched_phrase')}\" -- treated as expedite.",
        })

    brief = {
        "schema_version": 1,
        "request_type": request_type,
        "is_quote_request": bool(parsed.get("is_quote_request", request_type not in
                                            ("not_a_quote", "question", "unknown"))),
        "customer": {
            "company": str(customer.get("company", ""))[:120],
            "contact": str(customer.get("contact", ""))[:120],
            "email": str(customer.get("email", ""))[:160],
            "phone": str(customer.get("phone", ""))[:60],
        },
        "job_reference": _clean_model_text(parsed.get("job_reference"))[:80],
        "what_is_being_quoted": _clean_model_text(parsed.get("what_is_being_quoted"))[:400],
        "base_type": _one_of(parsed.get("base_type"), _ALLOWED_BASE_TYPES, "unclear"),
        "scope": scope,
        "bom_is_authoritative": _pick_bool(parsed.get("bom_is_authoritative"),
                                           bool(facts["boms"])),
        "delivery": {
            "requested": str(delivery.get("requested", ""))[:120],
            "urgency": _one_of(delivery.get("urgency"), _ALLOWED_URGENCY, "unstated"),
            "quote_delivery_with_price": bool(delivery.get("quote_delivery_with_price")),
        },
        "special_instructions": _as_list(parsed.get("special_instructions"), 4),
        "open_questions": _as_list(parsed.get("open_questions"), 3),
        "risks": _as_list(parsed.get("risks"), 4),
        "recommended_action": action,
        "reasoning": _clean_model_text(parsed.get("reasoning"))[:1200],

        # --- computed, authoritative ---
        "counts": tot,
        "flags": flags,
        "reconciliation": {
            "compared": rec.get("compared", False),
            "cad_file": rec.get("cad_file", ""),
            "matched": len(rec.get("matched", [])),
            "bom_only": rec.get("bom_only", []),
            "cad_only": rec.get("cad_only", []),
            "qty_mismatch": rec.get("qty_mismatch", []),
        },
        "line_items": _line_items(facts),
        "attachments": facts.get("attachments", []),
        "llm": {
            "ok": llm.get("ok", False),
            "model": llm.get("model", ""),
            "seconds": llm.get("seconds", 0.0),
            "error": llm.get("error", ""),
        },
    }
    if not llm.get("ok"):
        brief["degraded"] = True
        brief["flags"].append({
            "level": "warn", "code": "llm_unavailable",
            "message": "The local model did not answer, so this brief is the "
                       f"computed facts only. {llm.get('error', '')}".strip(),
        })
        if not brief["what_is_being_quoted"]:
            brief["what_is_being_quoted"] = _fallback_summary(facts)
    return brief


def _line_items(facts: dict) -> dict:
    """The BOM's line items, as computed. This is what the estimator checks.

    DIMENSIONS HERE ARE THE CUSTOMER'S STATED SIZES, NOT QUOTE SIZES.
    The BOM is reference material and its sizes are not always right, so they are
    named `bom_*` rather than `thickness`/`width`/`length` -- a consumer that
    wants a dimension has to reach for a key that says where it came from, and
    cannot get one by accident.

    They are still worth carrying: this brief is built when the RFQ arrives,
    before SolidWorks has opened anything, so the BOM is the only size
    information in existence at that moment. Once the macro runs,
    XT_Export_CAD_Dimensions.csv is the authority and these are superseded.
    """
    bom = next((b for b in facts["boms"] if not b.get("error")), None)
    if not bom:
        return {"plates": [], "purchased": [], "fasteners": []}

    def plate(p):
        d = p.get("dims") or {}
        return {
            "detail": p.get("detail", ""),
            "name": p.get("description", ""),
            "qty": p.get("qty", 0),
            # bom_* prefix: stated by the customer, superseded by the CAD export.
            "bom_thickness": d.get("thickness"),
            "bom_width": d.get("width"),
            "bom_length": d.get("length"),
            "material_raw": p.get("material_raw", ""),
            "grade": p.get("grade", ""),
            "grade_label": p.get("grade_label", ""),
            # Whether the GRADE is usable, which is what the BOM is trusted for.
            # Says nothing about the sizes.
            "grade_priceable": bool(p.get("grade") in PRICEABLE_GRADES),
        }

    def bought(p):
        return {
            "detail": p.get("detail", ""),
            "name": p.get("description", ""),
            "qty": p.get("qty", 0),
            "vendor": p.get("vendor", ""),
            "part_number": p.get("part_number", ""),
            "comment": p.get("comment", ""),
            "priceable": bool(p.get("part_number")),
        }

    def fast(p):
        return {
            "detail": p.get("detail", ""),
            "name": p.get("description", ""),
            "qty": p.get("qty", 0),
            "comment": p.get("comment", ""),
            "marked_include": p.get("marked_include"),
        }

    return {
        "plates": [plate(p) for p in bom.get("plates", [])],
        "purchased": [bought(p) for p in bom.get("purchased", [])],
        "fasteners": [fast(p) for p in bom.get("fasteners", [])],
    }


def _fallback_summary(facts: dict) -> str:
    """A usable one-liner when the model is down."""
    t = facts.get("totals", {})
    parts = []
    if t.get("plate_lines"):
        grades = ", ".join(f"{n}x {GRADE_LABELS.get(g, g)}" for g, n in
                           sorted(t.get("grades", {}).items(), key=lambda kv: -kv[1]))
        parts.append(f"{t['plate_pieces']} steel plate(s) [{grades}]")
    if t.get("purchased_lines"):
        parts.append(f"{t['purchased_pieces']} purchased component(s) "
                     f"on {len(t.get('vendors', {}))} vendor(s)")
    if t.get("fastener_lines"):
        parts.append(f"{t['fastener_pieces']} fastener(s)")
    if not parts:
        return "Could not determine scope from the attachments."
    return "Mold base: " + "; ".join(parts) + "."


# ===========================================================================
# Public entry point
# ===========================================================================

def analyze(subject: str, body: str, attachment_paths, from_addr: str = "",
            model: str = "", use_llm: bool = True, deep: bool = False,
            timeout: int = LLM_TIMEOUT_S) -> dict:
    """Read an RFQ and return the quote brief. Never raises on bad input."""
    paths = [Path(p) for p in (attachment_paths or [])]
    facts = gather_facts(subject, body, paths, from_addr=from_addr)
    if use_llm:
        prompt = build_prompt_deep(facts) if deep else build_prompt(facts)
        llm = ask_model(prompt, model=model, timeout=timeout)
    else:
        llm = {"ok": False, "model": "", "error": "LLM disabled by caller",
               "seconds": 0.0}
    brief = merge_brief(facts, llm)
    brief["source"] = _source_of(facts)
    return brief


def analyze_dir(directory, subject: str = "", body: str = "", from_addr: str = "",
                model: str = "", use_llm: bool = True) -> dict:
    """Analyze every file in an attachment folder (what quote_from_message saves)."""
    d = Path(directory)
    files = sorted(p for p in d.glob("*") if p.is_file()) if d.is_dir() else []
    return analyze(subject, body, files, from_addr=from_addr, model=model, use_llm=use_llm)


def _source_of(facts: dict) -> dict:
    """What the brief was built from. `body` is kept so a re-run needs no mailbox."""
    return {
        "subject": facts["email"]["subject"],
        "from": facts["email"]["from"],
        "body": facts["email"]["body"][:8000],
        "attachment_names": [a["name"] for a in facts["attachments"]],
    }


def analyze_async(target_dir, subject: str, body: str, attachment_paths,
                  from_addr: str = "", model: str = "", deep: bool = False) -> dict:
    """Save the computed brief now; let the model refine it in the background.

    WHY THIS IS SPLIT
        gather_facts() finishes in well under a second -- it is file parsing. The
        model costs ~45 s on this box because generation runs at ~4.7 tok/s, and
        that is the floor, not a tuning problem. Blocking an inbox poller or an
        HTTP handler for 45 s per email is not acceptable, and blocking it for
        five emails is far worse.

        So the caller gets the facts brief immediately -- counts, line items,
        BOM<->CAD reconciliation, flags, and the fallback summary, which is most
        of the value -- written to email_ai_brief.json with llm.pending True. A
        daemon thread then re-runs with the model and rewrites the same file.
        Poll the file, or read it once and again later; a reader that never comes
        back still has a correct, complete brief.
    """
    import threading

    paths = [Path(p) for p in (attachment_paths or [])]
    facts = gather_facts(subject, body, paths, from_addr=from_addr)
    brief = merge_brief(facts, {"ok": False, "model": pick_model(model), "seconds": 0.0,
                                "error": "model still running"})
    brief["source"] = _source_of(facts)
    brief["llm"]["pending"] = True
    save_brief(target_dir, brief)

    def _refine():
        try:
            prompt = build_prompt_deep(facts) if deep else build_prompt(facts)
            llm = ask_model(prompt, model=model)
            final = merge_brief(facts, llm)
            final["source"] = brief["source"]
            final["llm"]["pending"] = False
            save_brief(target_dir, final)
        except Exception as e:
            # A crash here must not lose the facts brief already on disk.
            stale = load_brief(target_dir) or brief
            stale.setdefault("llm", {})["pending"] = False
            stale["llm"]["error"] = f"{type(e).__name__}: {e}"
            save_brief(target_dir, stale)

    threading.Thread(target=_refine, daemon=True, name="cms-email-ai").start()
    return brief


def save_brief(target_dir, brief: dict) -> Path:
    d = Path(target_dir)
    d.mkdir(parents=True, exist_ok=True)
    path = d / BRIEF_FILENAME
    path.write_text(json.dumps(brief, indent=2, default=str), encoding="utf-8")
    return path


def load_brief(target_dir) -> dict | None:
    path = Path(target_dir) / BRIEF_FILENAME
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def handoff_lines(brief: dict) -> dict:
    """Extra key=value lines for cms_email.txt.

    ReadHandoffFile in Module6121.bas (~1949) is a Select Case with no Case
    Else, so keys it does not know are ignored. That makes it safe to add these
    now and teach the macro to read them later, without a lock-step change.
    """
    rec = brief.get("reconciliation", {})
    counts = brief.get("counts", {})
    grades = counts.get("grades", {})
    dominant = max(grades.items(), key=lambda kv: kv[1])[0] if grades else ""
    return {
        "AiRequestType": brief.get("request_type", ""),
        "AiBaseType": brief.get("base_type", ""),
        "AiAction": brief.get("recommended_action", ""),
        "AiSummary": (brief.get("what_is_being_quoted", "") or "").replace("\n", " ")[:300],
        "AiUrgency": brief.get("delivery", {}).get("urgency", ""),
        "AiRequestedDelivery": brief.get("delivery", {}).get("requested", ""),
        "AiJobRef": brief.get("job_reference", ""),
        "AiCustomerName": brief.get("customer", {}).get("company", ""),
        "AiPlateLines": str(counts.get("plate_lines", 0)),
        "AiPlatePieces": str(counts.get("plate_pieces", 0)),
        "AiPurchasedLines": str(counts.get("purchased_lines", 0)),
        "AiFastenerLines": str(counts.get("fastener_lines", 0)),
        "AiDominantGrade": dominant,
        "AiBomOnly": ";".join(str(i.get("detail", "")) for i in rec.get("bom_only", []))[:300],
        "AiCadOnly": ";".join(str(i.get("product", "")) for i in rec.get("cad_only", []))[:300],
        "AiQtyMismatch": ";".join(str(i.get("detail", "")) for i in rec.get("qty_mismatch", []))[:300],
        "AiFlagCount": str(len(brief.get("flags", []))),
        "AiBlockers": ";".join(f["message"] for f in brief.get("flags", [])
                               if f.get("level") == "blocker")[:300],
        "AiModel": brief.get("llm", {}).get("model", ""),
        "AiOk": "1" if brief.get("llm", {}).get("ok") else "0",
    }


# ===========================================================================
# CLI
# ===========================================================================

def format_text(brief: dict) -> str:
    """Human-readable brief for the terminal and for pasting into a job note."""
    L: list[str] = []
    add = L.append
    llm = brief.get("llm", {})
    add("=" * 74)
    add("CMS QUOTE BRIEF")
    add("=" * 74)
    src = brief.get("source", {})
    add(f"From:     {src.get('from') or '(unknown)'}")
    add(f"Subject:  {src.get('subject') or '(none)'}")
    add(f"Files:    {', '.join(src.get('attachment_names') or []) or '(none)'}")
    add(f"Model:    {llm.get('model') or '(none)'}"
        + (f"  [{llm.get('seconds')}s]" if llm.get("seconds") else "")
        + ("" if llm.get("ok") else "   *** NOT AVAILABLE -- computed facts only ***"))
    if llm.get("error"):
        add(f"          {llm['error']}")
    add("")
    add(f"REQUEST TYPE:  {brief.get('request_type')}    "
        f"BASE TYPE: {brief.get('base_type')}")
    add(f"ACTION:        {brief.get('recommended_action')}")
    add(f"WHAT:          {brief.get('what_is_being_quoted')}")
    cust = brief.get("customer", {})
    who = " / ".join(v for v in (cust.get("company"), cust.get("contact"),
                                 cust.get("email"), cust.get("phone")) if v)
    add(f"CUSTOMER:      {who or '(not stated)'}")
    if brief.get("job_reference"):
        add(f"THEIR JOB #:   {brief['job_reference']}")
    d = brief.get("delivery", {})
    add(f"DELIVERY:      {d.get('urgency')}"
        + (f" -- '{d.get('requested')}'" if d.get("requested") else "")
        + ("  (quote a delivery date with the price)" if d.get("quote_delivery_with_price") else ""))

    c = brief.get("counts", {})
    add("")
    add("-- COUNTS (computed) " + "-" * 53)
    add(f"  Steel plates:    {c.get('plate_lines', 0)} lines / {c.get('plate_pieces', 0)} pieces")
    for g, n in sorted((c.get("grades") or {}).items(), key=lambda kv: -kv[1]):
        add(f"       {n:>3} x {GRADE_LABELS.get(g, g)}")
    add(f"  Purchased:       {c.get('purchased_lines', 0)} lines / {c.get('purchased_pieces', 0)} pieces")
    for v, n in sorted((c.get("vendors") or {}).items(), key=lambda kv: -kv[1]):
        add(f"       {n:>3} lines  {v}")
    add(f"  Fasteners:       {c.get('fastener_lines', 0)} lines / {c.get('fastener_pieces', 0)} pieces")
    add(f"  CAD parts:       {c.get('cad_products', 0)} distinct / {c.get('cad_instances', 0)} placed")

    items = brief.get("line_items", {})
    if items.get("plates"):
        add("")
        add("-- STEEL ON THE BOM " + "-" * 54)
        add("   Sizes below are what the CUSTOMER stated. Quote off the CAD.")
        for p in items["plates"]:
            size = (f"{p['bom_thickness']} x {p['bom_width']} x {p['bom_length']}"
                    if p.get("bom_thickness") else "size not stated")
            mark = "" if p.get("grade_priceable") else "   <-- grade unclear"
            add(f"  {str(p['detail'] or '?'):<18} qty {p['qty']:<3} {size:<28} "
                f"{p.get('grade_label') or p.get('material_raw') or '?'}{mark}")
    if items.get("purchased"):
        add("")
        add("-- PURCHASED COMPONENTS " + "-" * 50)
        for p in items["purchased"]:
            mark = "" if p.get("priceable") else "   <-- no part #"
            add(f"  {str(p['detail'] or '?'):<18} qty {p['qty']:<3} "
                f"{(p.get('vendor') or '?'):<16} {(p.get('part_number') or '?'):<16} "
                f"{p['name'][:34]}{mark}")

    rec = brief.get("reconciliation", {})
    add("")
    add("-- BOM vs CAD " + "-" * 59)
    if rec.get("compared"):
        add(f"  Matched:  {rec.get('matched', 0)} lines against {rec.get('cad_file')}")
        for i in rec.get("bom_only", []):
            extra = " [include column blank]" if i.get("marked_include") is False else ""
            note = f" -- {i['comment']}" if i.get("comment") else ""
            add(f"  BOM only: {i.get('detail')} (qty {i.get('qty')}){extra}{note}")
        for i in rec.get("cad_only", []):
            add(f"  CAD only: {i.get('product')} x{i.get('instances')}")
        for i in rec.get("qty_mismatch", []):
            add(f"  QTY:      {i.get('detail')} BOM {i.get('bom_qty')} vs CAD {i.get('cad_instances')}")
        if not (rec.get("bom_only") or rec.get("cad_only") or rec.get("qty_mismatch")):
            add("  Clean -- every BOM line has a CAD part at the same quantity.")
    else:
        add("  Not compared.")

    for title, key in (("SPECIAL INSTRUCTIONS", "special_instructions"),
                       ("OPEN QUESTIONS FOR THE CUSTOMER", "open_questions"),
                       ("RISKS", "risks")):
        vals = brief.get(key) or []
        if vals:
            add("")
            add(f"-- {title} " + "-" * max(4, 72 - len(title)))
            for v in vals:
                add(f"  - {v}")

    flags = brief.get("flags") or []
    if flags:
        add("")
        add("-- FLAGS " + "-" * 64)
        for f in flags:
            add(f"  [{f.get('level', '?').upper():<7}] {f.get('message')}")

    if brief.get("reasoning"):
        add("")
        add("-- WHY " + "-" * 66)
        for line in _wrap(brief["reasoning"], 70):
            add(f"  {line}")
    add("=" * 74)
    return "\n".join(L)


def _wrap(text: str, width: int) -> list[str]:
    words = str(text).split()
    lines, cur = [], ""
    for w in words:
        if len(cur) + len(w) + 1 > width:
            lines.append(cur)
            cur = w
        else:
            cur = f"{cur} {w}".strip()
    if cur:
        lines.append(cur)
    return lines


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Read an RFQ email + attachments with the local Ollama model.")
    ap.add_argument("files", nargs="*", help="attachment paths (BOM, CAD, PDF, ZIP)")
    ap.add_argument("--dir", help="analyze every file in this folder instead")
    ap.add_argument("--subject", default="", help="email subject line")
    ap.add_argument("--body", default="", help="email body text")
    ap.add_argument("--body-file", help="read the email body from this file")
    ap.add_argument("--from", dest="from_addr", default="", help="sender")
    ap.add_argument("--model", default="", help="Ollama model tag (default: fastest installed)")
    ap.add_argument("--no-llm", action="store_true", help="computed facts only, skip Ollama")
    ap.add_argument("--deep", action="store_true",
                    help="send the full BOM/CAD digest (much slower; pair with a big --model)")
    ap.add_argument("--timeout", type=int, default=LLM_TIMEOUT_S,
                    help=f"seconds to wait for the model (default {LLM_TIMEOUT_S})")
    ap.add_argument("--json", action="store_true", help="print the brief as JSON")
    ap.add_argument("--save", help="write the brief JSON to this folder")
    ap.add_argument("--facts", action="store_true", help="dump raw computed facts and exit")
    ap.add_argument("--prompt", action="store_true", help="print the model prompt and exit")
    args = ap.parse_args(argv)

    body = args.body
    if args.body_file:
        body = Path(args.body_file).read_text(encoding="utf-8", errors="replace")

    if args.dir:
        files = sorted(p for p in Path(args.dir).glob("*") if p.is_file())
    else:
        files = [Path(f) for f in args.files]
    if not files:
        ap.error("give at least one attachment path, or --dir")

    if args.facts or args.prompt:
        facts = gather_facts(args.subject, body, files, from_addr=args.from_addr)
        if args.prompt:
            print(build_prompt_deep(facts) if args.deep else build_prompt(facts))
        else:
            print(json.dumps(facts, indent=2, default=str))
        return 0

    brief = analyze(args.subject, body, files, from_addr=args.from_addr,
                    model=args.model, use_llm=not args.no_llm, deep=args.deep,
                    timeout=args.timeout)
    print(json.dumps(brief, indent=2, default=str) if args.json else format_text(brief))
    if args.save:
        print(f"\nSaved: {save_brief(args.save, brief)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
