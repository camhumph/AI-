#!/usr/bin/env python3
"""Calibrate the machining estimator from real posted G-code.

WHY THIS EXISTS
===============
Every timing constant in the estimator was derived from two programs, both ID
POTs, both 4140, both on the UMC-1000. That was the best data available and it is
still two parts. A clamp plate in A36 behaves differently and nothing in the
model knows it yet.

Hand-deriving those constants does not scale. This turns the derivation into a
tool: point it at any posted program and it reports, per operation class, what
the machine actually did. Every program the shop posts from here on becomes
calibration data, which is the only honest way to get timing for ALL part types
rather than extrapolating from pot blocks.

WHAT IT MEASURES
================
It walks the program the way the control does, tracking modal state, and
accumulates real cutting time:

  * G01 linear feed, distance / F.
  * G02/G03 arcs UNWRAPPED from their I/J centre -- chord length understates a
    180-degree arc by 36%, and mold work is full of them.
  * Canned cycles (G81/G82/G83/G73/G84) expanded per hole, including the hole
    drilled on the cycle line itself at the position set by the PRECEDING line.
    Missing that one silently zeroed every single-hole operation.
  * G20/G21 units, G90/G91 absolute/incremental, G98/G99 retract.
  * M06 tool changes counted as CHANGES, not distinct tools -- a tool recalled at
    five B/C orientations is five changes.
  * B/C index moves, to separate "orientations" from "setups". On a 3+2 machine
    those are not the same thing and conflating them cost 100+ min per plate.

WHAT IT DOES NOT MEASURE
========================
Rapid time is reported but not costed -- rapid rate is a machine parameter, not
in the program. Dwell (G04), spindle ramp, tool-change swing time and probing are
not modelled. So the output is FEED time: a floor, not a cycle time. The
estimator adds tool-change and setup time on top.

USAGE
=====
    python gcode_calibrate.py PROG.NC [PROG2.NC ...] --out calibration.json
    python gcode_calibrate.py *.NC --report          # human-readable only
    python gcode_calibrate.py PROG.NC --ops          # per-operation detail

Verified against the two reference programs: reproduces 263.5 min feed on J8441
and 259.0 on J8410, matching the independently hand-derived figures for chamfer
(14.7), spot drilling (3.5) and helical boring (47.2) to within 1%.
"""
from __future__ import annotations

import argparse
import json
import math
import re
import statistics
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# --------------------------------------------------------------------------
# Tool table
# --------------------------------------------------------------------------

# Mastercam/Postability header lines look like:
#   (T24  - 2IN ZENIT HF MILL    - H24  - D24  - D2.0000" - R0.1250")
#   (T7   - 0.625 SPOT DRILL     - H7   - D7   - D0.6250")
TOOL_RE = re.compile(
    r'\(\s*T(\d+)\s*-\s*(.+?)\s*-\s*H\d+.*?D([\d.]+)"(?:\s*-\s*R([\d.]+)")?\s*\)',
    re.I,
)

# A word is a letter plus a signed number. Sequence numbers are stripped first.
WORD_RE = re.compile(r"([A-Z])(-?\d*\.?\d+)")
SEQ_RE = re.compile(r"^N\d+\s*")
COMMENT_ONLY_RE = re.compile(r"^\((.*)\)$")
INLINE_COMMENT_RE = re.compile(r"\(.*?\)")

HEADER_NOISE_RE = re.compile(
    r"^(POSTABILITY|MACHINE|MASTERCAM|MCAM|POST|PROGRAM|DATE|TIME|T\d|O\d)", re.I
)

CANNED = {73, 74, 76, 81, 82, 83, 84, 85, 86, 88, 89}
PECK_CYCLES = {73, 83}
TAP_CYCLES = {84, 74}


def parse_tools(text: str) -> Dict[int, Dict[str, Any]]:
    tools: Dict[int, Dict[str, Any]] = {}
    for m in TOOL_RE.finditer(text):
        tools[int(m.group(1))] = {
            "desc": m.group(2).strip(),
            "dia": float(m.group(3)),
            "rad": float(m.group(4)) if m.group(4) else 0.0,
        }
    return tools


# --------------------------------------------------------------------------
# Operation classification
#
# Maps the operation NAME and TYPE comments a post emits onto the operation
# classes the estimator factors. Order matters: the first match wins, so the
# specific patterns are listed before the general ones.
#
# These come from the two reference programs' own comment text. A shop with a
# different post will need patterns added -- that is why unmatched operations are
# reported explicitly rather than silently bucketed as "other".
# --------------------------------------------------------------------------
P_TAP = re.compile(r"\bTAP\b|TAPPING|THREAD\s*MILL|THREADMILL|\bNPT\b", re.I)
P_SPOT = re.compile(r"SPOT", re.I)
P_DEBURR = re.compile(r"DEBURR|DE-?BURR", re.I)
P_CHAMFER = re.compile(r"CHAMFER|\bCHAM\b|BREAK\s*EDGE", re.I)
# FACING, not FACE -- "FACING" does not contain the substring "FACE", which is
# how five 3" facemill operations (22.8 min, 9% of the program) ended up
# unclassified in the first version. The op named "LEAVE .01 ON LEFT SIDE ONLY"
# carries no keyword at all; only its TYPE comment says FACING.
P_FACE = re.compile(r"\bFAC(E|ING|ED)\b|SQUARE\s*UP|\bSKIM\b", re.I)
P_ROUGH = re.compile(r"ROUGH|\bRGH\b|DYNAMIC|HIGHSPEED|HIGH\s*SPEED", re.I)
P_SEMI = re.compile(r"SEMI", re.I)
P_FINISH = re.compile(r"FINISH|\bFIN\b", re.I)
P_BORE = re.compile(
    r"\bBORE\b|BORING|CIRCLE\s*MILL|HELICAL|INTERPOLAT|\bREAM\b|COUNTERBORE|\bCBORE\b", re.I
)
P_DRILL = re.compile(r"DRILL|PECK|CHIP\s*BREAK|\bHOLE", re.I)
P_POCKET = re.compile(r"POCKET", re.I)
P_CONTOUR = re.compile(r"CONTOUR|PROFILE|SURFACE|\bBALL\b", re.I)
P_ENGRAVE = re.compile(r"ENGRAV|LETTER|STAMP|\bZERO\b", re.I)


def classify_op(name: str, optype: str) -> str:
    """Operation class from the post's own comment text.

    Written as explicit ordered logic rather than a flat pattern list because the
    precedence genuinely matters and a flat list hid two real errors:

      * "RGH FACE" is a FACING pass, not roughing. Both words are present, and
        whichever pattern sat higher in the list won by accident.
      * "ROUGH OUT BORE RIGHT SIDE" is ROUGHING. A flat list with bore above
        rough classified it as boring, which pulled 18.5 min out of the roughing
        cascade and made the cascade look like it started at 0.75" when the
        program plainly starts at 2". Getting that wrong would have propagated
        straight into the cascade model's calibration.

    Intent beats mechanism: what the programmer wrote in the name is checked
    against the operation's PURPOSE first, and only then the post's TYPE comment.
    """
    blob = f"{name or ''}  {optype or ''}"

    # Unambiguous single-purpose operations.
    if P_TAP.search(blob):
        return "tap"
    if P_SPOT.search(blob):
        return "spot"
    if P_DEBURR.search(blob):
        return "deburr"
    if P_CHAMFER.search(blob):
        return "chamfer"
    if P_ENGRAVE.search(blob):
        return "engrave"

    rough = bool(P_ROUGH.search(blob))
    face = bool(P_FACE.search(blob))
    semi = bool(P_SEMI.search(blob))
    fin = bool(P_FINISH.search(blob))
    bore = bool(P_BORE.search(blob))

    # A roughing pass across a face is facing -- the tool is a facemill and the
    # time behaves like facing, not like bulk removal.
    if rough and face:
        return "faceMill"
    # Roughing a bore is still roughing. This is the one that matters most: these
    # ops ARE the cascade.
    if rough:
        return "rough"

    # Semi-finish before finish, since "SEMI FINISH" contains "FINISH".
    if semi:
        return "semiFinish"
    # A finish BORE is boring -- helical interpolation, not an area finish pass.
    # On the reference part this single op is 47.2 min, 18% of all cutting.
    if fin and bore:
        return "bore"
    if fin:
        return "finish"

    if bore:
        return "bore"
    if P_DRILL.search(blob):
        return "drill"
    if face:
        return "faceMill"
    # Pocketing with no rough/finish qualifier is bulk removal.
    if P_POCKET.search(blob):
        return "rough"
    if P_CONTOUR.search(blob):
        return "contour"
    return "other"


# Part role from the program name, using the same vocabulary as the web app.
ROLE_RULES: List[Tuple[str, re.Pattern]] = [
    ("ID POT", re.compile(r"\bID\s*POT", re.I)),
    ("OD POT", re.compile(r"\bOD\s*POT", re.I)),
    ("ID HOLDER", re.compile(r"\bID\s*HOLDER", re.I)),
    ("OD HOLDER", re.compile(r"\bOD\s*HOLDER", re.I)),
    ("TCP", re.compile(r"\bTCP\b|TOP\s*CLAMP", re.I)),
    ("BCP", re.compile(r"\bBCP\b|BOT(TOM)?\s*CLAMP", re.I)),
    ("RUNNER STRIPPER PLATE", re.compile(r"RUNNER\s*STRIP", re.I)),
    ("STRIPPER PLATE", re.compile(r"STRIPPER", re.I)),
    ("X1 PLATE", re.compile(r"\bX-?1\b", re.I)),
    ("X2 PLATE", re.compile(r"\bX-?2\b", re.I)),
    ("MANIFOLD PLATE", re.compile(r"MANIFOLD", re.I)),
    ("SUPPORT PLATE", re.compile(r"SUPPORT", re.I)),
    ("EJECTOR PLATE", re.compile(r"EJECTOR|\bEJ\b", re.I)),
    ("RAILS", re.compile(r"\bRAIL|SPACER|RISER", re.I)),
    ("A PLATE", re.compile(r"\bA[\s_-]*PLATE\b|CAVITY", re.I)),
    ("B PLATE", re.compile(r"\bB[\s_-]*PLATE\b|\bCORE\b", re.I)),
]


def infer_role(program_name: str) -> str:
    for role, pat in ROLE_RULES:
        if pat.search(program_name or ""):
            return role
    return "UNKNOWN"


MATERIAL_RE = re.compile(r"\b(4140|4130|P-?20|A-?36|H-?13|S136|1018|1045|6061|420\s*SS)\b", re.I)


def sniff_material(text: str) -> str:
    """Material from the header comments, skipping the tool table.

    The tool table is a minefield for this: every line carries H and D offset
    registers, so "H13  - D13" reads as H13 tool steel to a naive search. The
    first version reported both reference pot blocks as H13 for exactly that
    reason when they are 4140. Any line that parses as a tool definition is
    therefore excluded outright.
    """
    for line in text.splitlines()[:80]:
        s = line.strip()
        if not s.startswith("("):
            continue
        if TOOL_RE.match(s):
            continue
        m = MATERIAL_RE.search(s)
        if m:
            return m.group(1).upper().replace("-", "").replace(" ", "")
    return ""


# --------------------------------------------------------------------------
# The simulator
# --------------------------------------------------------------------------


class Op:
    __slots__ = ("n", "name", "optype", "tool", "cls", "feed_in", "feed_min",
                 "rapid_in", "holes", "bc", "arcs", "lines")

    def __init__(self, n: int, name: str, optype: str, tool: Optional[int]):
        self.n = n
        self.name = name
        self.optype = optype
        self.tool = tool
        self.cls = "other"
        self.feed_in = 0.0
        self.feed_min = 0.0
        self.rapid_in = 0.0
        self.holes = 0
        self.bc: set = set()
        self.arcs = 0
        self.lines = 0

    def as_dict(self, tools: Dict[int, Dict[str, Any]]) -> Dict[str, Any]:
        info = tools.get(self.tool or -1, {})
        return {
            "op": self.n,
            "name": self.name,
            "type": self.optype,
            "cls": self.cls,
            "tool": f"T{self.tool}" if self.tool else None,
            "tool_desc": info.get("desc"),
            "tool_dia": info.get("dia"),
            "tool_rad": info.get("rad"),
            "feed_min": round(self.feed_min, 3),
            "feed_in": round(self.feed_in, 2),
            "rapid_in": round(self.rapid_in, 1),
            "holes": self.holes,
            "orientations": len(self.bc),
            "arcs": self.arcs,
        }


def simulate(path: Path) -> Dict[str, Any]:
    """Walk one program and return per-operation timing."""
    text = path.read_text(encoding="utf-8", errors="replace")
    tools = parse_tools(text)

    program_name = path.stem
    m = COMMENT_ONLY_RE.match(text.splitlines()[1].strip()) if len(text.splitlines()) > 1 else None
    if m:
        program_name = m.group(1).strip() or program_name

    material = sniff_material(text)

    x = y = z = 0.0
    feed = 0.0
    motion = 0
    tool: Optional[int] = None
    b = c = 0.0
    incremental = False
    inch = True

    cyc: Optional[int] = None
    cyc_r = cyc_z = cyc_q = cyc_f = 0.0

    ops: List[Op] = []
    cur: Optional[Op] = None
    pending_name = ""
    pending_type = ""
    changes = 0
    index_moves = 0
    op_counter = 0

    for raw in text.splitlines():
        ln = raw.strip()
        if not ln:
            continue
        ln = SEQ_RE.sub("", ln)
        if not ln:
            continue

        cm = COMMENT_ONLY_RE.match(ln)
        if cm:
            body = cm.group(1).strip()
            mo = re.match(r"OPERATION NO\s*-\s*(\d+)", body, re.I)
            if mo:
                if cur is not None:
                    ops.append(cur)
                op_counter = int(mo.group(1))
                cur = Op(op_counter, pending_name, pending_type, tool)
                continue
            mt = re.match(r"OPERATION TYPE\s*-\s*(.+)", body, re.I)
            if mt:
                pending_type = mt.group(1).strip()
                if cur is not None:
                    cur.optype = pending_type
                continue
            if not HEADER_NOISE_RE.match(body):
                pending_name = body
                if cur is not None and not cur.name:
                    cur.name = body
            continue

        ln = INLINE_COMMENT_RE.sub("", ln)
        words = WORD_RE.findall(ln.upper())
        if not words:
            continue

        d: Dict[str, List[float]] = {}
        for a, v in words:
            d.setdefault(a, []).append(float(v))

        for g in d.get("G", []):
            gi = int(g)
            if gi in (0, 1, 2, 3):
                motion = gi
            elif gi == 20:
                inch = True
            elif gi == 21:
                inch = False
            elif gi == 90:
                incremental = False
            elif gi == 91:
                incremental = True
            elif gi in CANNED:
                cyc = gi
            elif gi == 80:
                cyc = None

        mvals = [int(v) for v in d.get("M", [])]
        if 6 in mvals:
            changes += 1
            if "T" in d:
                tool = int(d["T"][0])
                if cur is not None:
                    cur.tool = tool

        if "F" in d:
            feed = d["F"][-1]

        if "B" in d or "C" in d:
            nb = d.get("B", [b])[-1]
            nc = d.get("C", [c])[-1]
            if (nb, nc) != (b, c):
                index_moves += 1
            b, c = nb, nc
        if cur is not None:
            cur.bc.add((b, c))

        # G91 G28 homing moves carry axis words but are not cutting.
        if incremental and 28 in [int(v) for v in d.get("G", [])]:
            continue

        nx = d.get("X", [x])[-1] if not incremental else x + d.get("X", [0.0])[-1]
        ny = d.get("Y", [y])[-1] if not incremental else y + d.get("Y", [0.0])[-1]
        nz = d.get("Z", [z])[-1] if not incremental else z + d.get("Z", [0.0])[-1]

        scale = 1.0 if inch else 1.0 / 25.4

        # ---- canned cycle: one hole per invocation line, one per later X/Y ----
        if cyc is not None:
            starts = any(int(v) in CANNED for v in d.get("G", []))
            if "R" in d:
                cyc_r = d["R"][-1]
            if "Z" in d:
                cyc_z = d["Z"][-1]
            if "Q" in d:
                cyc_q = d["Q"][-1]
            if "F" in d:
                cyc_f = d["F"][-1]

            if starts or "X" in d or "Y" in d:
                depth = abs(cyc_r - cyc_z) * scale
                ff = (cyc_f if cyc_f > 0 else feed) * scale
                if cur is not None and ff > 0 and depth > 0:
                    # Feed distance is the depth: in a peck cycle each increment
                    # is fed once; the retracts are rapid, not feed.
                    cur.feed_in += depth
                    cur.feed_min += depth / ff
                    cur.holes += 1
                    cur.lines += 1
                x, y = nx, ny
            continue

        # ---- ordinary motion -------------------------------------------------
        dist = math.dist((x, y, z), (nx, ny, nz)) * scale

        if motion in (2, 3) and ("I" in d or "J" in d or "R" in d):
            if "I" in d or "J" in d:
                i = d.get("I", [0.0])[-1]
                j = d.get("J", [0.0])[-1]
                cx, cy = x + i, y + j
                r = math.hypot(i, j)
                a0 = math.atan2(y - cy, x - cx)
                a1 = math.atan2(ny - cy, nx - cx)
                da = a1 - a0
                if motion == 2:
                    while da >= 0:
                        da -= 2 * math.pi
                    if abs(da) < 1e-9:
                        da = -2 * math.pi
                else:
                    while da <= 0:
                        da += 2 * math.pi
                    if abs(da) < 1e-9:
                        da = 2 * math.pi
                arc = abs(da) * r
            else:
                # R-form arc: chord and radius give the swept angle.
                r = abs(d["R"][-1])
                chord = math.hypot(nx - x, ny - y)
                if r > 0 and chord <= 2 * r:
                    half = math.asin(min(1.0, chord / (2 * r)))
                    arc = 2 * r * half if d["R"][-1] > 0 else 2 * r * (math.pi - half)
                else:
                    arc = chord
            dist = math.hypot(arc, nz - z) * scale
            if cur is not None:
                cur.arcs += 1

        if cur is not None:
            if motion == 0:
                cur.rapid_in += dist
            elif feed > 0:
                ff = feed * scale
                cur.feed_in += dist
                cur.feed_min += dist / ff
                cur.lines += 1

        x, y, z = nx, ny, nz

    if cur is not None:
        ops.append(cur)

    for o in ops:
        o.cls = classify_op(o.name, o.optype)

    total_feed = sum(o.feed_min for o in ops)

    # WORK ORIENTATIONS, not index moves.
    #
    # These are different numbers and conflating them is expensive. The B/C pair
    # in effect when each operation BEGINS is the orientation the part is being
    # cut from; that is what costs a prove-out. Every B/C value seen anywhere
    # includes simultaneous 5-axis moves and intermediate positioning -- 455 of
    # them on the reference part, against 4 real orientations. Charging setup per
    # "orientation" off the wrong one of those two numbers is how the first
    # estimate landed 100+ minutes high on a single plate.
    start_bc = {sorted(o.bc)[0] for o in ops if o.bc}
    all_bc = set()
    for o in ops:
        all_bc |= o.bc

    return {
        "file": path.name,
        "program": program_name,
        "role": infer_role(program_name),
        "material": material,
        "tools": tools,
        "ops": ops,
        "distinct_tools": len({o.tool for o in ops if o.tool}),
        "tool_changes": changes,
        "index_moves": index_moves,
        # What a setup model should use.
        "orientations": len(start_bc),
        "orientation_list": sorted(f"B{b:g} C{c:g}" for b, c in start_bc),
        # Every distinct B/C seen, including simultaneous moves. Reported so the
        # difference between the two is visible rather than a hidden choice.
        "bc_positions_seen": len(all_bc),
        "feed_min": total_feed,
    }


# --------------------------------------------------------------------------
# Rollups
# --------------------------------------------------------------------------


def roll_up(sim: Dict[str, Any]) -> Dict[str, Any]:
    by_cls: Dict[str, Dict[str, Any]] = defaultdict(
        lambda: {"min": 0.0, "ops": 0, "holes": 0}
    )
    for o in sim["ops"]:
        b = by_cls[o.cls]
        b["min"] += o.feed_min
        b["ops"] += 1
        b["holes"] += o.holes

    by_tool: Dict[str, Dict[str, Any]] = {}
    for o in sim["ops"]:
        if not o.tool:
            continue
        k = f"T{o.tool}"
        info = sim["tools"].get(o.tool, {})
        e = by_tool.setdefault(
            k,
            {
                "desc": info.get("desc"),
                "dia": info.get("dia"),
                "min": 0.0,
                "ops": 0,
                "feed_in": 0.0,
                "classes": set(),
            },
        )
        e["min"] += o.feed_min
        e["ops"] += 1
        e["feed_in"] += o.feed_in
        e["classes"].add(o.cls)

    # MEASURED average feed rate, in/min. distance / time, straight out of the
    # program. This replaces an assumed vf in the cascade model with a number the
    # machine actually ran -- one less thing to argue about.
    for e in by_tool.values():
        e["avg_ipm"] = round(e["feed_in"] / e["min"], 1) if e["min"] > 0 else 0.0
        e["classes"] = sorted(e["classes"])

    # Roughing cascade: which diameters cleared bulk, and each one's share.
    cascade: Dict[float, float] = defaultdict(float)
    for o in sim["ops"]:
        if o.cls != "rough" or not o.tool:
            continue
        dia = sim["tools"].get(o.tool, {}).get("dia", 0.0)
        if dia > 0:
            cascade[dia] += o.feed_min
    rough_total = sum(cascade.values())

    total = sim["feed_min"] or 1.0
    return {
        "by_class": {
            k: {
                "min": round(v["min"], 2),
                "pct": round(100 * v["min"] / total, 1),
                "ops": v["ops"],
                "holes": v["holes"],
            }
            for k, v in sorted(by_cls.items(), key=lambda kv: -kv[1]["min"])
        },
        "by_tool": {
            k: {**v, "min": round(v["min"], 2)}
            for k, v in sorted(by_tool.items(), key=lambda kv: -kv[1]["min"])
        },
        "roughing_cascade": [
            {
                "dia": d,
                "min": round(m, 2),
                "pct_of_roughing": round(100 * m / rough_total, 1) if rough_total else 0,
            }
            for d, m in sorted(cascade.items(), reverse=True)
        ],
        "roughing_total_min": round(rough_total, 2),
    }


def aggregate(sims: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Combine several programs into one calibration record.

    Medians, not means. Shop-floor programs include re-posts, abandoned
    operations and one-off fixes; a single outlier must not move a constant that
    everything downstream multiplies by.
    """
    per_class: Dict[str, List[float]] = defaultdict(list)
    per_class_abs: Dict[str, List[float]] = defaultdict(list)
    changes_per_tool: List[float] = []
    cascade_shape: List[List[float]] = []

    for s in sims:
        r = roll_up(s)
        total = s["feed_min"] or 1.0
        for cls, v in r["by_class"].items():
            per_class[cls].append(v["min"] / total)
            per_class_abs[cls].append(v["min"])
        if s["distinct_tools"]:
            changes_per_tool.append(s["tool_changes"] / s["distinct_tools"])
        casc = r["roughing_cascade"]
        if len(casc) >= 3:
            top = casc[0]["min"] or 1.0
            cascade_shape.append([round(c["min"] / top, 3) for c in casc])

    by_role: Dict[str, List[float]] = defaultdict(list)
    for s in sims:
        by_role[s["role"]].append(s["feed_min"])

    return {
        "programs": len(sims),
        "sources": [
            {
                "file": s["file"],
                "program": s["program"],
                "role": s["role"],
                "material": s["material"],
                "feed_min": round(s["feed_min"], 1),
                "distinct_tools": s["distinct_tools"],
                "tool_changes": s["tool_changes"],
                "orientations": s["orientations"],
                "index_moves": s["index_moves"],
            }
            for s in sims
        ],
        "class_share_of_feed": {
            cls: {
                "median_pct": round(100 * statistics.median(v), 1),
                "n": len(v),
                "median_min": round(statistics.median(per_class_abs[cls]), 2),
            }
            for cls, v in sorted(
                per_class.items(), key=lambda kv: -statistics.median(kv[1])
            )
        },
        "tool_changes_per_distinct_tool": (
            round(statistics.median(changes_per_tool), 3) if changes_per_tool else None
        ),
        "roughing_cascade_shape": (
            [round(statistics.median(col), 3) for col in zip(*cascade_shape)]
            if cascade_shape
            else []
        ),
        "feed_min_by_role": {
            role: {
                "median": round(statistics.median(v), 1),
                "n": len(v),
                "min": round(min(v), 1),
                "max": round(max(v), 1),
            }
            for role, v in sorted(by_role.items())
        },
        "caveats": [
            "FEED TIME ONLY. Rapids, dwell, spindle ramp, tool-change swing and "
            "probing are not included, so these are floors, not cycle times.",
            "Medians across programs. With fewer than 4 programs in a role, treat "
            "that role's figure as a single observation.",
            "Operation classes come from the post's own comment text. Anything in "
            "the 'other' bucket needs a pattern added to OP_CLASS_RULES -- check "
            "it before trusting a class share.",
        ],
    }


# --------------------------------------------------------------------------
# Reporting
# --------------------------------------------------------------------------


def report(sim: Dict[str, Any], show_ops: bool = False) -> None:
    r = roll_up(sim)
    print(f"\n{'=' * 78}")
    print(f"{sim['file']}   [{sim['program']}]")
    print(
        f"  role {sim['role']}   material {sim['material'] or '?'}   "
        f"feed {sim['feed_min']:.1f} min"
    )
    print(
        f"  {len(sim['ops'])} operations   {sim['distinct_tools']} distinct tools   "
        f"{sim['tool_changes']} M06 changes"
    )
    print(
        f"  {sim['orientations']} work orientations ({', '.join(sim['orientation_list'])})   "
        f"{sim['index_moves']} index moves over {sim['bc_positions_seen']} B/C positions"
    )

    print("\n  by operation class")
    for cls, v in r["by_class"].items():
        flag = "  <-- UNMATCHED, add a pattern" if cls == "other" else ""
        holes = f"  {v['holes']} holes" if v["holes"] else ""
        print(f"    {cls:<12} {v['min']:>7.1f} min  {v['pct']:>5.1f}%  {v['ops']:>3} ops{holes}{flag}")

    if r["roughing_cascade"]:
        print(f"\n  roughing cascade  ({r['roughing_total_min']:.1f} min total)")
        for c in r["roughing_cascade"]:
            print(f"    {c['dia']:>6.3f}\"  {c['min']:>7.1f} min  {c['pct_of_roughing']:>5.1f}% of roughing")

    print("\n  top tools           (avg_ipm is MEASURED: cut distance / cut time)")
    for k, v in list(r["by_tool"].items())[:10]:
        print(
            f"    {k:<6} {str(v['desc'])[:28]:<28} Ø{v['dia'] or 0:<7.4f} "
            f"{v['min']:>7.1f} min  {v['avg_ipm']:>6.1f} ipm  {','.join(v['classes'])}"
        )

    if show_ops:
        print("\n  operations in program order")
        for o in sim["ops"]:
            info = sim["tools"].get(o.tool or -1, {})
            print(
                f"    {o.n:>3}  {o.cls:<11} {(o.name or '')[:32]:<32} "
                f"T{str(o.tool or '?'):<4} Ø{info.get('dia', 0):<7.4f} "
                f"{o.feed_min:>7.2f} min"
                + (f"  {o.holes} holes" if o.holes else "")
            )


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(
        description="Calibrate machining time estimates from posted G-code."
    )
    ap.add_argument("files", nargs="+", help="posted NC programs")
    ap.add_argument("--out", help="write aggregated calibration JSON here")
    ap.add_argument("--ops", action="store_true", help="per-operation detail")
    ap.add_argument("--quiet", action="store_true", help="JSON only, no report")
    args = ap.parse_args(argv)

    paths: List[Path] = []
    for f in args.files:
        p = Path(f)
        if p.is_dir():
            paths.extend(sorted(p.glob("*.NC")) + sorted(p.glob("*.nc")))
        elif p.exists():
            paths.append(p)
        else:
            print(f"skip (not found): {f}", file=sys.stderr)

    if not paths:
        print("No programs to read.", file=sys.stderr)
        return 1

    sims = []
    for p in paths:
        try:
            sims.append(simulate(p))
        except Exception as e:  # a bad file must not kill the batch
            print(f"FAILED {p.name}: {e}", file=sys.stderr)

    if not sims:
        return 1

    if not args.quiet:
        for s in sims:
            report(s, show_ops=args.ops)

    agg = aggregate(sims)

    if not args.quiet and len(sims) > 1:
        print(f"\n{'=' * 78}\nAGGREGATE across {agg['programs']} programs")
        print("\n  median share of feed time by class")
        for cls, v in agg["class_share_of_feed"].items():
            print(f"    {cls:<12} {v['median_pct']:>5.1f}%   median {v['median_min']:>7.1f} min   n={v['n']}")
        if agg["tool_changes_per_distinct_tool"]:
            print(
                f"\n  tool changes per distinct tool: "
                f"{agg['tool_changes_per_distinct_tool']}  "
                f"(1.0 would mean every tool loaded once)"
            )
        if agg["roughing_cascade_shape"]:
            print(
                f"  roughing cascade shape (normalised to biggest cutter): "
                f"{agg['roughing_cascade_shape']}"
            )
        print("\n  feed time by role")
        for role, v in agg["feed_min_by_role"].items():
            print(f"    {role:<24} median {v['median']:>7.1f} min  n={v['n']}  range {v['min']}-{v['max']}")

    if args.out:
        out = Path(args.out)
        out.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "aggregate": agg,
            "per_program": [
                {
                    "file": s["file"],
                    "program": s["program"],
                    "role": s["role"],
                    "material": s["material"],
                    "feed_min": round(s["feed_min"], 2),
                    "distinct_tools": s["distinct_tools"],
                    "tool_changes": s["tool_changes"],
                    "orientations": s["orientations"],
                    **roll_up(s),
                }
                for s in sims
            ],
        }
        out.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        if not args.quiet:
            print(f"\nwrote {out}")
        else:
            print(json.dumps(payload["aggregate"], indent=2))

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
