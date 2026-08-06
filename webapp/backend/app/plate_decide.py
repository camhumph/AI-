"""Decide a part's role from STL mesh geometry + size + stack position.

WHY THIS EXISTS
---------------
The classifier decides from a bounding box and a footprint ratio, and two of the
columns it leans on are not reliable:

  * Mass_or_Vol is mass only if the CAD has a material. C17879 has none, so
    every part -- insulation and 4140 alike -- reads ~0.036 lb/in^3, which is
    SolidWorks' 1 g/cc default. Mass cannot separate steel from insulation there.
  * SolidFillPct read 0.000 for C17879 idx 8 and 100.000 for its ejector plates.
    The 0.000 is wrong on its face for a real solid.

A triangle mesh does not care about material. Volume from the divergence theorem
is exact for a closed mesh, so `mesh_volume / bbox_volume` is a fill ratio that
holds up where SolidFillPct does not.

But the thing that actually settles C17879 is POSITION, and nothing was using it.
The macro's ejector-plate test takes (baseW, baseL, baseFoot) -- footprint only,
no Z. So it could not ask either question that matters:

    idx 8   0.250 x 23.500 x 31.250 at z=+10.752   <- ABOVE the top clamp plate
    idx 11  0.625 x 15.625 x 16.313 at z=-5.252    <- inside the rail gap
    idx  6  1.125 x 15.625 x 16.313 at z=-6.127    <- inside the rail gap

An ejector plate cannot sit outside the stationary clamp plate, and cannot be
full-footprint, because it has to travel inside the rail gap. A full-footprint
quarter-inch sheet outside the outermost clamp plate is insulation. Both are
one comparison away once the stack is known.

ORDER OF EVIDENCE
-----------------
    1. BOM detail-number match      (bom_roles -- a lookup, not a guess)
    2. CAD component name token     (shop names like A-PLATE, EJ-RET-PLATE)
    3. Stack position + footprint   (this module)
    4. Nothing -- say so, do not invent a name

Sizes still come from the CAD, never the BOM -- see bom_roles.HINT_FIELDS.
"""
from __future__ import annotations

import csv
import re
import struct
from pathlib import Path

from . import bom_roles, job_housekeeping

# A plate whose footprint is within this fraction of the base is "full".
FULL_FOOTPRINT_MIN = 0.90
# Quarter-inch sheet stock, the thickness insulation ships in.
INSULATOR_MAX_T = 0.375
# An ejector plate covers a real share of the base and is roughly square; a rail
# is a long bar and a pillar is tiny. See the note in decide().
EJECTOR_MIN_AREA_FRAC = 0.15
EJECTOR_MAX_ASPECT = 4.0


# ---------------------------------------------------------------- STL reading

def read_stl(path: Path) -> dict | None:
    """Mesh volume and bounding box from a binary or ASCII STL.

    Volume is the divergence-theorem sum over triangles: for a closed mesh,
    sum(dot(v0, cross(v1, v2))) / 6. Sign depends on winding, so the absolute
    value is taken.
    """
    try:
        data = path.read_bytes()
    except Exception:
        return None
    tris = []
    # Binary: 80-byte header, uint32 count, then 50 bytes per facet.
    if len(data) >= 84:
        n = struct.unpack_from("<I", data, 80)[0]
        if n and len(data) >= 84 + n * 50:
            for i in range(n):
                off = 84 + i * 50 + 12          # skip the normal
                v = struct.unpack_from("<9f", data, off)
                tris.append((v[0:3], v[3:6], v[6:9]))
    if not tris:
        # ASCII fallback.
        try:
            text = data.decode("utf-8", "replace")
        except Exception:
            return None
        verts = [tuple(float(x) for x in m.groups())
                 for m in re.finditer(r"vertex\s+([-\d.eE+]+)\s+([-\d.eE+]+)\s+([-\d.eE+]+)", text)]
        for i in range(0, len(verts) - 2, 3):
            tris.append((verts[i], verts[i + 1], verts[i + 2]))
    if not tris:
        return None

    vol6 = 0.0
    lo = [float("inf")] * 3
    hi = [float("-inf")] * 3
    for a, b, c in tris:
        cx = b[1] * c[2] - b[2] * c[1]
        cy = b[2] * c[0] - b[0] * c[2]
        cz = b[0] * c[1] - b[1] * c[0]
        vol6 += a[0] * cx + a[1] * cy + a[2] * cz
        for p in (a, b, c):
            for k in range(3):
                if p[k] < lo[k]:
                    lo[k] = p[k]
                if p[k] > hi[k]:
                    hi[k] = p[k]
    dims = sorted((hi[k] - lo[k] for k in range(3)), reverse=True)
    vol = abs(vol6) / 6.0
    # STLs come out in millimetres on some jobs despite FORCE_INCH_STL_EXPORT --
    # C17879's read 123.825 / 603.250 / 800.100 for a 4.875 x 23.750 x 31.500
    # plate. A mold plate is never 800 inches, so a longest side over 200 means
    # mm and everything scales by 25.4 (volume by 25.4^3).
    if dims[0] > 200.0:
        dims = [d / 25.4 for d in dims]
        vol /= 25.4 ** 3
    bbox = dims[0] * dims[1] * dims[2]
    return {
        "file": path.name,
        "triangles": len(tris),
        "mesh_volume": vol,
        "bbox_volume": bbox,
        "fill_pct": (100.0 * vol / bbox) if bbox > 0 else 0.0,
        "length": dims[0], "width": dims[1], "thickness": dims[2],
    }


def load_job_stls(job_dir) -> list[dict]:
    """Every per-plate STL in the job's stl\\ folder, measured."""
    d = Path(job_dir) / "stl"
    if not d.is_dir():
        return []
    out = []
    for p in sorted(d.glob("*")):
        if p.suffix.lower() != ".stl" or not p.is_file():
            continue
        m = read_stl(p)
        if m:
            out.append(m)
    return out


# ------------------------------------------------------------------ CAD parts

def _f(v, default=0.0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def load_parts(job_dir) -> list[dict]:
    csv_path = job_housekeeping.find_job_file(job_dir, "XT_Export_CAD_Dimensions.csv")
    if csv_path is None:
        return []
    parts = []
    with csv_path.open(encoding="utf-8-sig", errors="replace", newline="") as f:
        for r in csv.DictReader(f):
            idx = (r.get("Index") or "").strip()
            if not idx:
                continue
            t, w, l = _f(r.get("Thickness")), _f(r.get("Width")), _f(r.get("Length"))
            parts.append({
                "index": idx,
                "component": (r.get("Component") or "").strip(),
                "thickness": t, "width": w, "length": l,
                "face_area": w * l,
                "bbox_volume": _f(r.get("BBoxVolume_cuin")),
                "mass_or_vol": _f(r.get("Mass_or_Vol")),
                "fill_pct_cad": _f(r.get("SolidFillPct")),
                "z": _f(r.get("CenterZ")),
                "nthru": int(_f(r.get("NThruHoles"))),
                "ncbore": int(_f(r.get("NCbore"))),
                "npockets": int(_f(r.get("NPockets"))),
            })
    return parts


def attach_mesh(parts: list[dict], meshes: list[dict], tol: float = 0.02) -> int:
    """Match each measured STL to a CAD part by bounding box. Returns match count.

    Matching on the box rather than the filename on purpose: the per-plate STLs
    are named from the role the macro assigned, so on a job that named a plate
    wrong the filename is wrong too. The geometry inside it is still correct.
    """
    used = set()
    n = 0
    for p in parts:
        want = sorted((p["thickness"], p["width"], p["length"]), reverse=True)
        for i, m in enumerate(meshes):
            if i in used:
                continue
            got = sorted((m["thickness"], m["width"], m["length"]), reverse=True)
            if all(abs(a - b) <= max(tol, 0.01 * max(a, 1.0)) for a, b in zip(want, got)):
                p["mesh"] = m
                p["fill_pct_mesh"] = m["fill_pct"]
                used.add(i)
                n += 1
                break
    return n


# -------------------------------------------------------------------- the stack

def build_stack(parts: list[dict]) -> dict:
    """Work out the base footprint, the clamp-plate span and the rail gap."""
    if not parts:
        return {}
    base_area = max(p["face_area"] for p in parts)
    full = [p for p in parts
            if p["face_area"] >= FULL_FOOTPRINT_MIN * base_area and p["thickness"] > 0]
    full.sort(key=lambda p: -p["z"])

    stack = {
        "base_area": base_area,
        "full_indices": [p["index"] for p in full],
        "n_full": len(full),
        "top_z": None, "bottom_z": None,
        "rail_gap": None,
    }
    if not full:
        return stack
    # Outermost full plates that are plausibly clamp plates (not the thin sheet).
    solid_full = [p for p in full if p["thickness"] > INSULATOR_MAX_T]
    if solid_full:
        stack["top_z"] = solid_full[0]["z"] + solid_full[0]["thickness"] / 2.0
        stack["bottom_z"] = solid_full[-1]["z"] - solid_full[-1]["thickness"] / 2.0
        stack["top_plate"] = solid_full[0]["index"]
        stack["bottom_plate"] = solid_full[-1]["index"]

    # The rail gap: the largest vertical space between two adjacent full plates.
    # On a standard base that space is exactly where the ejector plates travel.
    best = None
    for a, b in zip(solid_full, solid_full[1:]):
        a_bot = a["z"] - a["thickness"] / 2.0
        b_top = b["z"] + b["thickness"] / 2.0
        gap = a_bot - b_top
        if gap > 0.25 and (best is None or gap > best[0]):
            best = (gap, b_top, a_bot, a["index"], b["index"])
    if best:
        stack["rail_gap"] = {"height": best[0], "low_z": best[1], "high_z": best[2],
                             "above": best[3], "below": best[4]}
    return stack


def decide(job_dir) -> dict:
    """Role per CAD part from BOM > CAD name > stack position. Never invents."""
    job_dir = Path(job_dir)
    parts = load_parts(job_dir)
    if not parts:
        return {"error": "No XT_Export_CAD_Dimensions.csv for this job.", "decisions": []}

    meshes = load_job_stls(job_dir)
    matched = attach_mesh(parts, meshes)
    stack = build_stack(parts)
    hints = bom_roles.load_hints(job_dir) or bom_roles.build_hints(job_dir).get("hints", [])
    if isinstance(hints, list):
        hints = {h["cad_index"]: h for h in hints if h.get("cad_index")}

    gap = stack.get("rail_gap")
    top_z, bottom_z = stack.get("top_z"), stack.get("bottom_z")
    base_area = stack.get("base_area", 0)

    decisions = []
    for p in parts:
        role, conf, why = "", "", ""
        is_full = base_area and p["face_area"] >= FULL_FOOTPRINT_MIN * base_area

        # 1. BOM lookup
        h = hints.get(p["index"])
        if h and h.get("role"):
            role, conf = h["role"], "HIGH"
            why = f"BOM detail {h.get('detail')} names it '{h.get('bom_name','')[:40]}'"

        # 2. CAD component name token
        if not role:
            named, phrase = bom_roles.role_for_description(
                re.sub(r"[-_]+", " ", p["component"].split("/")[-1]))
            if named:
                role, conf = named, "HIGH"
                why = f"CAD name contains '{phrase}'"

        # 3. Stack position
        if not role:
            # Insulation: full-footprint sheet stock OUTSIDE the clamp plates.
            if (is_full and p["thickness"] <= INSULATOR_MAX_T
                    and top_z is not None
                    and (p["z"] > top_z or p["z"] < bottom_z)):
                role, conf = "insulator_sheet", "HIGH"
                side = "above the top" if p["z"] > top_z else "below the bottom"
                why = (f"full footprint, {p['thickness']:.3f} thick, and sits {side} "
                       f"clamp plate (z={p['z']:.3f}) -- sheet insulation, not steel")
            # Ejector plates: PLATE-SHAPED, sub-footprint, INSIDE the rail gap.
            #
            # Position alone is not enough. On C17879 the rail gap also contains
            # the rails themselves (1.875 x 3.000 x 31.500), eight support
            # pillars (0.938 x 0.938 x 5.375) and four dowel pins (0.500 x 0.500
            # x 1.250) -- all correctly inside the gap, none of them plates. Two
            # more tests separate them:
            #   area  -- an ejector plate covers a real share of the base
            #            (idx 6/7 are 255 in^2, 34% of the 748 in^2 base; the
            #            pillars are 5 in^2, under 1%)
            #   shape -- an ejector plate is roughly square (16.313/15.625 =
            #            1.04:1) where a rail is a long bar (31.5/3.0 = 10.5:1)
            elif (gap and not is_full
                  and gap["low_z"] - 0.05 <= p["z"] <= gap["high_z"] + 0.05
                  and base_area and p["face_area"] >= EJECTOR_MIN_AREA_FRAC * base_area
                  and p["width"] > 0
                  and (p["length"] / p["width"]) <= EJECTOR_MAX_ASPECT):
                role, conf = "ejector_stack_plate", "MEDIUM"
                why = (f"sub-footprint and inside the {gap['height']:.3f} rail gap "
                       f"(z={p['z']:.3f} within {gap['low_z']:.3f}..{gap['high_z']:.3f})")
            elif is_full:
                role, conf = "full_footprint_plate", "LOW"
                why = "full footprint, but position alone does not name it"

        d = {
            "index": p["index"], "component": p["component"],
            "thickness": round(p["thickness"], 4), "width": round(p["width"], 4),
            "length": round(p["length"], 4), "z": round(p["z"], 4),
            "role": role, "confidence": conf or "NONE",
            "why": why or "no BOM row, no name token, no positional rule matched",
            "full_footprint": bool(is_full),
            "fill_pct_cad": round(p["fill_pct_cad"], 2),
        }
        if "fill_pct_mesh" in p:
            d["fill_pct_mesh"] = round(p["fill_pct_mesh"], 2)
            d["mesh_volume"] = round(p["mesh"]["mesh_volume"], 3)
            # The disagreement worth surfacing: a real solid cannot be 0% full.
            if p["fill_pct_cad"] <= 0.01 < p["fill_pct_mesh"]:
                d["why"] += (f"  [SolidFillPct says 0.00 but the mesh is "
                             f"{p['fill_pct_mesh']:.1f}% full -- CAD column unreliable here]")
        decisions.append(d)

    return {
        "stack": stack,
        "stl_files": len(meshes),
        "stl_matched": matched,
        "bom_hints": len(hints),
        "decisions": decisions,
    }


def ejector_candidates(job_dir) -> list[dict]:
    """The parts a rail-gap test says are ejector-stack plates, grouped by size."""
    r = decide(job_dir)
    out = [d for d in r.get("decisions", [])
           if d["role"] in ("ejector_stack_plate", "ejector_plate", "bottom_ejector_plate")]
    out.sort(key=lambda d: (-d["thickness"], d["index"]))
    return out
