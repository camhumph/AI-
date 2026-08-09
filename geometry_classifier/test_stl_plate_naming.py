"""Regression tests for STL-driven plate naming.

Run from the repo root:
    set PYTHONPATH=C:\\CMS_AI
    python geometry_classifier\\test_stl_plate_naming.py

The invariants here are the ones that made the C178 naming wrong in the first
place, so each test names the failure it is guarding against:

* A RAIL'S THICKNESS IS ITS STACK HEIGHT. The CAD export sorts a part's three
  sides ascending and calls the smallest "Thickness", which for a rail is its
  width. C17879's rails are 1.875 x 3.000 x 31.500 and the shop orders them
  "3 x 1.875 x 31.5" -- 3" is the height.
* A PLATE THAT IS NOT FULL FOOTPRINT IS STILL A PLATE. C17879's four real
  ejector plates are a third of the mold footprint, and requiring a full
  footprint is exactly why they reached no steel sheet at all.
* A FAILED FEATURE EXTRACTION IS NOT A MEASUREMENT OF ZERO. C17879 index 8 is a
  0.250" sheet with SolidFillPct 0 and no holes recorded. Reading that as a plain
  featureless plate is how it became "Ejector Plate".
* MESH GEOMETRY IS IN MILLIMETRES AND STACKED ON Y. Nothing in an STL records
  units or which axis is up, so both are measured, never assumed.
"""
from __future__ import annotations

import struct
import sys
from pathlib import Path

import numpy as np

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from geometry_classifier import stl_geometry as sg
from geometry_classifier import stl_plate_naming as spn


# ─────────────────────────────────────────────────────────────────────────────
# Synthetic meshes, so these tests need no job folder
# ─────────────────────────────────────────────────────────────────────────────
def _quad(p0, p1, p2, p3):
    return [(p0, p1, p2), (p0, p2, p3)]


def _box(x0, y0, z0, x1, y1, z1):
    """A closed axis-aligned box as a triangle list."""
    return (
        _quad((x0, y0, z0), (x1, y0, z0), (x1, y1, z0), (x0, y1, z0))       # bottom
        + _quad((x0, y0, z1), (x0, y1, z1), (x1, y1, z1), (x1, y0, z1))     # top
        + _quad((x0, y0, z0), (x0, y0, z1), (x1, y0, z1), (x1, y0, z0))     # -y
        + _quad((x0, y1, z0), (x1, y1, z0), (x1, y1, z1), (x0, y1, z1))     # +y
        + _quad((x0, y0, z0), (x0, y1, z0), (x0, y1, z1), (x0, y0, z1))     # -x
        + _quad((x1, y0, z0), (x1, y0, z1), (x1, y1, z1), (x1, y1, z0))     # +x
    )


def _write_binary_stl(path: Path, tris, header=b"test") -> None:
    head = header[:80].ljust(80, b"\x00")
    out = bytearray(head + struct.pack("<I", len(tris)))
    for a, b, c in tris:
        out += struct.pack("<3f", 0.0, 0.0, 0.0)
        for v in (a, b, c):
            out += struct.pack("<3f", *v)
        out += b"\x00\x00"
    path.write_bytes(bytes(out))


def _tris(tris):
    return np.asarray(tris, dtype=np.float64).reshape(-1, 3, 3)


# ─────────────────────────────────────────────────────────────────────────────
def test_reader_roundtrips_binary_and_ascii(tmp: Path):
    tris = _box(0, 0, 0, 2, 3, 4)
    p = tmp / "box.stl"
    _write_binary_stl(p, tris)
    got = sg.read_stl(p)
    assert got.shape == (12, 3, 3), got.shape

    # A binary STL from SolidWorks still starts with the ASCII word "solid" in its
    # header, so format must be decided by the length check, not that prefix.
    p2 = tmp / "solid_prefixed.stl"
    _write_binary_stl(p2, tris, header=b"solid box written by SolidWorks")
    got2 = sg.read_stl(p2)
    assert got2.shape == (12, 3, 3), "a 'solid'-prefixed binary STL must not parse as ASCII"

    ascii_text = "solid s\n" + "".join(
        "facet normal 0 0 0\n outer loop\n"
        + "".join(f"  vertex {v[0]} {v[1]} {v[2]}\n" for v in t)
        + " endloop\nendfacet\n"
        for t in tris
    ) + "endsolid s\n"
    p3 = tmp / "ascii.stl"
    p3.write_text(ascii_text, encoding="utf-8")
    assert sg.read_stl(p3).shape == (12, 3, 3)


def test_volume_and_units_are_measured_not_assumed(tmp: Path):
    # 2 x 3 x 4 inches = 24 cuin.
    g = sg.analyze_body(_tris(_box(0, 0, 0, 2, 3, 4)), stack_axis=1)
    assert abs(g.volume_cuin - 24.0) < 0.01, g.volume_cuin
    assert abs(g.thickness_in - 3.0) < 1e-6, "stack_axis=1 means Y is the thickness"
    assert abs(g.plan_short_in - 2.0) < 1e-6 and abs(g.plan_long_in - 4.0) < 1e-6

    # The same plate in millimetres must come back in inches.
    mm = _tris(_box(0, 0, 0, 2 * 25.4, 3 * 25.4, 4 * 25.4))
    unit, scale = sg.detect_units(mm)
    assert unit == "mm", f"a 101 mm part read as {unit}"
    g_mm = sg.analyze_body(mm * scale, stack_axis=1)
    assert abs(g_mm.volume_cuin - 24.0) < 0.05, g_mm.volume_cuin

    # A genuinely small inch part must not be mistaken for millimetres.
    unit_small, _ = sg.detect_units(_tris(_box(0, 0, 0, 1, 0.5, 2)))
    assert unit_small == "in", "a 2-inch part was misread as millimetres"


def test_stack_axis_is_detected_from_the_bodies():
    # Three plates stacked along Y, each spanning the full X/Z footprint: the
    # real layout of every job in the workspace.
    extents = np.array([[24.0, 1.5, 31.0], [24.0, 4.9, 31.0], [24.0, 3.9, 31.0]])
    assert sg.detect_stack_axis(extents) == 1
    # Same mold rebuilt stacking on Z must be detected as Z, not assumed to be Y.
    assert sg.detect_stack_axis(extents[:, [0, 2, 1]]) == 2


def test_pockets_are_separated_by_which_face_they_open_on():
    """A cavity plate and a core plate differ only in which way the pocket faces.

    This is the single strongest cue in the final pass, so it has to survive the
    rasteriser. The plate here is 6 x 2 x 8 with a 2 x 2 pocket 0.5 deep cut into
    its TOP (+Y) face only.
    """
    outer = _box(0, 0, 0, 6, 2, 8)
    # Remove the top face and rebuild it around a pocket opening.
    body = [t for t in outer if not all(abs(v[1] - 2.0) < 1e-9 for v in t)]
    body += _quad((0, 2, 0), (0, 2, 8), (2, 2, 8), (2, 2, 0))
    body += _quad((4, 2, 0), (4, 2, 8), (6, 2, 8), (6, 2, 0))
    body += _quad((2, 2, 0), (2, 2, 3), (4, 2, 3), (4, 2, 0))
    body += _quad((2, 2, 5), (2, 2, 8), (4, 2, 8), (4, 2, 5))
    # Pocket walls and floor at y = 1.5.
    body += _quad((2, 1.5, 3), (4, 1.5, 3), (4, 1.5, 5), (2, 1.5, 5))
    body += _quad((2, 2, 3), (2, 1.5, 3), (2, 1.5, 5), (2, 2, 5))
    body += _quad((4, 2, 3), (4, 2, 5), (4, 1.5, 5), (4, 1.5, 3))
    body += _quad((2, 2, 3), (4, 2, 3), (4, 1.5, 3), (2, 1.5, 3))
    body += _quad((2, 2, 5), (2, 1.5, 5), (4, 1.5, 5), (4, 2, 5))

    g = sg.analyze_body(_tris(body), stack_axis=1, cell_in=0.05, cross_axis=False)
    assert g.pocket_top_in2 > 3.0, f"a 2x2 pocket on the top face measured {g.pocket_top_in2}"
    assert abs(g.pocket_top_max_depth - 0.5) < 0.06, g.pocket_top_max_depth
    assert g.pocket_bottom_in2 < 0.5, (
        f"the bottom face is solid but measured {g.pocket_bottom_in2} in2 of pocket -- "
        "top and bottom have been confused"
    )


def test_an_ejector_housing_is_seen_as_clamp_plate_plus_rails():
    """One CAD body, three pieces of steel.

    C18027's ejector housing is modelled as a single U-shaped solid: a 0.875"
    bottom clamp plate with two 3.000" rails standing on it. The CAD export lists
    it as one 3.875" plate, so the steel sheet ordered a solid 3.875" block and
    the two rails were never quoted at all. Nothing in the export can reveal
    this -- only the mesh can.

    Built here as a 8 x 3.875 x 12 channel with a 5.5-wide pocket 3.000 deep.
    """
    W, H, L = 8.0, 3.875, 12.0
    floor, wall = 0.875, 3.0
    x0, x1 = 1.25, 6.75  # the channel opening, leaving 1.25 of rail each side
    body = []
    body += _quad((0, 0, 0), (W, 0, 0), (W, 0, L), (0, 0, L))            # underside
    for a, b in ((0.0, x0), (x1, W)):                                     # rail tops
        body += _quad((a, H, 0), (a, H, L), (b, H, L), (b, H, 0))
    body += _quad((x0, floor, 0), (x1, floor, 0), (x1, floor, L), (x0, floor, L))
    body += _quad((x0, floor, 0), (x0, floor, L), (x0, H, L), (x0, H, 0))
    body += _quad((x1, floor, 0), (x1, H, 0), (x1, H, L), (x1, floor, L))
    body += _quad((0, 0, 0), (0, 0, L), (0, H, L), (0, H, 0))
    body += _quad((W, 0, 0), (W, H, 0), (W, H, L), (W, 0, L))
    body += _quad((0, 0, 0), (0, H, 0), (x0, H, 0), (x0, floor, 0))
    body += _quad((x1, 0, 0), (x1, floor, 0), (W, H, 0), (W, 0, 0))
    body += _quad((x0, 0, 0), (x0, floor, 0), (x1, floor, 0), (x1, 0, 0))
    body += _quad((0, 0, L), (x0, floor, L), (x0, H, L), (0, H, L))
    body += _quad((x1, 0, L), (W, 0, L), (W, H, L), (x1, floor, L))
    body += _quad((x0, 0, L), (x1, 0, L), (x1, floor, L), (x0, floor, L))

    g = sg.analyze_body(_tris(body), stack_axis=1, cell_in=0.05, cross_axis=False)
    assert g.is_open_channel, (
        f"a {H}in body that is only {g.full_thickness_pct:.0f}% full height, with a "
        f"{g.pocket_top_max_depth:.2f}in pocket, was not recognised as an ejector housing"
    )
    assert abs(g.channel_depth_in - wall) < 0.08, f"rail height {g.channel_depth_in}"
    assert abs(g.channel_floor_in - floor) < 0.08, f"clamp thickness {g.channel_floor_in}"
    assert len(g.standing_walls) == 2, (
        f"{len(g.standing_walls)} rails found, expected 2 -- bolt holes through a "
        f"rail must not split it into fragments"
    )
    for w in g.standing_walls:
        assert abs(w.height_in - wall) < 0.08, "a rail stands the channel depth, not the whole body"
        assert abs(w.width_in - 1.25) < 0.1, f"rail width {w.width_in}"
        assert abs(w.length_in - L) < 0.15, f"rail length {w.length_in}"

    # A plain solid plate must NOT trip this.
    plain = sg.analyze_body(_tris(_box(0, 0, 0, 8, 1.375, 12)), stack_axis=1,
                            cell_in=0.05, cross_axis=False)
    assert not plain.is_open_channel, "a solid plate was mistaken for an ejector housing"


def test_manifold_plate_needs_hot_runner_evidence():
    """manifold_plate was handed out on position alone.

    On C18027 that invented a manifold this shop does not run -- and because the
    stack is named outward from the A/B pair, it pushed the real A plate down to
    b_plate and the real B plate down to support_plate. One unjustified name
    silently re-labelled three plates.
    """
    from geometry_classifier.qwen_classify_xt_csv import (
        _looks_like_hot_runner,
        _stack_names_around_ab,
    )

    # C18027 index 3: cross-drilling reported by CAD, but the mesh finds one
    # waterline, and the bore is a sprue, not a runner channel.
    assert not _looks_like_hot_runner(
        {"NCrossAxis": "1", "MaxBoreDia": "2.438", "Thickness": "1.375"}
    )
    # A genuine hot-runner plate: drilled across, bored, and thick with it.
    assert _looks_like_hot_runner(
        {"NCrossAxis": "12", "MaxBoreDia": "1.500", "Thickness": "2.500"}
    )

    # Five plates, A/B found one position too low, no hot-runner evidence.
    plates = [
        {"NCrossAxis": "0", "MaxBoreDia": "0.5", "Thickness": "0.875"},   # top clamp
        {"NCrossAxis": "1", "MaxBoreDia": "2.438", "Thickness": "1.375"}, # the real A
        {"NCrossAxis": "1", "MaxBoreDia": "1.438", "Thickness": "1.375"}, # the real B
        {"NCrossAxis": "0", "MaxBoreDia": "0.812", "Thickness": "1.375"}, # support
        {"NCrossAxis": "0", "MaxBoreDia": "1.0", "Thickness": "3.875"},   # bottom clamp
    ]
    names = _stack_names_around_ab(plates, a_at=2, top_clamp_present=True)
    assert "manifold_plate" not in names, f"unevidenced manifold survived: {names}"
    assert names == [
        "top_clamp_plate", "a_plate", "b_plate", "support_plate", "bottom_clamp_plate"
    ], names

    # With real evidence the manifold stays and nothing shifts.
    hot = list(plates)
    hot[1] = {"NCrossAxis": "14", "MaxBoreDia": "1.25", "Thickness": "2.000"}
    names_hot = _stack_names_around_ab(hot, a_at=2, top_clamp_present=True)
    assert names_hot[1] == "manifold_plate", names_hot
    assert names_hot[2] == "a_plate", names_hot


def test_volume_joins_a_mesh_to_its_cad_row():
    """The join key. SolidWorks' bbox x fill and mesh integration agree closely,
    and no two plates in a base share a volume."""
    rows = [
        # 4.875 x 11.875 x 23.5 at 40.383% fill -> 549.4 cuin (real C17880 A plate)
        {"Index": "1", "Component": "a", "Thickness": "4.875", "Width": "11.875",
         "Length": "23.500", "BBoxVolume_cuin": "1360.433", "SolidFillPct": "40.383",
         "CenterX": "-6.25", "CenterY": "0", "CenterZ": "2.440", "Qty": "1"},
        # 1.875 x 11.875 x 23.5 at 50.943% -> 266.6 cuin (the B plate)
        {"Index": "3", "Component": "b", "Thickness": "1.875", "Width": "11.875",
         "Length": "23.500", "BBoxVolume_cuin": "523.244", "SolidFillPct": "50.943",
         "CenterX": "-6.25", "CenterY": "0", "CenterZ": "-0.939", "Qty": "1"},
    ]
    parsed = [
        {
            "index": r["Index"], "component": r["Component"], "qty": 1,
            "thickness": float(r["Thickness"]), "width": float(r["Width"]),
            "length": float(r["Length"]), "bbox_volume": float(r["BBoxVolume_cuin"]),
            "solid_fill_pct": float(r["SolidFillPct"]),
            "center_x": float(r["CenterX"]), "center_y": float(r["CenterY"]),
            "center_z": float(r["CenterZ"]),
            "n_thru": 0, "n_cbore": 0, "n_cross": 0, "max_bore": 0.0,
            "hole_sig": "", "n_pockets": 0, "pocket_area": 0.0,
            "pocket_up": 0.0, "pocket_dn": 0.0, "raw": r,
        }
        for r in rows
    ]
    assert abs(spn.csv_volume(parsed[0]) - 549.4) < 0.5, spn.csv_volume(parsed[0])
    assert abs(spn.csv_volume(parsed[1]) - 266.6) < 0.5, spn.csv_volume(parsed[1])
    # The two volumes are far enough apart that a 2% join can never cross them.
    assert spn.csv_volume(parsed[0]) / spn.csv_volume(parsed[1]) > 1.5


def _model_from(rows_spec):
    """Build a StackModel from (index, thk, w, l, cx, cy, cz, fill) tuples."""
    parsed = []
    for idx, thk, w, l, cx, cy, cz, fill in rows_spec:
        parsed.append(
            {
                "index": str(idx), "component": f"part-{idx}", "qty": 1,
                "thickness": thk, "width": w, "length": l,
                "bbox_volume": thk * w * l, "solid_fill_pct": fill,
                "center_x": cx, "center_y": cy, "center_z": cz,
                "n_thru": 0, "n_cbore": 0, "n_cross": 0, "max_bore": 0.0,
                "hole_sig": "", "n_pockets": 0, "pocket_area": 0.0,
                "pocket_up": 0.0, "pocket_dn": 0.0, "raw": {},
            }
        )
    return spn.build_stack_model(parsed, job="TEST")


# C17879's real stack, trimmed to what the derivation needs.
C17879_STACK = [
    # idx  thk     w       l      cx     cy      cz     fill
    ("8",  0.250, 23.500, 31.250,  0.0,   0.0,  10.752,  0.0),   # the mystery sheet
    ("4",  1.875, 23.750, 31.500,  0.0,   0.0,   9.689, 96.2),   # top clamp
    ("1",  4.875, 23.750, 31.500,  0.0,   0.0,   6.314, 97.0),   # A
    ("2",  3.875, 23.750, 31.500,  0.0,   0.0,   1.939, 69.7),   # B
    ("3",  3.875, 23.750, 31.500,  0.0,   0.0,  -1.939, 72.6),   # support
    ("11", 0.625, 15.625, 16.313,  8.531, 0.0,  -5.252, 100.0),  # split ejector
    ("12", 0.625, 15.625, 16.313, -8.531, 0.0,  -5.252, 100.0),
    ("9",  1.875,  3.000, 31.500,  0.0,  -8.813, -5.377, 94.8),  # rails
    ("10", 1.875,  3.000, 31.500,  0.0,   8.813, -5.377, 94.8),
    ("6",  1.125, 15.625, 16.313,  8.531, 0.0,  -6.127, 100.0),
    ("7",  1.125, 15.625, 16.313, -8.531, 0.0,  -6.127, 100.0),
    ("5",  1.375, 23.750, 31.500,  0.0,   0.0,  -7.564, 98.4),   # bottom clamp
]


def test_ejector_box_comes_from_the_plate_gap_not_the_rails():
    """The box must be found without trusting a rail's thickness.

    A rail's stack height is precisely what the CAD export cannot say, so locating
    the box from the rails is circular. C17879's opening is 3.000" -- which is what
    the shop's own steel sheet orders the rails as.
    """
    m = _model_from(C17879_STACK)
    assert m.ejector_box is not None, "no ejector box found in a base that plainly has one"
    lo, hi = m.ejector_box
    assert abs((hi - lo) - 3.000) < 0.01, f"box height {hi - lo:.4f}, expected 3.000"


def test_rail_thickness_becomes_the_stack_height():
    m = _model_from(C17879_STACK)
    rails = [p for p in m.parts if p.index in {"9", "10"}]
    assert rails, "rails missing from the model"
    for r in rails:
        assert r.is_rail_like, f"index {r.index} was not recognised as a rail"
        assert abs(r.thickness - 3.000) < 0.01, (
            f"rail {r.index} thickness is {r.thickness:.4f}; the CAD export's sorted "
            f"1.875 was used instead of the 3.000 stack height"
        )
        assert r.in_stack, "a rail is quoted steel and must appear in the stack listing"


def test_half_footprint_ejector_plates_stay_in_the_stack():
    m = _model_from(C17879_STACK)
    for idx in ("6", "7", "11", "12"):
        p = m.by_index()[idx]
        assert p.in_stack, (
            f"index {idx} ({p.plan_area_frac:.0%} of the footprint) was dropped from the "
            f"stack; this is the C17879 bug where four real ejector plates reached no sheet"
        )
        assert p.inside_ejector_box, f"index {idx} sits between the rails but was not flagged"
        assert p.mirror_twins, f"index {idx} is half of a split set but has no twin"


def test_a_sheet_above_the_top_clamp_plate_is_outside_the_stack():
    """C17879 index 8 sits above everything and was named "Ejector Plate".

    The fix is not a rule that renames it -- it is that the model can see it is
    top of stack with no measured geometry, so it cannot be inside the ejector box.
    """
    m = _model_from(C17879_STACK)
    p = m.by_index()["8"]
    assert p.order_from_top == 1, f"the 0.250 sheet is at order {p.order_from_top}, expected 1"
    assert not p.inside_ejector_box, "a part above the top clamp plate cannot be in the ejector box"
    assert p.weight_lb is None, (
        "a part whose feature walk failed must report an unknown weight, not 0 lb"
    )
    assert p.weight_lb_if_solid > 0, "an upper-bound weight should still be available"


# C18184's real stack: five plates on a 9.875 x 11.875 footprint stacked along
# **Y**, with CenterZ 0.000 on every one of them.
C18184_STACK = [
    # idx  thk     w       l      cx     cy      cz    fill
    ("3",  1.875, 9.875, 11.875, 0.0,  2.813,  0.0,  84.8),
    ("2",  1.875, 9.875, 11.875, 0.0,  0.937,  0.0,  73.1),
    ("5",  1.375, 9.875, 11.875, 0.0, -0.688,  0.0,  90.9),
    ("4",  1.375, 9.875, 11.875, 0.0, -2.063,  0.0,  67.7),
    ("1",  3.625, 9.875, 11.875, 0.0, -4.563,  0.0,  90.7),
]


def test_the_stack_axis_is_measured_not_assumed():
    """Not every base stacks along Z.

    C18184 stacks along Y with CenterZ 0.000 on all five plates. Reading Z there
    put every plate on one stack level, which makes order, gaps, the ejector box
    and mirror twins all meaningless at once. The deterministic rules already got
    this right for that job ("stack_axis": "CenterY").
    """
    m = _model_from(C18184_STACK)
    assert m.stack_axis_field == "center_y", (
        f"stack axis detected as {m.stack_axis_field}; C18184 stacks along Y"
    )
    levels = spn._stack_levels([p for p in m.parts if p.in_stack])
    assert len(levels) == 5, f"{len(levels)} stack levels for a 5-plate stack"

    order = [p.index for p in m.in_stack_parts()]
    assert order == ["3", "2", "5", "4", "1"], f"top-to-bottom order wrong: {order}"
    for p in m.parts:
        assert p.stack_pos == p.center_y, "stack_pos must follow the detected axis"

    # And a Z-stacked base must still come out as Z.
    z = _model_from(C17879_STACK)
    assert z.stack_axis_field == "center_z", z.stack_axis_field


def test_side_by_side_parts_share_a_stack_level():
    """Twins at one height are one level. Treating them as two produced negative
    gaps, which read as an interference that is not there."""
    m = _model_from(C17879_STACK)
    a, b = m.by_index()["11"], m.by_index()["12"]
    assert a.order_from_top == b.order_from_top, "mirror twins landed on different levels"
    for p in m.parts:
        if p.is_rail_like:
            continue  # rails stand alongside, so they carry no gaps at all
        for gap in (p.gap_above, p.gap_below):
            assert gap is None or gap >= 0.0, f"index {p.index} has a negative gap {gap}"


def test_the_model_is_never_shown_the_macros_old_guess(tmp: Path):
    """The macro writes its current role into the STL name AND its mesh header.

    Feeding either back as evidence is what lets a wrong name survive a re-run, so
    neither may reach a prompt.
    """
    p = tmp / "Job-C17880_Plate 4.STL"
    _write_binary_stl(p, _box(0, 0, 0, 2, 1, 4), header=b"solid Job-C17880_Plate 4")
    assert "Plate 4" in sg._solid_name(p), "test fixture is not carrying the stale label"

    m = _model_from(C17879_STACK)
    part = m.by_index()["11"]
    part.stl = sg.analyze_body(
        sg.read_stl(p), stack_axis=1, source=p.name,
        label_hint=sg._solid_name(p), cross_axis=False,
    )
    spn.refresh_derived(m)
    prompt = spn.build_final_prompt(m, "x.csv", ["a_plate", "ejector_plate"], rules=None)
    assert "Plate 4" not in prompt, "the macro's stale plate name leaked into the final prompt"
    assert p.name not in prompt, "the STL filename leaked into the final prompt"


def test_hardware_is_not_handed_to_the_model():
    """127 fasteners against 10 plates: asking for all of them spends the whole
    generation on the parts that were never wrong."""
    stack = list(C17879_STACK) + [
        (f"h{i}", 0.5, 0.5, 1.25, 13.0, 8.7, -6.877, 0.0) for i in range(40)
    ]
    m = _model_from(stack)
    struct = spn.structural_parts(m)
    hw = spn.hardware_parts(m)
    assert len(struct) == 12, f"expected the 12 structural parts, got {len(struct)}"
    assert len(hw) == 40, f"expected 40 hardware parts, got {len(hw)}"

    rules = {"classifications": [{"index": p.index, "role": "ejector_pin"} for p in hw]}
    prompt = spn.build_candidate_prompt(m, "x.csv", ["a_plate", "ejector_pin"], rules=rules)
    assert "do not re-classify" in prompt.lower()
    assert '"rules_role":"ejector_pin"' in prompt.replace(" ", ""), (
        "hardware must arrive pre-named so the model can spot a plate hiding in it"
    )

    # And measuring must not be requested for parts that are plainly fasteners.
    wanted = spn.candidates_to_measure(m, None)
    assert not any(i.startswith("h") for i in wanted), (
        f"fasteners were queued for the mesh pass: {sorted(i for i in wanted if i.startswith('h'))[:5]}"
    )


def test_a_short_answer_is_patched_not_discarded():
    m = _model_from(C17879_STACK)
    roles = ["a_plate", "b_plate", "top_clamp_plate", "hardware_other", "full_footprint_plate"]
    partial = {"classifications": [{"index": "1", "role": "a_plate", "confidence": "HIGH"}]}
    fallback = {
        "classifications": [
            {"index": p.index, "role": "full_footprint_plate", "quote": True} for p in m.parts
        ]
    }
    expect = {p.index for p in spn.structural_parts(m)}
    data, problems = spn.validate_classifications(partial, m, roles, expect=expect)
    assert problems and "not classified" in problems[0]
    added = spn.fill_missing(data, fallback, m, "geometry rules")
    assert added == len(m.parts) - 1
    assert len(data["classifications"]) == len(m.parts)
    patched = next(c for c in data["classifications"] if c["index"] == "3")
    assert "[geometry rules]" in patched["reason"], "a patched row must say where it came from"


def test_a_structurally_impossible_answer_is_caught():
    """A weak local model must not be able to put nonsense on a steel sheet.

    qwen3:8b named C17880's topmost plate "a_plate" and the A plate below it
    "top_clamp_plate", then put the B plate four levels further down -- and the
    mesh pass confirmed all of it rather than correcting any of it. The result
    was materially worse than the deterministic rules it was meant to improve on.

    These checks are not naming opinions. A top clamp plate that is not on top
    cannot be true of any mold base, whatever the shop calls its plates.
    """
    stack = [
        # idx  thk     w       l      cx    cy      cz      fill
        ("4",  1.454, 11.875, 23.5,  0.0,  0.0,  5.604,  92.7),
        ("1",  4.875, 11.875, 23.5,  0.0,  0.0,  2.440,  40.4),
        ("3",  1.875, 11.875, 23.5,  0.0,  0.0, -0.939,  50.9),
        ("5",  1.375, 11.875, 23.5,  0.0,  0.0, -2.565,  83.3),
        ("2",  1.878, 11.875, 23.5,  0.0,  0.0, -4.188,  83.1),
        ("6",  0.875, 11.875, 25.5,  0.0,  0.0, -9.565,  94.2),
    ]
    m = _model_from(stack)

    # What qwen3:8b actually returned for this job.
    bad = {"classifications": [
        {"index": "4", "role": "a_plate"},
        {"index": "1", "role": "top_clamp_plate"},
        {"index": "3", "role": "support_plate"},
        {"index": "2", "role": "b_plate"},
        {"index": "6", "role": "bottom_clamp_plate"},
    ]}
    problems = spn.structural_sanity_problems(bad, m)
    assert any("topmost" in p for p in problems), problems
    assert any("levels apart" in p for p in problems), problems

    # The correct reading trips nothing.
    good = {"classifications": [
        {"index": "4", "role": "top_clamp_plate"},
        {"index": "1", "role": "a_plate"},
        {"index": "3", "role": "b_plate"},
        {"index": "5", "role": "support_plate"},
        {"index": "2", "role": "support_plate"},
        {"index": "6", "role": "bottom_clamp_plate"},
    ]}
    assert spn.structural_sanity_problems(good, m) == [], spn.structural_sanity_problems(good, m)

    # B above A is impossible whichever way round the rest is.
    flipped = {"classifications": [
        {"index": "1", "role": "b_plate"},
        {"index": "3", "role": "a_plate"},
    ]}
    assert any("sits above" in p for p in spn.structural_sanity_problems(flipped, m))


def test_the_cad_half_tokens_settle_the_parting_line():
    """When the CAD tree says which half a plate is in, that is not negotiable."""
    rows = [
        {"index": "1", "component": "24-258--bs-quote_1-1/24-258--1200-a00_1-1", "qty": 1,
         "thickness": 4.875, "width": 11.875, "length": 23.5, "bbox_volume": 1360.4,
         "solid_fill_pct": 40.4, "center_x": 0.0, "center_y": 0.0, "center_z": 2.44,
         "n_thru": 0, "n_cbore": 0, "n_cross": 0, "max_bore": 0.0, "hole_sig": "",
         "n_pockets": 0, "pocket_area": 0.0, "pocket_up": 0.0, "pocket_dn": 0.0, "raw": {}},
        {"index": "3", "component": "24-258--bm-quote_1-1/24-258--1500-a00_1-1", "qty": 1,
         "thickness": 1.875, "width": 11.875, "length": 23.5, "bbox_volume": 523.2,
         "solid_fill_pct": 50.9, "center_x": 0.0, "center_y": 0.0, "center_z": -0.939,
         "n_thru": 0, "n_cbore": 0, "n_cross": 0, "max_bore": 0.0, "hole_sig": "",
         "n_pockets": 0, "pocket_area": 0.0, "pocket_up": 0.0, "pocket_dn": 0.0, "raw": {}},
    ]
    m = spn.build_stack_model(rows, job="C17880")
    assert m.by_index()["1"].half == "stationary", m.by_index()["1"].half
    assert m.by_index()["3"].half == "moving", m.by_index()["3"].half

    swapped = {"classifications": [
        {"index": "1", "role": "b_plate"},
        {"index": "3", "role": "a_plate"},
    ]}
    problems = spn.structural_sanity_problems(swapped, m)
    assert any("stationary half" in p or "moving half" in p for p in problems), problems


def test_a_second_a_plate_is_caught():
    m = _model_from(C17879_STACK)
    roles = ["a_plate", "full_footprint_plate"]
    doubled = {
        "classifications": [
            {"index": "1", "role": "a_plate", "confidence": "HIGH"},
            {"index": "2", "role": "a_plate", "confidence": "HIGH"},
        ]
    }
    data, problems = spn.validate_classifications(doubled, m, roles, expect={"1", "2"})
    assert any("second a_plate" in p for p in problems), problems
    kept = {c["index"]: c["role"] for c in data["classifications"]}
    assert kept["1"] == "a_plate" and kept["2"] == "full_footprint_plate"


if __name__ == "__main__":
    import tempfile

    with tempfile.TemporaryDirectory() as d:
        tmp = Path(d)
        test_reader_roundtrips_binary_and_ascii(tmp)
        test_volume_and_units_are_measured_not_assumed(tmp)
        test_stack_axis_is_detected_from_the_bodies()
        test_pockets_are_separated_by_which_face_they_open_on()
        test_an_ejector_housing_is_seen_as_clamp_plate_plus_rails()
        test_manifold_plate_needs_hot_runner_evidence()
        test_volume_joins_a_mesh_to_its_cad_row()
        test_ejector_box_comes_from_the_plate_gap_not_the_rails()
        test_rail_thickness_becomes_the_stack_height()
        test_half_footprint_ejector_plates_stay_in_the_stack()
        test_a_sheet_above_the_top_clamp_plate_is_outside_the_stack()
        test_the_stack_axis_is_measured_not_assumed()
        test_side_by_side_parts_share_a_stack_level()
        test_the_model_is_never_shown_the_macros_old_guess(tmp)
        test_hardware_is_not_handed_to_the_model()
        test_a_short_answer_is_patched_not_discarded()
        test_a_structurally_impossible_answer_is_caught()
        test_the_cad_half_tokens_settle_the_parting_line()
        test_a_second_a_plate_is_caught()
    print("OK: STL plate naming regressions passed")
