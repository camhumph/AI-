"""Read real geometry out of the exported STL meshes.

Why this exists
---------------
Everything downstream of Module6121 names plates from
``XT_Export_CAD_Dimensions.csv``, which carries only a sorted bounding box plus
whatever hole/pocket counts the SolidWorks feature walk managed to collect. That
export loses two things the naming depends on:

* **Which box side is the stack axis.** The CSV sorts Thickness/Width/Length
  ascending, so a rail's real stack height is thrown away and the smallest side
  takes its place.
* **Anything the feature walk missed.** In the C178 jobs a large share of rows
  come back ``SolidFillPct=0`` with zero holes and zero pockets -- the walk gave
  up, and the plate then has nothing left to name it by. C17879 index 8 is the
  clearest case: a 0.250" sheet with no features at all, which the naming rules
  then labelled "Ejector Plate" while the four real ejector plates went unnamed.

The STL has neither problem. It is the same solid, in the frame the macro
already oriented, and its geometry cannot "fail to extract" -- the triangles
*are* the part.

Facts about these files, established by measuring them
-----------------------------------------------------
* **Units are millimetres**, despite ``FORCE_INCH_STL_EXPORT`` in the macro.
  Nothing in an STL records units, so this module measures and decides; see
  :func:`detect_units`. Everything it returns is inches.
* **The stack axis is Y, not Z.** A job's ``XT_Export_CAD_Dimensions.csv``
  ``CenterZ`` equals the STL's Y centre to the thousandth. This module therefore
  never assumes an axis: it reports all three extents and takes ``stack_axis``
  from the caller, with :func:`detect_stack_axis` to infer it from a set of
  bodies.
* **The merged ``<job>.stl`` cannot be split into parts.** Plates that touch
  face to face share welded vertices, so connected components fuse them --
  C17879's top clamp + A + B come back as one 10.625" body. Use the per-plate
  files in the job's ``stl\\`` folder, which are one clean body each.
* **``<job> component.stl`` is the complement, not a superset.** It is the
  assembly with the quoted plates *hidden*, so it holds hardware only.

This module is deliberately dependency-free apart from numpy. The shop PC runs
``pip install -r requirements.txt`` from a batch file on every start, and adding
trimesh/scipy there means a compiler toolchain and a much slower cold start for
no capability we need.
"""
from __future__ import annotations

import math
import struct
from dataclasses import dataclass, field, asdict
from pathlib import Path

import numpy as np

# Rasterisation cell size, in inches, for the plan-view pass. 0.05" resolves a
# 1/4" hole into ~20 cells, which is enough to count it and size it; smaller
# holes are found by the wall-cluster pass instead, which is exact.
DEFAULT_CELL_IN = 0.05

# Refuse to build a plan grid bigger than this; coarsen the cell instead.
MAX_GRID_CELLS = 1_200_000

# Cap on how many (triangle, cell) pairs we materialise at once, to keep the
# vectorised rasteriser's peak memory near 100 MB rather than unbounded.
RASTER_CHUNK_PAIRS = 2_000_000

# Two vertices closer than this in every axis are the same vertex. Tight on
# purpose: separate plates in an assembly sit face to face, and a loose weld
# fuses them into one body.
WELD_TOL_IN = 1e-5

MM_PER_IN = 25.4


# ─────────────────────────────────────────────────────────────────────────────
# Reading
# ─────────────────────────────────────────────────────────────────────────────
def read_stl(path: str | Path) -> np.ndarray:
    """Return an ``(n_triangles, 3, 3)`` float64 array of triangle corners.

    Handles binary and ASCII STL. A binary STL written by SolidWorks still
    begins with the ASCII word ``solid`` inside its 80-byte header, so the format
    is decided by whether the triangle count matches the file length -- never by
    that prefix.
    """
    raw = Path(path).read_bytes()
    if len(raw) < 84:
        return _read_ascii_stl(raw.decode("utf-8", errors="replace"))

    n_tris = struct.unpack("<I", raw[80:84])[0]
    if len(raw) == 84 + 50 * n_tris:
        return _read_binary_stl(raw, n_tris)
    return _read_ascii_stl(raw.decode("utf-8", errors="replace"))


def _read_binary_stl(raw: bytes, n_tris: int) -> np.ndarray:
    # Each record is 50 bytes: 3 normal floats, 9 corner floats, 2 attribute
    # bytes. Reading it as a 50-byte stride and slicing the corners out avoids a
    # per-triangle Python loop entirely.
    rec = np.frombuffer(raw, dtype=np.uint8, count=50 * n_tris, offset=84)
    rec = rec.reshape(n_tris, 50)
    corners = rec[:, 12:48].copy().view("<f4").reshape(n_tris, 3, 3)
    return corners.astype(np.float64)


def _read_ascii_stl(text: str) -> np.ndarray:
    verts: list[tuple[float, float, float]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line.startswith("vertex"):
            continue
        parts = line.split()
        if len(parts) >= 4:
            try:
                verts.append((float(parts[1]), float(parts[2]), float(parts[3])))
            except ValueError:
                continue
    if not verts:
        return np.zeros((0, 3, 3), dtype=np.float64)
    arr = np.asarray(verts, dtype=np.float64)
    # Drop a trailing partial triangle rather than raising: a truncated export
    # should still yield the bodies that did make it out.
    return arr[: (len(arr) // 3) * 3].reshape(-1, 3, 3)


# ─────────────────────────────────────────────────────────────────────────────
# Units
# ─────────────────────────────────────────────────────────────────────────────
def _fraction_snap_score(dims: np.ndarray) -> float:
    """How close ``dims`` sit to 1/64" multiples, averaged. 1.0 is a perfect fit.

    Mold steel is specified in binary fractions, so the correct unit reading
    lands on 1/64" boundaries and the wrong one scatters. This is the tiebreak
    when a part is small enough that its size alone is ambiguous.
    """
    if not len(dims):
        return 0.0
    sixty_fourths = dims * 64.0
    err = np.abs(sixty_fourths - np.round(sixty_fourths))
    return float(1.0 - np.mean(np.minimum(err, 0.5)) * 2.0)


def detect_units(tris: np.ndarray) -> tuple[str, float]:
    """Guess whether a mesh is in millimetres or inches.

    Returns ``(unit_name, scale_to_inches)``.

    A mold plate runs roughly 0.1" to 120". Anything whose longest side exceeds
    150 in file units is millimetres -- a 150" plate does not exist in this shop,
    and 150 mm is a routine 5.9" block. In the ambiguous middle the 1/64"
    snapping test decides.
    """
    if len(tris) == 0:
        return "in", 1.0
    pts = tris.reshape(-1, 3)
    extent = pts.max(axis=0) - pts.min(axis=0)
    longest = float(extent.max())

    if longest > 150.0:
        return "mm", 1.0 / MM_PER_IN
    if longest < 6.0:
        # Under 6 file units: as mm that is a 1/4" chip, so inches is the only
        # reading that describes a real part.
        return "in", 1.0

    as_in = _fraction_snap_score(extent)
    as_mm = _fraction_snap_score(extent / MM_PER_IN)
    if as_mm > as_in + 0.02:
        return "mm", 1.0 / MM_PER_IN
    return "in", 1.0


# ─────────────────────────────────────────────────────────────────────────────
# Splitting a multi-body mesh
# ─────────────────────────────────────────────────────────────────────────────
def _union_find(n: int, pairs_a: np.ndarray, pairs_b: np.ndarray) -> np.ndarray:
    """Label ``n`` items into components given edges ``(pairs_a[i], pairs_b[i])``."""
    parent = np.arange(n, dtype=np.int64)

    def find(i: int) -> int:
        root = i
        while parent[root] != root:
            root = parent[root]
        while parent[i] != root:
            parent[i], i = root, parent[i]
        return root

    for a, b in zip(pairs_a.tolist(), pairs_b.tolist()):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[max(ra, rb)] = min(ra, rb)
    return np.array([find(int(i)) for i in range(n)], dtype=np.int64)


def _weld(tris: np.ndarray, weld_tol: float = WELD_TOL_IN) -> np.ndarray:
    """Per-corner vertex ids, with coincident corners sharing an id."""
    keys = np.round(tris.reshape(-1, 3) / weld_tol).astype(np.int64)
    _, vert_ids = np.unique(keys, axis=0, return_inverse=True)
    return vert_ids.reshape(-1, 3)


def split_bodies(
    tris: np.ndarray,
    weld_tol: float = WELD_TOL_IN,
    min_triangles: int = 12,
) -> list[np.ndarray]:
    """Split ``tris`` into connected components, largest first.

    Only useful on a mesh whose parts do not touch -- see the module docstring on
    ``<job>.stl``, where face-to-face plates fuse into one component.
    """
    if len(tris) == 0:
        return []
    vert_ids = _weld(tris, weld_tol)

    # Any two triangles sharing a welded vertex are one body. Sorting
    # (vertex_id, triangle) lets us emit each vertex's triangle run as edges to a
    # single spanning triangle, which is all union-find needs.
    order = np.argsort(vert_ids.ravel(), kind="stable")
    sorted_v = vert_ids.ravel()[order]
    sorted_t = (order // 3).astype(np.int64)
    run_start = np.concatenate(([0], np.flatnonzero(np.diff(sorted_v)) + 1))
    # For each entry, the first triangle of its vertex's run: one edge per entry.
    first_of_run = np.repeat(
        sorted_t[run_start], np.diff(np.concatenate((run_start, [len(sorted_v)])))
    )
    roots = _union_find(len(tris), first_of_run, sorted_t)

    bodies: list[np.ndarray] = []
    for root in np.unique(roots):
        idx = np.flatnonzero(roots == root)
        if len(idx) >= min_triangles:
            bodies.append(idx)
    bodies.sort(key=len, reverse=True)
    return bodies


# ─────────────────────────────────────────────────────────────────────────────
# Feature containers
# ─────────────────────────────────────────────────────────────────────────────
@dataclass
class HoleGroup:
    """Holes of one nominal size sharing one axis, counted together."""

    axis: str                    # "thickness" (through the faces) or "cross"
    diameter_in: float
    count: int
    through: bool
    depth_in: float

    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class BodyGeometry:
    """Everything measurable about one solid body. All lengths in inches."""

    # --- identity / provenance ----------------------------------------------
    source: str = ""             # file this body came from
    body_index: int = 0
    label_hint: str = ""         # name in the file/header, UNTRUSTED (see below)
    n_triangles: int = 0
    units: str = "in"

    # --- position and size, in the assembly frame ---------------------------
    # ext/center are indexed by raw axis (0=x, 1=y, 2=z) so nothing is sorted
    # away. thickness_in is the extent along the stack axis -- the real plate
    # thickness, which is the whole point of reading the mesh.
    ext_x: float = 0.0
    ext_y: float = 0.0
    ext_z: float = 0.0
    center_x: float = 0.0
    center_y: float = 0.0
    center_z: float = 0.0
    stack_axis: int = 1
    stack_center: float = 0.0    # position up the stack
    stack_min: float = 0.0
    stack_max: float = 0.0
    thickness_in: float = 0.0
    plan_short_in: float = 0.0   # smaller in-plane side
    plan_long_in: float = 0.0    # larger in-plane side

    # --- mass properties ----------------------------------------------------
    volume_cuin: float = 0.0
    surface_area_in2: float = 0.0
    bbox_volume_cuin: float = 0.0
    solid_fill_pct: float = 0.0

    # --- plan view ----------------------------------------------------------
    footprint_in2: float = 0.0      # outline area, through-holes counted inside
    solid_section_in2: float = 0.0  # outline minus through-holes
    outline_fill_pct: float = 0.0   # footprint / (plan_short*plan_long); 100 = rectangle
    full_thickness_in2: float = 0.0
    full_thickness_pct: float = 0.0

    # --- pockets, split by which face they open on --------------------------
    # "top" means the +stack_axis face. Which side a pocket opens on separates a
    # cavity plate from a core plate, and an ejector retainer from its back-up.
    pocket_top_in2: float = 0.0
    pocket_top_max_depth: float = 0.0
    pocket_bottom_in2: float = 0.0
    pocket_bottom_max_depth: float = 0.0
    n_pocket_regions_top: int = 0
    n_pocket_regions_bottom: int = 0

    # --- faces --------------------------------------------------------------
    top_face_in2: float = 0.0
    bottom_face_in2: float = 0.0
    n_solid_spans_max: int = 0    # >1 means a ray re-enters: internal structure

    # --- holes --------------------------------------------------------------
    holes: list[HoleGroup] = field(default_factory=list)
    n_thru_holes: int = 0
    n_counterbores: int = 0
    n_cross_holes: int = 0
    max_bore_dia: float = 0.0
    hole_signature: str = ""

    # --- diagnostics --------------------------------------------------------
    watertight_ratio: float = 0.0
    notes: list[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        d = asdict(self)
        d["holes"] = [h.to_dict() if isinstance(h, HoleGroup) else h for h in self.holes]
        return d


# ─────────────────────────────────────────────────────────────────────────────
# Plan-view rasterisation
# ─────────────────────────────────────────────────────────────────────────────
def _column_spans(tris: np.ndarray, axis: int, cell_in: float):
    """Rasterise ``tris`` into columns along ``axis`` and pair up solid spans.

    A ray along ``axis`` crosses the surface an even number of times for a closed
    solid, so sorting a column's crossings and taking them two at a time gives
    the solid spans. Triangles parallel to the ray -- the walls of a hole --
    contribute no crossing and drop out on their own, which is why a through-hole
    appears as a column with no solid instead of needing a special case.

    Returns ``(n_u, n_v, cell, cols, enter, exit_)``.
    """
    u_ax, v_ax = [a for a in (0, 1, 2) if a != axis]
    pts = tris.reshape(-1, 3)
    lo, hi = pts.min(axis=0), pts.max(axis=0)
    span_u = max(hi[u_ax] - lo[u_ax], 1e-9)
    span_v = max(hi[v_ax] - lo[v_ax], 1e-9)

    cell = cell_in
    while (span_u / cell + 2) * (span_v / cell + 2) > MAX_GRID_CELLS:
        cell *= 1.5

    n_u = max(int(math.ceil(span_u / cell)), 1)
    n_v = max(int(math.ceil(span_v / cell)), 1)
    # Sample at cell centres, half a cell in from the bounding box, so a face
    # lying exactly on the boundary is never sampled edge-on.
    origin_u = lo[u_ax] + cell / 2.0
    origin_v = lo[v_ax] + cell / 2.0
    empty = (np.zeros(0, dtype=np.int64), np.zeros(0), np.zeros(0))

    p0, p1, p2 = tris[:, 0, :], tris[:, 1, :], tris[:, 2, :]
    e1_u, e1_v = p1[:, u_ax] - p0[:, u_ax], p1[:, v_ax] - p0[:, v_ax]
    e2_u, e2_v = p2[:, u_ax] - p0[:, u_ax], p2[:, v_ax] - p0[:, v_ax]
    det = e1_u * e2_v - e1_v * e2_u

    tu = tris[:, :, u_ax]
    tv = tris[:, :, v_ax]
    lo_iu = np.clip(np.ceil((tu.min(axis=1) - origin_u) / cell), 0, n_u - 1).astype(np.int64)
    hi_iu = np.clip(np.floor((tu.max(axis=1) - origin_u) / cell), 0, n_u - 1).astype(np.int64)
    lo_iv = np.clip(np.ceil((tv.min(axis=1) - origin_v) / cell), 0, n_v - 1).astype(np.int64)
    hi_iv = np.clip(np.floor((tv.max(axis=1) - origin_v) / cell), 0, n_v - 1).astype(np.int64)

    n_cu = np.maximum(hi_iu - lo_iu + 1, 0)
    n_cv = np.maximum(hi_iv - lo_iv + 1, 0)
    # |det| near zero: triangle is parallel to the ray, no crossing to record.
    live = (np.abs(det) > 1e-12) & (n_cu > 0) & (n_cv > 0)
    if not live.any():
        return (n_u, n_v, cell, *empty)

    idx_all = np.flatnonzero(live)
    pairs_all = (n_cu[idx_all] * n_cv[idx_all]).astype(np.int64)

    col_chunks: list[np.ndarray] = []
    depth_chunks: list[np.ndarray] = []

    # Walk the triangles in chunks of bounded total (triangle, cell) pairs. One
    # coarse triangle covering a whole plate face can be worth 10^5 pairs on its
    # own, so the chunk boundary is by pair count, not triangle count.
    start = 0
    cumulative = np.cumsum(pairs_all)
    while start < len(idx_all):
        budget = cumulative[start - 1] if start else 0
        stop = int(np.searchsorted(cumulative, budget + RASTER_CHUNK_PAIRS, side="right"))
        stop = max(stop, start + 1)
        idx = idx_all[start:stop]
        start = stop

        pairs = (n_cu[idx] * n_cv[idx]).astype(np.int64)
        tri_of_pair = np.repeat(idx, pairs)
        # Local pair number within each triangle's own bbox, via a ramp built
        # from the per-triangle offsets.
        offsets = np.concatenate(([0], np.cumsum(pairs)[:-1]))
        local = np.arange(len(tri_of_pair), dtype=np.int64) - np.repeat(offsets, pairs)
        nv_of_pair = np.repeat(n_cv[idx], pairs)
        iu = np.repeat(lo_iu[idx], pairs) + local // nv_of_pair
        iv = np.repeat(lo_iv[idx], pairs) + local % nv_of_pair

        uu = origin_u + iu * cell
        vv = origin_v + iv * cell
        du = uu - p0[tri_of_pair, u_ax]
        dv = vv - p0[tri_of_pair, v_ax]
        inv = 1.0 / det[tri_of_pair]
        b1 = (du * e2_v[tri_of_pair] - dv * e2_u[tri_of_pair]) * inv
        b2 = (dv * e1_u[tri_of_pair] - du * e1_v[tri_of_pair]) * inv
        inside = (b1 >= -1e-9) & (b2 >= -1e-9) & (b1 + b2 <= 1.0 + 1e-9)
        if not inside.any():
            continue

        t_in = tri_of_pair[inside]
        b1i, b2i = b1[inside], b2[inside]
        w = (
            (1.0 - b1i - b2i) * tris[t_in, 0, axis]
            + b1i * tris[t_in, 1, axis]
            + b2i * tris[t_in, 2, axis]
        )
        col_chunks.append(iu[inside] * n_v + iv[inside])
        depth_chunks.append(w)

    if not col_chunks:
        return (n_u, n_v, cell, *empty)

    col_ids = np.concatenate(col_chunks)
    depths = np.concatenate(depth_chunks)

    # Pair crossings into spans, column by column, vectorised: sort by
    # (column, depth) then take alternate entries as enter/exit.
    order = np.lexsort((depths, col_ids))
    c, d = col_ids[order], depths[order]
    run_start = np.concatenate(([0], np.flatnonzero(np.diff(c)) + 1))
    run_len = np.diff(np.concatenate((run_start, [len(c)])))
    # Position within each column's run.
    pos = np.arange(len(c), dtype=np.int64) - np.repeat(run_start, run_len)
    # Drop the last crossing of any odd-length run: the ray clipped a shared
    # edge, and guessing which crossing is spurious is worse than losing one.
    keep = ~((run_len % 2 == 1)[np.repeat(np.arange(len(run_len)), run_len)] & (pos == np.repeat(run_len - 1, run_len)))
    c, d, pos = c[keep], d[keep], pos[keep]
    enters = pos % 2 == 0
    if enters.sum() != (~enters).sum():
        return (n_u, n_v, cell, *empty)
    return n_u, n_v, cell, c[enters], d[enters], d[~enters]


def _sparse_blobs(mask_flat: np.ndarray, n_u: int, n_v: int) -> list[np.ndarray]:
    """4-connected components of a boolean grid, working only on its True cells.

    Labelling the whole grid would be the obvious approach and is far too slow in
    pure Python. Through-holes and pockets are a small fraction of the cells, so
    union-find over just those cells and their neighbours is orders of magnitude
    cheaper -- and keeps this module free of scipy.
    """
    cells = np.flatnonzero(mask_flat)
    if len(cells) == 0:
        return []
    # Map cell id -> dense index, so union-find runs over the True cells only.
    rank = np.full(n_u * n_v, -1, dtype=np.int64)
    rank[cells] = np.arange(len(cells))

    iu, iv = cells // n_v, cells % n_v
    edges_a: list[np.ndarray] = []
    edges_b: list[np.ndarray] = []
    # Right neighbour (v+1) and down neighbour (u+1); the mirrored pairs are
    # implied by union-find being symmetric.
    right = np.flatnonzero(iv + 1 < n_v)
    if len(right):
        nb = rank[cells[right] + 1]
        ok = nb >= 0
        edges_a.append(rank[cells[right][ok]])
        edges_b.append(nb[ok])
    down = np.flatnonzero(iu + 1 < n_u)
    if len(down):
        nb = rank[cells[down] + n_v]
        ok = nb >= 0
        edges_a.append(rank[cells[down][ok]])
        edges_b.append(nb[ok])

    if edges_a:
        roots = _union_find(len(cells), np.concatenate(edges_a), np.concatenate(edges_b))
    else:
        roots = np.arange(len(cells), dtype=np.int64)

    order = np.argsort(roots, kind="stable")
    sorted_roots = roots[order]
    starts = np.concatenate(([0], np.flatnonzero(np.diff(sorted_roots)) + 1))
    stops = np.concatenate((starts[1:], [len(sorted_roots)]))
    return [cells[order[s:e]] for s, e in zip(starts, stops)]


def _watertight_ratio(tris: np.ndarray) -> float:
    """Fraction of edges shared by exactly two triangles.

    Near 1.0 means a closed solid whose measurements can be trusted; a low value
    flags a mesh whose numbers are suspect, and lands in ``notes``.
    """
    if len(tris) == 0:
        return 0.0
    vid = _weld(tris)
    edges = np.sort(
        np.concatenate([vid[:, [0, 1]], vid[:, [1, 2]], vid[:, [2, 0]]], axis=0), axis=1
    )
    _, counts = np.unique(edges, axis=0, return_counts=True)
    if not len(counts):
        return 0.0
    return float(np.count_nonzero(counts == 2) / len(counts))


# ─────────────────────────────────────────────────────────────────────────────
# Per-body analysis
# ─────────────────────────────────────────────────────────────────────────────
def analyze_body(
    tris: np.ndarray,
    stack_axis: int = 1,
    cell_in: float = DEFAULT_CELL_IN,
    source: str = "",
    body_index: int = 0,
    label_hint: str = "",
    units: str = "in",
    cross_axis: bool = True,
) -> BodyGeometry:
    """Measure one solid body. ``tris`` must already be in inches."""
    g = BodyGeometry(
        source=source,
        body_index=body_index,
        label_hint=label_hint,
        n_triangles=int(len(tris)),
        units=units,
        stack_axis=stack_axis,
    )
    if len(tris) == 0:
        g.notes.append("empty body")
        return g

    pts = tris.reshape(-1, 3)
    lo, hi = pts.min(axis=0), pts.max(axis=0)
    ext = hi - lo
    ctr = (lo + hi) / 2.0
    g.ext_x, g.ext_y, g.ext_z = (float(v) for v in ext)
    g.center_x, g.center_y, g.center_z = (float(v) for v in ctr)
    g.stack_center = float(ctr[stack_axis])
    g.stack_min, g.stack_max = float(lo[stack_axis]), float(hi[stack_axis])
    g.thickness_in = float(ext[stack_axis])
    plan = sorted(float(ext[a]) for a in (0, 1, 2) if a != stack_axis)
    g.plan_short_in, g.plan_long_in = plan[0], plan[1]
    g.bbox_volume_cuin = float(ext[0] * ext[1] * ext[2])

    p0, p1, p2 = tris[:, 0, :], tris[:, 1, :], tris[:, 2, :]
    cross = np.cross(p1 - p0, p2 - p0)
    norm = np.linalg.norm(cross, axis=1)
    tri_area = 0.5 * norm
    g.surface_area_in2 = float(tri_area.sum())
    # Divergence theorem: signed tetrahedron volumes about the origin sum to the
    # enclosed volume. Facet winding sets the sign, so take the magnitude.
    g.volume_cuin = float(abs(np.einsum("ij,ij->", p0, cross) / 6.0))
    if g.bbox_volume_cuin > 0:
        g.solid_fill_pct = float(100.0 * g.volume_cuin / g.bbox_volume_cuin)

    g.watertight_ratio = _watertight_ratio(tris)
    if g.watertight_ratio < 0.98:
        g.notes.append(f"mesh not closed (edge match {g.watertight_ratio:.2f})")

    # Planar area lying exactly on the two stack-axis faces.
    n_s = cross[:, stack_axis]
    flat = np.abs(n_s) > 0.999 * norm + 1e-30
    face_pos = tris[:, :, stack_axis].mean(axis=1)
    plane_tol = max(1e-4, 0.002 * max(g.thickness_in, 1e-6))
    g.top_face_in2 = float(tri_area[flat & (np.abs(face_pos - g.stack_max) < plane_tol)].sum())
    g.bottom_face_in2 = float(tri_area[flat & (np.abs(face_pos - g.stack_min) < plane_tol)].sum())

    # ---- plan view ---------------------------------------------------------
    n_u, n_v, cell, cols, enter, exit_ = _column_spans(tris, stack_axis, cell_in)
    if len(cols) == 0:
        g.notes.append("plan-view raster empty; only bounding box is reliable")
        return g

    cell_area = cell * cell
    n_cols = n_u * n_v
    occupied = np.zeros(n_cols, dtype=bool)
    occupied[cols] = True
    g.n_solid_spans_max = int(np.bincount(cols, minlength=n_cols).max())

    top_of = np.full(n_cols, -np.inf)
    bot_of = np.full(n_cols, np.inf)
    np.maximum.at(top_of, cols, exit_)
    np.minimum.at(bot_of, cols, enter)
    solid_depth = np.zeros(n_cols)
    np.add.at(solid_depth, cols, exit_ - enter)

    g.solid_section_in2 = float(occupied.sum() * cell_area)

    # Footprint = the outline, including whatever it encloses. Components of the
    # empty region that never touch the border are interior voids, i.e.
    # through-holes, and belong inside the outline.
    empty_flat = ~occupied
    border = np.zeros((n_u, n_v), dtype=bool)
    border[0, :] = border[-1, :] = True
    border[:, 0] = border[:, -1] = True
    border_flat = border.reshape(-1)
    interior_void = np.zeros(n_cols, dtype=bool)
    void_blobs: list[np.ndarray] = []
    for blob in _sparse_blobs(empty_flat, n_u, n_v):
        if border_flat[blob].any():
            continue
        interior_void[blob] = True
        void_blobs.append(blob)

    footprint_mask = occupied | interior_void
    g.footprint_in2 = float(footprint_mask.sum() * cell_area)
    if g.plan_short_in * g.plan_long_in > 0:
        g.outline_fill_pct = float(
            100.0 * g.footprint_in2 / (g.plan_short_in * g.plan_long_in)
        )

    depth_tol = max(1.5 * cell, 0.02 * max(g.thickness_in, 1e-6))
    full = occupied & (solid_depth >= g.thickness_in - depth_tol)
    g.full_thickness_in2 = float(full.sum() * cell_area)
    if g.footprint_in2 > 0:
        g.full_thickness_pct = float(100.0 * g.full_thickness_in2 / g.footprint_in2)

    # ---- pockets and counterbores -----------------------------------------
    surf_tol = max(1.5 * cell, 0.01 * max(g.thickness_in, 1e-6))
    top_short = occupied & (top_of < g.stack_max - surf_tol)
    bot_short = occupied & (bot_of > g.stack_min + surf_tol)

    g.pocket_top_in2 = float(top_short.sum() * cell_area)
    g.pocket_bottom_in2 = float(bot_short.sum() * cell_area)
    if top_short.any():
        g.pocket_top_max_depth = float((g.stack_max - top_of[top_short]).max())
    if bot_short.any():
        g.pocket_bottom_max_depth = float((bot_of[bot_short] - g.stack_min).max())

    holes: list[HoleGroup] = []
    bore_sizes: list[tuple[float, float]] = []
    for mask, ref, sign in ((top_short, g.stack_max, -1.0), (bot_short, g.stack_min, 1.0)):
        regions = _sparse_blobs(mask, n_u, n_v)
        if sign < 0:
            g.n_pocket_regions_top = len(regions)
        else:
            g.n_pocket_regions_bottom = len(regions)
        for blob in regions:
            area = float(len(blob) * cell_area)
            dia = 2.0 * math.sqrt(area / math.pi)
            depth = (
                float((ref - top_of[blob]).max())
                if sign < 0
                else float((bot_of[blob] - ref).max())
            )
            # Round-ish, no wider than 3", with real depth: that is a bore, not a
            # milled cavity. Cavities are already carried by pocket_*_in2.
            if dia <= 3.0 and depth > 0.02 and dia >= 0.5 * cell:
                bore_sizes.append((dia, depth))
    g.n_counterbores = len(bore_sizes)

    thru_sizes = []
    for blob in void_blobs:
        area = float(len(blob) * cell_area)
        dia = 2.0 * math.sqrt(area / math.pi)
        if dia >= 0.5 * cell:
            thru_sizes.append(dia)
    g.n_thru_holes = len(thru_sizes)

    def _bucket(sizes, depths, axis_name, through, frac=64.0):
        # Bucket to 1/64" so one physical hole size stays one group despite
        # rasterisation noise.
        buckets: dict[float, list[float]] = {}
        for d, dep in zip(sizes, depths):
            buckets.setdefault(round(d * frac) / frac, []).append(dep)
        for key, deps in sorted(buckets.items(), reverse=True):
            holes.append(
                HoleGroup(
                    axis=axis_name,
                    diameter_in=round(key, 4),
                    count=len(deps),
                    through=through,
                    depth_in=round(max(deps), 4),
                )
            )

    _bucket(thru_sizes, [g.thickness_in] * len(thru_sizes), "thickness", True)
    _bucket([d for d, _ in bore_sizes], [dp for _, dp in bore_sizes], "thickness", False)

    # ---- cross-axis holes --------------------------------------------------
    # Waterlines and side-action clearance run across the plate rather than
    # through its faces, so they are invisible to the pass above. The same
    # rasteriser aimed along the two in-plane axes finds them.
    if cross_axis:
        for axis in (a for a in (0, 1, 2) if a != stack_axis):
            try:
                a_u, a_v, a_cell, a_cols, _, _ = _column_spans(
                    tris, axis, max(cell_in, 0.06)
                )
            except Exception as exc:  # one bad body must not sink the run
                g.notes.append(f"cross-axis scan failed on axis {axis}: {exc}")
                continue
            if len(a_cols) == 0:
                continue
            occ = np.zeros(a_u * a_v, dtype=bool)
            occ[a_cols] = True
            b = np.zeros((a_u, a_v), dtype=bool)
            b[0, :] = b[-1, :] = True
            b[:, 0] = b[:, -1] = True
            b_flat = b.reshape(-1)
            sizes = []
            for blob in _sparse_blobs(~occ, a_u, a_v):
                if b_flat[blob].any():
                    continue
                dia = 2.0 * math.sqrt((len(blob) * a_cell * a_cell) / math.pi)
                if dia >= 0.5 * a_cell:
                    sizes.append(dia)
            if sizes:
                span = float(ext[axis])
                _bucket(sizes, [span] * len(sizes), "cross", True, frac=32.0)
                g.n_cross_holes += len(sizes)

    g.holes = sorted(holes, key=lambda h: (-h.diameter_in, h.axis))
    if g.holes:
        g.max_bore_dia = float(max(h.diameter_in for h in g.holes))
        g.hole_signature = "|".join(
            f"{h.diameter_in:.3f}{'' if h.axis == 'thickness' else 'c'}"
            f"{'' if h.through else 'b'}x{h.count}"
            for h in g.holes[:12]
        )
    return g


# ─────────────────────────────────────────────────────────────────────────────
# Stack-axis inference
# ─────────────────────────────────────────────────────────────────────────────
def detect_stack_axis(extents: np.ndarray) -> int:
    """Pick the stack axis from a set of body bounding-box extents.

    ``extents`` is ``(n_bodies, 3)``. Plates fill the assembly in the two
    in-plane axes and only stack up along the third, so the stack axis is the one
    where a typical body's extent is the smallest fraction of the assembly's.
    Measured this way it comes out as Y for every job in this workspace -- but it
    is measured, not assumed, so a re-oriented export cannot silently break it.
    """
    if extents.ndim != 2 or len(extents) == 0:
        return 1
    assembly = extents.max(axis=0)
    assembly[assembly <= 0] = 1e-9
    return int(np.argmin(np.median(extents, axis=0) / assembly))


# ─────────────────────────────────────────────────────────────────────────────
# File-level entry points
# ─────────────────────────────────────────────────────────────────────────────
def _solid_name(path: Path) -> str:
    """The name in an STL's 80-byte header, kept only as an untrusted hint.

    The macro writes the plate role it *currently* believes into both the file
    name and this header. Feeding that back as evidence would make the naming
    circular, which is how a wrong name survives a re-run -- so it travels as
    ``label_hint`` and the prompt builder does not show it to the model.
    """
    try:
        head = path.read_bytes()[:80]
    except OSError:
        return ""
    if head[:5] != b"solid":
        return ""
    return head[5:].decode("ascii", errors="replace").strip().strip("\x00")


def read_scaled(path: str | Path) -> tuple[np.ndarray, str]:
    """Read an STL and convert it to inches. Returns ``(triangles, unit_name)``."""
    tris = read_stl(path)
    if len(tris) == 0:
        return tris, "in"
    unit_name, scale = detect_units(tris)
    return (tris * scale if scale != 1.0 else tris), unit_name


def quick_measure(tris: np.ndarray, stack_axis: int = 1) -> dict:
    """Bounding box, volume and centre only -- no rasterisation.

    This is the cheap half of :func:`analyze_body`, split out so a caller can
    match a mesh to its CAD row (volume is the join key) and decide whether the
    expensive feature pass is worth running on it at all. On a job with 130 parts
    that is the difference between measuring ten plates and measuring everything.
    """
    if len(tris) == 0:
        return {"volume_cuin": 0.0, "ext": (0.0, 0.0, 0.0), "center": (0.0, 0.0, 0.0)}
    pts = tris.reshape(-1, 3)
    lo, hi = pts.min(axis=0), pts.max(axis=0)
    p0, p1, p2 = tris[:, 0, :], tris[:, 1, :], tris[:, 2, :]
    cross = np.cross(p1 - p0, p2 - p0)
    ext = hi - lo
    plan = sorted(float(ext[a]) for a in (0, 1, 2) if a != stack_axis)
    return {
        "volume_cuin": float(abs(np.einsum("ij,ij->", p0, cross) / 6.0)),
        "ext": tuple(float(v) for v in ext),
        "center": tuple(float(v) for v in (lo + hi) / 2.0),
        "thickness_in": float(ext[stack_axis]),
        "plan_short_in": plan[0],
        "plan_long_in": plan[1],
        "n_triangles": int(len(tris)),
    }


def analyze_stl(
    path: str | Path,
    stack_axis: int | None = None,
    cell_in: float = DEFAULT_CELL_IN,
    split: bool = False,
    cross_axis: bool = True,
    max_bodies: int = 400,
) -> list[BodyGeometry]:
    """Analyse an STL file and return one :class:`BodyGeometry` per body.

    ``split=False`` (the default) suits the per-plate files in a job's ``stl\\``
    folder, which hold one part each. ``split=True`` runs connected components
    first -- only useful on a mesh whose parts genuinely do not touch.
    """
    p = Path(path)
    tris = read_stl(p)
    if len(tris) == 0:
        return []

    unit_name, scale = detect_units(tris)
    if scale != 1.0:
        tris = tris * scale
    hint = _solid_name(p) or p.stem

    groups = split_bodies(tris) if split else [np.arange(len(tris))]
    if not groups:
        return []
    if stack_axis is None:
        ext = np.array(
            [np.ptp(tris[i].reshape(-1, 3), axis=0) for i in groups], dtype=np.float64
        )
        stack_axis = detect_stack_axis(ext)

    out: list[BodyGeometry] = []
    for i, idx in enumerate(groups[:max_bodies]):
        out.append(
            analyze_body(
                tris[idx],
                stack_axis=stack_axis,
                cell_in=cell_in,
                source=p.name,
                body_index=i,
                label_hint=hint if len(groups) == 1 else "",
                units=unit_name,
                cross_axis=cross_axis,
            )
        )
    return out


def analyze_stl_folder(
    folder: str | Path,
    cell_in: float = DEFAULT_CELL_IN,
    cross_axis: bool = True,
    stack_axis: int | None = None,
) -> list[BodyGeometry]:
    """Analyse a job's ``stl\\`` folder -- one steel part per file.

    The stack axis is decided once across every file, so all bodies are measured
    in a single consistent frame and their stack positions are comparable.
    """
    files = sorted(Path(folder).glob("*.[sS][tT][lL]"))
    if not files:
        return []

    meshes: list[tuple[Path, np.ndarray, str]] = []
    for p in files:
        tris = read_stl(p)
        if len(tris) == 0:
            continue
        unit_name, scale = detect_units(tris)
        meshes.append((p, tris * scale if scale != 1.0 else tris, unit_name))
    if not meshes:
        return []

    if stack_axis is None:
        stack_axis = detect_stack_axis(
            np.array([np.ptp(t.reshape(-1, 3), axis=0) for _, t, _ in meshes])
        )

    out: list[BodyGeometry] = []
    for i, (p, tris, unit_name) in enumerate(meshes):
        out.append(
            analyze_body(
                tris,
                stack_axis=stack_axis,
                cell_in=cell_in,
                source=p.name,
                body_index=i,
                label_hint=p.stem,
                units=unit_name,
                cross_axis=cross_axis,
            )
        )
    return out
