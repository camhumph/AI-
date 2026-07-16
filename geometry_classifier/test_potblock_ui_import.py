"""Regression: pot-block XT must not become A/B/rails; import hoists STL/ISO."""

from __future__ import annotations

import csv
import shutil
import tempfile
from pathlib import Path


def _make_8222_rows():
    """Synthetic XT rows matching the 822200009 pot-block pattern from the UI."""
    # T, W, L roughly matching the user's Parts table
    dims = [
        (6.75, 6.852, 15.375),   # holder mislabeled A
        (6.75, 6.75, 15.375),    # holder mislabeled B
        (1.375, 15.875, 20),     # TCP
        (1.375, 15.875, 20),     # BCP
        (0.25, 11.75, 15.375),   # insulation mislabeled rail
        (0.25, 11.75, 15.375),
        (0.25, 6.72, 15.875),
        (0.25, 6.625, 8.505),
        (5.0, 6.0, 6.845),       # pot-like
        (5.0, 6.0, 6.75),
        (1.24, 1.24, 8.25),
        (1.375, 1.376, 1.376),
    ]
    rows = []
    for i, (t, w, l) in enumerate(dims, start=1):
        rows.append(
            {
                "i": i,
                "name": f"822200009__asm_objects_{i}-1",
                "t": t,
                "w": w,
                "l": l,
                "v": t * w * l,
                "x": 0.0,
                "y": float(i),
                "z": 0.0,
            }
        )
    return rows


def test_pot_block_guard_skips_ab_rails():
    from geometry_classifier.qwen_classify_xt_csv import (
        classify_geometry,
        looks_like_pot_block_geometry,
    )

    rows = _make_8222_rows()
    assert looks_like_pot_block_geometry(rows) is True
    result = classify_geometry(rows)
    assert result["job_analysis"].get("base_type") == "bms"
    roles = {c["role"] for c in result["classifications"]}
    assert "a_plate" not in roles
    assert "b_plate" not in roles
    assert "rail" not in roles and "rail_1" not in roles and "rail_2" not in roles


def test_import_hoists_stl_and_iso(tmp_path: Path | None = None):
    from webapp.backend.app import jobs

    root = Path(tempfile.mkdtemp(prefix="cms_hoist_")) if tmp_path is None else tmp_path
    src = root / "C18699"
    src.mkdir()
    (src / "base").mkdir()
    (src / "foo ISO.jpg").write_bytes(b"jpeg")
    (src / "foo BACK ISO.jpg").write_bytes(b"jpeg2")
    (src / "foo.stl").write_bytes(b"solid")
    (src / "base" / "foo.stl").write_bytes(b"solid")
    (src / "XT_Export_CAD_Dimensions.csv").write_text(
        "Index,Component,Qty,Thickness,Width,Length,BBoxVolume_cuin,Mass_or_Vol,CenterX,CenterY,CenterZ\n"
        "1,a,1,1.375,15.875,20,400,1,0,0,0\n"
        "2,b,1,1.375,15.875,20,400,1,0,1,0\n"
        "3,c,1,6.75,6.85,15.375,700,1,0,2,0\n"
        "4,d,1,6.75,6.75,15.375,690,1,0,3,0\n"
        "5,e,1,0.25,11.75,15.375,45,1,0,4,0\n"
        "6,f,1,0.25,11.75,15.375,45,1,0,5,0\n"
        "7,g,1,5,6,6.75,200,1,0,6,0\n"
        "8,h,1,5,6,6.8,204,1,0,7,0\n",
        encoding="utf-8",
    )
    (src / "Purchased Components Quote.csv").write_text(
        "Component,Vendor,PartNumber,Description,QTY,UnitPrice,Extended\n"
        "Leader Pin,DME,5213GL,Top Leader Pins,4,12.5,50\n"
        "TOTAL,,,,, ,50\n",
        encoding="utf-8",
    )

    old_root = jobs.config.JOBS_ROOT
    try:
        jobs.config.JOBS_ROOT = root / "jobs"
        jobs.config.JOBS_ROOT.mkdir()
        job = jobs.import_from_folder(str(src))
        job_dir = jobs.config.JOBS_ROOT / job["job_id"]
        assert (job_dir / "models" / "foo.stl").exists()
        assert (job_dir / "images" / "foo ISO.jpg").exists()
        assert (job_dir / "images" / "foo BACK ISO.jpg").exists()
        assert job.get("base_type") == "bms" or jobs._read_meta(job_dir).get("base_type") == "bms"
        assert len(job["models"]) >= 1
        assert len(job["images"]) >= 2
    finally:
        jobs.config.JOBS_ROOT = old_root
        if tmp_path is None:
            shutil.rmtree(root, ignore_errors=True)


if __name__ == "__main__":
    test_pot_block_guard_skips_ab_rails()
    test_import_hoists_stl_and_iso()
    print("OK: pot-block + STL/ISO hoist regressions passed")
