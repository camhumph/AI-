"""Regression: standard mold bases must NOT be forced onto the BMS branch."""

from __future__ import annotations

import csv
import json
import tempfile
from pathlib import Path


def _write_xt(path: Path, plates: list[tuple[float, float, float]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(
            [
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
        )
        for i, (t, wdim, l) in enumerate(plates, start=1):
            w.writerow([i, f"part_{i}", 1, t, wdim, l, t * wdim * l, 1, 0, float(i), 0])


def test_standard_stack_not_pot_block():
    """5+ full-footprint plates = standard mold — never BMS geometry."""
    from webapp.backend.app import jobs

    root = Path(tempfile.mkdtemp(prefix="cms_std_stack_"))
    job = root / "C18001"
    job.mkdir()
    # Typical DME stack: TCP, A, B, Support, BCP — all ~same footprint
    _write_xt(
        job / "XT_Export_CAD_Dimensions.csv",
        [
            (1.875, 15.875, 23.75),  # TCP
            (2.375, 15.875, 23.75),  # A
            (2.875, 15.875, 23.75),  # B
            (1.875, 15.875, 23.75),  # Support
            (1.875, 15.875, 23.75),  # BCP
            (2.0, 2.0, 15.875),      # rail
            (2.0, 2.0, 15.875),      # rail
            (0.875, 15.875, 23.75),  # ejector
        ],
    )
    assert jobs._xt_looks_like_pot_block(job) is False


def test_pot_block_geometry_still_detected():
    from webapp.backend.app import jobs

    root = Path(tempfile.mkdtemp(prefix="cms_pot_stack_"))
    job = root / "C18699"
    job.mkdir()
    _write_xt(
        job / "XT_Export_CAD_Dimensions.csv",
        [
            (1.375, 15.875, 20.0),   # TCP clamp
            (1.375, 15.875, 20.0),   # BCP clamp
            (6.75, 6.85, 15.375),    # holder
            (6.75, 6.75, 15.375),    # holder
            (5.0, 6.0, 6.845),       # pot
            (5.0, 6.0, 6.75),        # pot
            (0.25, 11.75, 15.375),   # insulation
            (0.25, 11.75, 15.375),
        ],
    )
    assert jobs._xt_looks_like_pot_block(job) is True


def test_folder_bms_not_from_random_csv_text():
    """Standard Dynacast folder must not become BMS just because a CSV mentions SMED."""
    from webapp.backend.app import jobs

    root = Path(tempfile.mkdtemp(prefix="cms_std_folder_"))
    job = root / "Dynacast-C18050"
    job.mkdir()
    (job / "notes.csv").write_text("Component,Note\nX,mentions SMED in passing\n", encoding="utf-8")
    assert jobs._folder_looks_like_bms(job) is False

    bms = root / "BMS-851100029-C18603"
    bms.mkdir()
    assert jobs._folder_looks_like_bms(bms) is True


def test_folder_bms_from_macro_log_pot_line():
    from webapp.backend.app import jobs

    root = Path(tempfile.mkdtemp(prefix="cms_log_bms_"))
    job = root / "C18610"
    job.mkdir()
    (job / "CMS_Base_Export_Log.txt").write_text(
        "Base type: POT / HOLDER BLOCK\nDONE JOB C18610\n", encoding="utf-8"
    )
    assert jobs._folder_looks_like_bms(job) is True

    std = root / "C18060"
    std.mkdir()
    (std / "CMS_Base_Export_Log.txt").write_text(
        "Base type: STANDARD MOLD BASE\nBase type STANDARD from geometry (nFull=5)\nDONE JOB C18060\n",
        encoding="utf-8",
    )
    assert jobs._folder_looks_like_bms(std) is False


def test_poll_completion_accepts_done_log(tmp_path: Path | None = None):
    from webapp.backend.app import quote_pipeline

    root = Path(tempfile.mkdtemp(prefix="cms_poll_")) if tmp_path is None else tmp_path
    job_id = "C18070"
    local = root / "workspace" / job_id
    local.mkdir(parents=True)
    (local / "XT_Export_CAD_Dimensions.csv").write_text(
        "Index,Component,Qty,Thickness,Width,Length\n1,a,1,1,10,10\n", encoding="utf-8"
    )
    (local / "CMS_Base_Export_Log.txt").write_text(
        "Base type: STANDARD MOLD BASE\nDONE JOB C18070. Output folder: x\nTOTAL JOB TIME: 12s\n",
        encoding="utf-8",
    )

    status_dir = root / "status"
    status_dir.mkdir()
    old_status = quote_pipeline.STATUS_DIR
    quote_pipeline.STATUS_DIR = status_dir

    quote_pipeline.set_status(job_id, phase="running", message="quoting", job_id=job_id, c_number=job_id)

    orig_find = quote_pipeline.find_local_job_folder

    def _find(jid: str):
        if jid == job_id:
            return local
        return None

    quote_pipeline.find_local_job_folder = _find
    orig_sync = quote_pipeline.sync_completed_job

    def _sync(jid, folder, base_type="standard"):
        quote_pipeline.set_status(
            jid, phase="completed", message="synced", job_id=jid, local_folder=folder
        )
        return {"job_id": jid}

    quote_pipeline.sync_completed_job = _sync
    try:
        st = quote_pipeline.poll_completion(job_id)
        assert st.get("outputs_found") is True, st
        assert st.get("phase") == "completed", st
    finally:
        quote_pipeline.STATUS_DIR = old_status
        quote_pipeline.find_local_job_folder = orig_find
        quote_pipeline.sync_completed_job = orig_sync


def test_classifier_standard_stack_not_bms_guard():
    from geometry_classifier.qwen_classify_xt_csv import (
        classify_geometry,
        looks_like_pot_block_geometry,
    )

    rows = []
    dims = [
        (1.875, 15.875, 23.75),
        (2.375, 15.875, 23.75),
        (2.875, 15.875, 23.75),
        (1.875, 15.875, 23.75),
        (1.875, 15.875, 23.75),
        (2.0, 2.0, 15.875),
        (2.0, 2.0, 15.875),
    ]
    for i, (t, w, l) in enumerate(dims, start=1):
        rows.append(
            {
                "i": i,
                "name": f"plate_{i}",
                "t": t,
                "w": w,
                "l": l,
                "v": t * w * l,
                "x": 0.0,
                "y": float(i),
                "z": 0.0,
            }
        )
    assert looks_like_pot_block_geometry(rows) is False
    result = classify_geometry(rows)
    assert result["job_analysis"].get("base_type") != "bms"
    roles = {c["role"] for c in result["classifications"]}
    # Should assign real plate roles, not all hardware_other from BMS guard
    assert "a_plate" in roles or "b_plate" in roles or "top_clamp_plate" in roles


if __name__ == "__main__":
    test_standard_stack_not_pot_block()
    test_pot_block_geometry_still_detected()
    test_folder_bms_not_from_random_csv_text()
    test_folder_bms_from_macro_log_pot_line()
    test_poll_completion_accepts_done_log()
    test_classifier_standard_stack_not_bms_guard()
    print("OK: standard vs BMS branch regressions passed")
