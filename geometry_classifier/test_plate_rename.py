"""Regression tests for plate renames.

Run from the repo root:
    set PYTHONPATH=C:\\CMS_AI
    python geometry_classifier\\test_plate_rename.py

The invariant these protect is the important one: A RENAME CHANGES THE LABEL,
NEVER THE ROLE. `role` decides the quote row and the price, so if a rename ever
starts re-keying it, typing in a text box would silently move money between rows
of the grade block with nothing in the UI to show it.
"""
from __future__ import annotations

import json
import sys
import tempfile
from pathlib import Path


def _make_job(root: Path) -> Path:
    jd = root / "jobs" / "C99999"
    jd.mkdir(parents=True, exist_ok=True)
    (jd / "meta.json").write_text(
        json.dumps({"display_name": "C99999", "base_type": "standard"}), encoding="utf-8"
    )
    (jd / "classification.json").write_text(
        json.dumps(
            {
                "job_analysis": {"stack_axis": "Z"},
                "classifications": [
                    {"index": "3", "role": "a_plate", "confidence": "HIGH", "quote": True, "reason": "x"},
                    {"index": "5", "role": "b_plate", "confidence": "HIGH", "quote": True, "reason": "x"},
                    {"index": "7", "role": "rail", "confidence": "HIGH", "quote": True, "reason": "x"},
                    {"index": "8", "role": "rail", "confidence": "HIGH", "quote": True, "reason": "x"},
                ],
            }
        ),
        encoding="utf-8",
    )
    (jd / "classification.csv").write_text(
        "Index,Component,Role,Confidence,Quote,Thickness,Width,Length,CenterX,CenterY,CenterZ,Reason\n"
        "3,ASM-1/2223605_A-PLATE_1,a_plate,HIGH,TRUE,2.375,9.875,16.0,0,0,3,x\n"
        "5,ASM-1/2223605_B-PLATE_2,b_plate,HIGH,TRUE,3.875,9.875,16.0,0,0,-2,x\n"
        "7,ASM-1/RAIL-TOP_1,rail,HIGH,TRUE,1.375,2.0,16.0,4,0,-6,x\n"
        "8,ASM-1/RAIL-BOTTOM_1,rail,HIGH,TRUE,1.375,2.0,16.0,-4,0,-6,x\n",
        encoding="utf-8",
    )
    return jd


def test_rename_propagates_but_never_moves_the_role():
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        _make_job(root)
        import os

        os.environ["CMS_JOBS_ROOT"] = str(root / "jobs")
        os.environ["CMS_DATA_DIR"] = str(root)
        for m in [k for k in list(sys.modules) if k.startswith("webapp.backend.app")]:
            del sys.modules[m]
        from webapp.backend.app import jobs, plate_names, pricing, vba_bridge

        jd = jobs.job_dir("C99999")
        assert jd is not None, "job_dir should find the job"

        plate_names.set_overrides(jd, {"3": "Cavity Plate", "7": "Left Spacer"})

        job = jobs.get_job("C99999")
        by_idx = {p["index"]: p for p in job["parts"]}

        # 1. Label changed.
        assert by_idx["3"]["role_label"] == "Cavity Plate", by_idx["3"]
        assert by_idx["7"]["role_label"] == "Left Spacer", by_idx["7"]
        # 2. ROLE did NOT. This is the invariant.
        assert by_idx["3"]["role"] == "a_plate", "rename must not re-key the role"
        assert by_idx["7"]["role"] == "rail", "rename must not re-key the role"
        # 3. Untouched plates keep their defaults.
        assert by_idx["5"]["role_label"] == "B Plate", by_idx["5"]
        # 4. The original label is kept so the UI can explain the pricing.
        assert by_idx["3"]["role_label_original"] == "A Plate", by_idx["3"]

        # 5. The parts table path (pricing) shows the override, beating the CAD
        #    component name -- it computes `display` independently of get_job.
        q = pricing.build_quote_sheet(job)
        cls = {i["index"]: i for i in q["line_items"] if i.get("section") == "classified"}
        assert cls["3"]["component"] == "Cavity Plate", cls["3"]
        assert cls["3"]["role"] == "a_plate", "quote row must not move"
        # An un-renamed PLATE shows its resolved shop name, not the raw CAD path.
        # This assertion used to require "B-PLATE" -- it was pinning the bug where
        # the CAD path won and the table read "2223605_B-PLATE".
        assert cls["5"]["component"] == "B Plate", cls["5"]
        # Hardware is the opposite: the CAD name IS the part number, so it stays.
        hard = {i["index"]: i for i in q["line_items"] if i.get("role") == "leader_pin"}
        if hard:
            assert "LDR-PIN" in next(iter(hard.values()))["component"], hard

        # 6. The bridge file carries it, so the next macro run writes the new
        #    name into the quote sheet and steel sheet instead of reverting it.
        vba_bridge.write_bridge_files("C99999", q, job.get("job_analysis", {}))
        import csv

        bridge = root / "vba_bridge" / "C99999_part_names.csv"
        rows = list(csv.DictReader(bridge.open(encoding="utf-8")))
        rn = {r["Index"]: r["ResolvedName"] for r in rows}
        assert rn["3"] == "Cavity Plate", rn
        assert rn["7"] == "Left Spacer", rn
        # 7. Renaming rail 1 must not renumber rail 2. _resolved_name advances a
        #    per-role counter, so short-circuiting on the override would have
        #    made this "Rail 1".
        assert rn["8"] == "Rail 2", f"rail numbering broke: {rn}"

        # 8. Overrides survive a re-classify, which overwrites classification.*
        #    wholesale. This is why the store lives in meta.json.
        cl = json.loads((jd / "classification.json").read_text(encoding="utf-8"))
        (jd / "classification.json").write_text(json.dumps(cl), encoding="utf-8")
        again = jobs.get_job("C99999")
        assert {p["index"]: p["role_label"] for p in again["parts"]}["3"] == "Cavity Plate"

        # 9. Clearing restores the role default.
        plate_names.set_overrides(jd, {"3": ""})
        after = jobs.get_job("C99999")
        assert {p["index"]: p["role_label"] for p in after["parts"]}["3"] == "A Plate"


def test_name_sanitising():
    from webapp.backend.app.plate_names import clean_name, MAX_NAME_LEN

    # Excel formula injection: these names get written into cells, and a leading
    # = + - @ is executed when the workbook opens.
    for bad in ("=cmd|calc", "+SUM(A1)", "@evil", "-1+1"):
        assert not clean_name(bad).startswith(("=", "+", "@", "-")), bad
    # The macro strips double quotes when writing the steel sheet, so a name
    # containing one would not match on the way back.
    assert '"' not in clean_name('A "X" Plate')
    assert clean_name("  Spaced   Out  ") == "Spaced Out"
    assert clean_name("\t\tTabbed") == "Tabbed"
    assert clean_name(None) == ""
    assert len(clean_name("x" * 200)) == MAX_NAME_LEN


def test_sheet_rewrite_degrades_without_excel():
    """No Excel/pywin32 must be a delay, not a crash and not a lost rename."""
    with tempfile.TemporaryDirectory() as td:
        root = Path(td)
        jd = _make_job(root)
        from webapp.backend.app import sheet_rename

        r = sheet_rename.rewrite_names(jd, [("A Plate", "Cavity Plate")])
        assert r["changed"] == 0
        assert r["skipped_reason"], "must explain why nothing was written"
        assert sheet_rename.rewrite_names(jd, [])["skipped_reason"] == "nothing to rename"


def test_role_keyed_rename_for_steel_rows():
    """A steel row has no CAD index, so its rename must key on the ROLE.

    Steel rows are numbered S1, S2, ... -- table positions, not part identities.
    Storing a rename under "S1" put it somewhere nothing reads, which is why
    typing a name and pressing Enter appeared to do nothing at all.
    """
    from webapp.backend.app import plate_names as pn

    rows = [
        {"index": "3", "role": "a_plate", "role_label": "A Plate"},
        {"index": "5", "role": "b_plate", "role_label": "B Plate"},
    ]
    ov = {pn.role_key("b_plate"): "Core Plate"}
    pn.apply_to_rows(rows, ov)
    by = {r["index"]: r for r in rows}
    assert by["5"]["role_label"] == "Core Plate", by["5"]
    assert by["5"]["role"] == "b_plate", "role must not move"
    assert by["3"]["role_label"] == "A Plate", by["3"]

    # A CAD-index rename is more specific and must win over a role-wide one.
    ov2 = {pn.role_key("b_plate"): "Core Plate", "5": "This One Only"}
    rows2 = [{"index": "5", "role": "b_plate", "role_label": "B Plate"}]
    pn.apply_to_rows(rows2, ov2)
    assert rows2[0]["role_label"] == "This One Only", rows2[0]

    # Aliases for the sheet-dimension join must resolve role keys too, or a
    # renamed plate loses its stock dimensions and its price changes.
    import tempfile, json as _json
    from pathlib import Path as _P

    with tempfile.TemporaryDirectory() as td:
        jd = _P(td)
        (jd / "meta.json").write_text(
            _json.dumps({"plate_name_overrides": {pn.role_key("b_plate"): "Core Plate"}}),
            encoding="utf-8",
        )
        al = pn.sheet_name_aliases(jd, {})
        assert al.get("b_plate") == ["core plate"], al


def test_two_mold_bases_are_told_apart():
    """C18616 holds two bases; the _2 suffix says which, and hardware must not."""
    from webapp.backend.app import mold_groups as mg

    assert mg.group_of("A/2223605_B-PLATE_2-1") == 2
    assert mg.group_of("A/2223605_B-PLATE-1") is None
    # Part numbers that merely contain _<digit> must NOT read as a mold number.
    for hw in (
        "A/LDR-PIN_2OD-X-7-3-4_PCS-1",
        "A/LBB_2ID-X-5-7-8-2",
        "A/GEB_1ID-X-1-3-8OD_DME-3",
        "A/EJ_LDR_PIN_D-1-X-6-25_PCS-4",
    ):
        assert mg.group_of(hw) is None, hw

    two = [
        {"index": "1", "role": "b_plate", "Component": "A/2223605_B-PLATE_2-1"},
        {"index": "2", "role": "b_plate", "Component": "A/2223605_B-PLATE-1"},
    ]
    g = mg.detect_groups(two)
    assert mg.group_count(g) == 2, g
    assert g["1"] == 2 and g["2"] == 1, g

    # One base: no markers anywhere, so no "(Mold n)" noise.
    one = [
        {"index": "1", "role": "a_plate", "Component": "A/2223605_A-PLATE-1"},
        {"index": "2", "role": "rail", "Component": "A/RAIL-TOP-1"},
        {"index": "3", "role": "rail", "Component": "A/RAIL-BOTTOM-1"},
    ]
    assert mg.detect_groups(one) == {}, "single-base job must be left alone"


if __name__ == "__main__":
    test_name_sanitising()
    test_sheet_rewrite_degrades_without_excel()
    test_role_keyed_rename_for_steel_rows()
    test_two_mold_bases_are_told_apart()
    test_rename_propagates_but_never_moves_the_role()
    print("OK: plate rename regressions passed")
