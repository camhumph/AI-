"""Seed a real demo job (T001015) so the app is fully functional out of the
box, without needing your actual C:\\CMS_Local_Workspace mounted yet.

T001015 is reconstructed from the existing correction ledger already in this
repo (geometry_classifier/outputs/XT_Export_CAD_Dimensions_CORRECT_ME.csv,
which is T001015 data) and the T001015 rendered views already checked into
datasets/cms_molds/label_queue/images/. It is then classified with the
*current* geometry_classifier rules so the demo showcases the latch-lock /
bottom_ejector_plate fixes.

Run once: python3 seed_demo_data.py
"""
import csv
import shutil
import sys
from pathlib import Path

BACKEND_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BACKEND_DIR))

from app import config, jobs  # noqa: E402

REPO_ROOT = BACKEND_DIR.parent.parent
CORRECTION_LEDGER = (
    REPO_ROOT / "geometry_classifier" / "outputs"
    / "XT_Export_CAD_Dimensions_CORRECT_ME.csv"
)
IMAGES_DIR = REPO_ROOT / "datasets" / "cms_molds" / "label_queue" / "images"


def build_raw_csv(job_dir: Path):
    if not CORRECTION_LEDGER.exists():
        print(f"WARNING: {CORRECTION_LEDGER} not found, skipping T001015 raw CSV")
        return False
    with CORRECTION_LEDGER.open("r", newline="", encoding="utf-8") as f:
        rows = list(csv.DictReader(f))
    fields = ["Index", "Component", "Thickness", "Width", "Length", "BBoxVolume_cuin",
              "CenterX", "CenterY", "CenterZ"]
    out_path = job_dir / "XT_Export_CAD_Dimensions.csv"
    with out_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fields)
        writer.writeheader()
        for r in rows:
            writer.writerow(
                {
                    "Index": r["Index"],
                    "Component": r["Component"],
                    "Thickness": r["Thickness"],
                    "Width": r["Width"],
                    "Length": r["Length"],
                    "BBoxVolume_cuin": "",
                    "CenterX": r["CenterX"],
                    "CenterY": r["CenterY"],
                    "CenterZ": r["CenterZ"],
                }
            )
    print(f"Wrote {out_path} ({len(rows)} rows)")
    return True


def copy_images(job_dir: Path):
    images_out = job_dir / "images"
    images_out.mkdir(exist_ok=True)
    if not IMAGES_DIR.exists():
        print(f"WARNING: {IMAGES_DIR} not found, skipping images")
        return
    count = 0
    for img in sorted(IMAGES_DIR.glob("t001015_asm_FULL_*.jpg")):
        shutil.copy2(img, images_out / img.name)
        count += 1
    print(f"Copied {count} images to {images_out}")


def main():
    job_dir = config.JOBS_ROOT / "T001015"
    job_dir.mkdir(parents=True, exist_ok=True)
    jobs.create_job(
        "T001015",
        display_name="T001015 -- Plate-Sequenced Latch-Lock Base",
        customer="Demo / Training Data",
    )
    has_raw = build_raw_csv(job_dir)
    copy_images(job_dir)

    if has_raw:
        print("Running classifier (rules-only)...")
        result = jobs.classify_job("T001015", mode="rules")
        print(f"Classified {len(result['parts'])} parts.")
        print(f"Sequenced latch-lock base: {result['job_analysis'].get('sequenced_latch_lock_base')}")

    print("Done. Demo job 'T001015' is ready.")


if __name__ == "__main__":
    main()
