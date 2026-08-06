"""Tidy a job folder: loose CAD into base\\, diagnostic CSVs rendered to PDF.

    python tidy_job_folder.py C18597                 # by job number
    python tidy_job_folder.py "C:\\path\\to\\job"      # by full path
    python tidy_job_folder.py C18597 --dry-run       # show, change nothing
    python tidy_job_folder.py C18597 --undo          # put the CAD files back
    python tidy_job_folder.py --all                  # every job in the workspace

The CSVs are NOT deleted. XT_Export_CAD_Dimensions.csv is the webapp's source of
truth and XT_Export_BOM_Match_Report.csv is read when pricing, so the PDFs are
written into a `diagnostics\\` subfolder and the CSVs stay put.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "webapp" / "backend"))

from app import job_housekeeping as hk  # noqa: E402

WORKSPACE = Path(r"C:\CMS_Local_Workspace")


def resolve(target: str) -> Path:
    p = Path(target)
    if p.is_dir():
        return p
    cand = WORKSPACE / target
    if cand.is_dir():
        return cand
    raise SystemExit(f"No such job folder: {target}")


def show(label: str, rep: dict) -> None:
    print(f"\n{label}")
    for key in ("moved", "written", "restored"):
        for item in rep.get(key, []) or []:
            print(f"   + {item}")
    for item in rep.get("skipped", []) or []:
        print(f"   . skipped: {item}")
    for item in rep.get("errors", []) or []:
        print(f"   ! ERROR: {item}")
    if rep.get("manifest"):
        print(f"   manifest: {rep['manifest']}")
    if rep.get("dir"):
        print(f"   output:   {rep['dir']}")


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("job", nargs="?", help="job number or folder path")
    ap.add_argument("--all", action="store_true", help="every job folder in the workspace")
    ap.add_argument("--dry-run", action="store_true", help="report only, change nothing")
    ap.add_argument("--undo", action="store_true", help="move the CAD files back to the job root")
    ap.add_argument("--cad-only", action="store_true")
    ap.add_argument("--pdf-only", action="store_true")
    args = ap.parse_args()

    if args.all:
        targets = [p for p in sorted(WORKSPACE.iterdir())
                   if p.is_dir() and not p.name.startswith(("AI_Bridge", "."))]
    elif args.job:
        targets = [resolve(args.job)]
    else:
        ap.print_help()
        return 2

    for job_dir in targets:
        print("=" * 72)
        print(job_dir)
        if args.undo:
            show("CAD restore", hk.undo_cad_move(job_dir))
            continue
        if not args.pdf_only:
            show("Loose CAD -> base\\", hk.tidy_cad_files(job_dir, dry_run=args.dry_run))
        if not args.cad_only:
            show("Diagnostics -> PDF", hk.diagnostics_to_pdf(job_dir, dry_run=args.dry_run))
    if args.dry_run:
        print("\n(dry run: nothing was changed)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
