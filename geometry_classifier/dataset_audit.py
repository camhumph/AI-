#!/usr/bin/env python3
"""Audit the YOLO datasets and training runs before anyone trusts a number.

WHY THIS EXISTS
===============
Two of the four datasets in this repo have **identical train and validation
images**. Every mAP ever reported for `major8_full_views` and
`right_view_major8` was therefore measured on the pictures the model was trained
on. Those numbers are not generalization estimates; they are memorisation
scores, and the highest one (0.746) came from the most leaked run.

That is not a small bookkeeping error. It is the difference between "this model
finds plates" and "this model has seen these four pictures". A leak like that
gets worse the more you train, so the reported figure improves while the real
one does not -- which is exactly the pattern in these runs.

This script makes that class of mistake impossible to ship unnoticed. Run it
before promoting any weights.

WHAT IT CHECKS
==============
  1. LEAKAGE      -- any image in both train and val.
  2. LABELS       -- images with no label file, and label files that are empty
                     (an empty label is a legitimate background image, but three
                     out of four "labelled" images being empty is not).
  3. BALANCE      -- box count per class per split. A class with no val boxes
                     cannot be evaluated at all.
  4. OVERFIT      -- for each run, the epoch of peak mAP50 vs the final epoch.
                     Promoting `best.pt` from a run that peaked at epoch 126 of
                     300 is fine; not noticing it peaked there is not.
  5. AUGMENTATION -- all-augmentation-off on a tiny dataset is a guaranteed
                     memoriser. Flagged from the run's own args.yaml.
  6. PROVENANCE   -- which file in models/ matches which run, by size, so
                     "which model is in production" has an answer.

USAGE
    python dataset_audit.py                       # audit everything
    python dataset_audit.py --root C:\\CMS_AI
    python dataset_audit.py --fix-split           # rebuild leaked splits
    python dataset_audit.py --fix-split --dry-run

`--fix-split` moves images so train and val are disjoint. It will NOT invent
data: with three labelled images there is no honest split, and it says so and
refuses rather than producing a 2/1 split that means nothing.
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import shutil
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

IMG_EXT = {".jpg", ".jpeg", ".png", ".bmp", ".webp"}

# Below this many labelled images, a train/val split is meaningless. Chosen so
# that a val set can hold at least 2 images and still leave 6 to train on.
MIN_IMAGES_FOR_SPLIT = 8


def read_yaml_names(p: Path) -> Dict[int, str]:
    """Class names from a data.yaml, without pulling in a YAML dependency."""
    names: Dict[int, str] = {}
    if not p.exists():
        return names
    in_names = False
    for line in p.read_text(encoding="utf-8", errors="replace").splitlines():
        s = line.rstrip()
        if s.startswith("names:"):
            in_names = True
            continue
        if in_names:
            if s and not s.startswith((" ", "\t", "-")):
                break
            t = s.strip()
            if not t:
                continue
            if ":" in t:
                k, v = t.split(":", 1)
                k = k.strip().lstrip("-").strip()
                if k.isdigit():
                    names[int(k)] = v.strip().strip("'\"")
    return names


def stem_set(d: Path) -> Dict[str, Path]:
    if not d.is_dir():
        return {}
    return {p.stem: p for p in sorted(d.iterdir()) if p.suffix.lower() in IMG_EXT}


def audit_dataset(ds: Path) -> Dict[str, Any]:
    names = read_yaml_names(ds / "data.yaml")
    out: Dict[str, Any] = {
        "dataset": ds.name,
        "classes": names,
        "splits": {},
        "problems": [],
    }

    per_split: Dict[str, Dict[str, Path]] = {}
    for split in ("train", "val", "test"):
        imgs = stem_set(ds / "images" / split)
        if not imgs:
            continue
        lbl_dir = ds / "labels" / split
        labelled = 0
        empty = 0
        missing: List[str] = []
        boxes: Counter = Counter()
        for stem in imgs:
            lp = lbl_dir / f"{stem}.txt"
            if not lp.exists():
                missing.append(stem)
                continue
            body = lp.read_text(encoding="utf-8", errors="replace").strip()
            if not body:
                empty += 1
                continue
            labelled += 1
            for line in body.splitlines():
                parts = line.split()
                if parts and parts[0].isdigit():
                    boxes[int(parts[0])] += 1

        per_split[split] = imgs
        out["splits"][split] = {
            "images": len(imgs),
            "labelled": labelled,
            "empty_labels": empty,
            "missing_labels": len(missing),
            "boxes": dict(sorted(boxes.items())),
            "box_total": sum(boxes.values()),
        }
        if missing:
            out["problems"].append(
                f"{split}: {len(missing)} image(s) have no label file "
                f"(e.g. {', '.join(missing[:3])}) -- YOLO trains on these as "
                f"backgrounds, which silently teaches 'nothing here'."
            )
        if empty:
            out["problems"].append(
                f"{split}: {empty} label file(s) are EMPTY. Legitimate as "
                f"negatives, but they contribute no boxes."
            )

    # ---- 1. LEAKAGE, the big one ----
    tr = set(per_split.get("train", {}))
    va = set(per_split.get("val", {}))
    overlap = sorted(tr & va)
    out["leakage"] = {
        "count": len(overlap),
        "examples": overlap[:5],
        "fraction_of_val": round(len(overlap) / len(va), 3) if va else 0.0,
    }
    if overlap:
        out["problems"].insert(
            0,
            f"*** LEAKAGE: {len(overlap)} of {len(va)} validation images are ALSO "
            f"in train ({100 * len(overlap) / len(va):.0f}% of val). Every metric "
            f"from this dataset is measured on training data and means nothing. ***",
        )

    # ---- 3. classes with no validation support ----
    tr_boxes = out["splits"].get("train", {}).get("boxes", {})
    va_boxes = out["splits"].get("val", {}).get("boxes", {})
    unevaluable = [
        f"{cid}:{names.get(cid, '?')}" for cid in tr_boxes if cid not in va_boxes
    ]
    if unevaluable:
        out["problems"].append(
            f"{len(unevaluable)} class(es) have training boxes but NO validation "
            f"boxes, so their accuracy is unmeasured: {', '.join(unevaluable[:8])}"
        )

    total_labelled = sum(
        s.get("labelled", 0) for s in out["splits"].values()
    )
    if total_labelled < MIN_IMAGES_FOR_SPLIT:
        out["problems"].append(
            f"Only {total_labelled} image(s) carry any boxes. Below "
            f"{MIN_IMAGES_FOR_SPLIT} there is no honest train/val split to make "
            f"-- this is a data collection problem, not a splitting problem."
        )
    out["total_labelled"] = total_labelled
    return out


def audit_run(run: Path) -> Optional[Dict[str, Any]]:
    res = run / "results.csv"
    if not res.exists():
        return None
    rows: List[Dict[str, str]] = []
    with res.open(newline="", encoding="utf-8", errors="replace") as fh:
        for r in csv.DictReader(fh):
            rows.append({k.strip(): (v or "").strip() for k, v in r.items() if k})
    if not rows:
        return None

    def col(row: Dict[str, str], want: str) -> float:
        for k, v in row.items():
            if want in k:
                try:
                    return float(v)
                except ValueError:
                    return 0.0
        return 0.0

    maps = [col(r, "mAP50(B)") or col(r, "mAP50") for r in rows]
    if not any(maps):
        return None
    best_i = max(range(len(maps)), key=lambda i: maps[i])

    # WHICH DATASET the run used. Without this the mAP table is misleading: a
    # run on a leaked dataset sits in the same column as a run on a clean one and
    # looks better. Digging this out by hand is how the production model turned
    # out to be the worst one.
    dataset = ""
    args_p = run / "args.yaml"
    if args_p.exists():
        for line in args_p.read_text(encoding="utf-8", errors="replace").splitlines():
            if line.strip().startswith("data:"):
                raw = line.split(":", 1)[1].strip()
                parts = raw.replace("\\", "/").split("/")
                dataset = parts[-2] if len(parts) >= 2 else raw
                break

    out: Dict[str, Any] = {
        "run": run.name,
        "dataset": dataset,
        "epochs": len(rows),
        "best_epoch": best_i + 1,
        "best_mAP50": round(maps[best_i], 5),
        "final_mAP50": round(maps[-1], 5),
        "problems": [],
    }
    if maps[-1] < maps[best_i] * 0.9 and len(rows) > 20:
        out["problems"].append(
            f"peaked at epoch {best_i + 1}/{len(rows)} (mAP50 {maps[best_i]:.3f}) "
            f"then fell to {maps[-1]:.3f} -- {100 * (1 - maps[-1] / maps[best_i]):.0f}% "
            f"below peak. Trained well past useful; best.pt is right but the run "
            f"wasted {len(rows) - best_i - 1} epochs and the final weights are worse."
        )

    args_f = run / "args.yaml"
    if args_f.exists():
        txt = args_f.read_text(encoding="utf-8", errors="replace")
        aug_keys = ["mosaic", "fliplr", "translate", "scale", "hsv_h", "hsv_s", "hsv_v", "erasing"]
        off = []
        for k in aug_keys:
            for line in txt.splitlines():
                if line.strip().startswith(f"{k}:"):
                    try:
                        if float(line.split(":", 1)[1].strip()) == 0.0:
                            off.append(k)
                    except ValueError:
                        pass
                    break
        if len(off) >= 6:
            out["problems"].append(
                f"ALL augmentation disabled ({', '.join(off)}). On a small dataset "
                f"that removes the only defence against memorising it."
            )
        for line in txt.splitlines():
            if line.strip().startswith("patience:") and line.strip().startswith("epochs:") is False:
                try:
                    pat = float(line.split(":", 1)[1].strip())
                    if pat >= len(rows):
                        out["problems"].append(
                            f"patience={pat:g} >= epochs={len(rows)}, so early "
                            f"stopping could never fire."
                        )
                except ValueError:
                    pass
    return out


def model_provenance(root: Path) -> List[Dict[str, Any]]:
    """Match models/*.pt to the run weights they were copied from, by size."""
    by_size: Dict[int, List[str]] = defaultdict(list)
    for w in (root / "runs").rglob("weights/*.pt"):
        by_size[w.stat().st_size].append(str(w.relative_to(root)))
    out = []
    for m in sorted((root / "models").glob("*.pt")):
        sz = m.stat().st_size
        out.append(
            {
                "model": m.name,
                "size": sz,
                "matches_run_weights": by_size.get(sz, []),
            }
        )
    return out


def fix_split(ds: Path, dry_run: bool) -> List[str]:
    """Make train and val disjoint, or refuse and say why."""
    msgs: List[str] = []
    tr = stem_set(ds / "images" / "train")
    va = stem_set(ds / "images" / "val")
    overlap = sorted(set(tr) & set(va))
    if not overlap:
        return [f"{ds.name}: train/val already disjoint, nothing to do."]

    # Count how many have real boxes -- only those are worth splitting.
    def has_boxes(stem: str, split: str) -> bool:
        p = ds / "labels" / split / f"{stem}.txt"
        return p.exists() and bool(p.read_text(encoding="utf-8", errors="replace").strip())

    labelled = [s for s in tr if has_boxes(s, "train")]
    if len(labelled) < MIN_IMAGES_FOR_SPLIT:
        return [
            f"{ds.name}: REFUSING to split. Only {len(labelled)} labelled image(s) "
            f"({MIN_IMAGES_FOR_SPLIT} needed). Splitting three images into two and "
            f"one produces a validation score with no meaning -- it would replace a "
            f"number that is obviously wrong with one that merely looks plausible. "
            f"Label more images first; there are unlabelled candidates in "
            f"label_queue/."
        ]

    # Deterministic 20% val, by hash so it is stable across runs.
    ranked = sorted(labelled, key=lambda s: hashlib.sha1(s.encode()).hexdigest())
    n_val = max(2, round(0.2 * len(ranked)))
    val_stems = set(ranked[:n_val])

    for stem in overlap:
        keep_in = "val" if stem in val_stems else "train"
        drop_from = "train" if keep_in == "val" else "val"
        for sub, ext_dir in (("images", "images"), ("labels", "labels")):
            src = ds / ext_dir / drop_from
            for p in list(src.glob(f"{stem}.*")):
                msgs.append(f"{ds.name}: remove {p.relative_to(ds)} (kept in {keep_in})")
                if not dry_run:
                    p.unlink()
    return msgs


def main(argv: Optional[List[str]] = None) -> int:
    ap = argparse.ArgumentParser(description="Audit YOLO datasets and runs.")
    ap.add_argument("--root", default=r"C:\CMS_AI", help="CMS_AI root")
    ap.add_argument("--fix-split", action="store_true", help="rebuild leaked splits")
    ap.add_argument("--dry-run", action="store_true", help="with --fix-split, show only")
    ap.add_argument("--json", help="write the full audit here")
    args = ap.parse_args(argv)

    root = Path(args.root)
    if not root.is_dir():
        print(f"root not found: {root}", file=sys.stderr)
        return 1

    report: Dict[str, Any] = {"root": str(root), "datasets": [], "runs": [], "models": []}

    ds_root = root / "datasets"
    datasets = [d for d in sorted(ds_root.iterdir()) if (d / "images").is_dir()] if ds_root.is_dir() else []

    print("=" * 78)
    print("DATASETS")
    print("=" * 78)
    for ds in datasets:
        a = audit_dataset(ds)
        report["datasets"].append(a)
        print(f"\n{ds.name}   ({len(a['classes'])} classes, {a['total_labelled']} labelled images)")
        for split, s in a["splits"].items():
            print(
                f"  {split:<6} {s['images']:>4} images  {s['labelled']:>4} with boxes  "
                f"{s['empty_labels']:>3} empty  {s['missing_labels']:>3} unlabelled  "
                f"{s['box_total']:>4} boxes"
            )
        if a["leakage"]["count"]:
            print(f"  LEAKAGE {a['leakage']['count']} shared images ({100 * a['leakage']['fraction_of_val']:.0f}% of val)")
        for p in a["problems"]:
            print(f"    - {p}")

    print("\n" + "=" * 78)
    print("TRAINING RUNS")
    print("=" * 78)
    runs_root = root / "runs" / "detect"
    if runs_root.is_dir():
        rs = []
        for run in sorted(runs_root.iterdir()):
            if not run.is_dir():
                continue
            a = audit_run(run)
            if a:
                rs.append(a)
        rs.sort(key=lambda d: -d["best_mAP50"])
        report["runs"] = rs
        # Mark every run whose dataset leaks, so the table cannot be read wrong.
        leaked_ds = {a["dataset"] for a in report["datasets"] if a["leakage"]["count"]}
        for a in rs:
            a["trustworthy"] = a["dataset"] not in leaked_ds and bool(a["dataset"])

        print(f"\n{'run':<34} {'dataset':<20} {'ep':>4} {'best@':>6} {'mAP50':>7} {'final':>7}  trust")
        for a in rs:
            mark = "  ok " if a["trustworthy"] else " LEAK"
            print(
                f"{a['run'][:34]:<34} {a['dataset'][:20]:<20} {a['epochs']:>4} "
                f"{a['best_epoch']:>6} {a['best_mAP50']:>7.4f} {a['final_mAP50']:>7.4f} {mark}"
            )

        honest = [a for a in rs if a["trustworthy"]]
        if honest:
            b = honest[0]
            print(
                f"\n  BEST HONEST RESULT: {b['run']} on {b['dataset']}, "
                f"mAP50 {b['best_mAP50']:.4f} at epoch {b['best_epoch']}/{b['epochs']}."
            )
            print(
                "  That is the only kind of number worth promoting. Anything marked "
                "LEAK was scored on its own training images."
            )
        for a in rs:
            if a["problems"]:
                print(f"\n  {a['run']}")
                for p in a["problems"]:
                    print(f"    - {p}")

    print("\n" + "=" * 78)
    print("MODEL PROVENANCE  (which models/*.pt came from which run)")
    print("=" * 78)
    prov = model_provenance(root)
    report["models"] = prov
    run_by_name = {a["run"]: a for a in report["runs"]}
    for m in prov:
        src = m["matches_run_weights"][0] if m["matches_run_weights"] else "NO MATCHING RUN"
        # Attach the run's score and trust flag. A model file on its own tells you
        # nothing; the pair (which run, was that run honest) is the whole story.
        run_name = src.replace("\\", "/").split("/")[2] if m["matches_run_weights"] else ""
        info = run_by_name.get(run_name)
        if info:
            m["run"] = run_name
            m["run_mAP50"] = info["best_mAP50"]
            m["run_dataset"] = info["dataset"]
            m["trustworthy"] = info["trustworthy"]
            tag = "ok" if info["trustworthy"] else "LEAKED"
            print(
                f"  {m['model']:<40} <- {run_name:<34} "
                f"mAP50 {info['best_mAP50']:.4f} on {info['dataset']} [{tag}]"
            )
        else:
            print(f"  {m['model']:<40} <- {src}")

    leaked = {a["dataset"] for a in report["datasets"] if a["leakage"]["count"]}
    if leaked:
        print("\n" + "!" * 78)
        print("Any run trained on these datasets has an UNTRUSTWORTHY score:")
        for d in sorted(leaked):
            print(f"  - {d}")
        print("Do not compare their mAP against a cleanly-split dataset.")
        print("!" * 78)

    if args.fix_split:
        print("\n" + "=" * 78)
        print("FIX SPLIT" + ("  (dry run)" if args.dry_run else ""))
        print("=" * 78)
        for ds in datasets:
            for line in fix_split(ds, args.dry_run):
                print(f"  {line}")

    if args.json:
        Path(args.json).write_text(json.dumps(report, indent=2, default=str), encoding="utf-8")
        print(f"\nwrote {args.json}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
