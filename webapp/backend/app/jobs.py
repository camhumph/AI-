"""Job discovery + classification loading/execution.

A "job" is a folder under config.JOBS_ROOT. Recognized contents:

  meta.json                          optional display metadata
  XT_Export_CAD_Dimensions.csv       raw SolidWorks CAD export (source of truth)
  classification.csv / classification.json   AI classification result (this
                                      app writes these; matches the shape
                                      produced by geometry_classifier/qwen_classify_xt_csv.py)
  images/*.jpg|*.png                 rendered views (ISO/front/back/left/right...)
  models/*.stl                       3D printable/viewable geometry
  documents/*                        quote sheets, steel sheets, PDFs, etc.
"""
import csv
import json
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

from . import config
from .roles import role_label, role_group

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp"}
MODEL_EXTS = {".stl"}
DOC_EXTS = {".pdf", ".csv", ".xlsx", ".xls", ".txt", ".docx"}


def _job_dir(job_id: str) -> Path:
    safe = job_id.strip().replace("..", "").replace("/", "_")
    return config.JOBS_ROOT / safe


def _list_assets(job_dir: Path, subfolder: str, exts: set) -> list:
    folder = job_dir / subfolder
    out = []
    if folder.exists():
        for p in sorted(folder.iterdir()):
            if p.is_file() and p.suffix.lower() in exts:
                out.append(
                    {
                        "name": p.name,
                        "url": f"/api/jobs/{job_dir.name}/file/{subfolder}/{p.name}",
                        "size": p.stat().st_size,
                    }
                )
    return out


def _read_meta(job_dir: Path) -> dict:
    meta_path = job_dir / "meta.json"
    if meta_path.exists():
        try:
            return json.loads(meta_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _read_classification(job_dir: Path):
    json_path = job_dir / "classification.json"
    if json_path.exists():
        try:
            return json.loads(json_path.read_text(encoding="utf-8"))
        except Exception:
            pass
    return None


def _read_classification_csv_by_index(job_dir: Path) -> dict:
    """classification.json holds index/role/confidence/reason/quote only.
    classification.csv (written by qwen_classify_xt_csv.py) also carries the
    original Component/Thickness/Width/Length/Center* columns -- merge those
    in so the UI, pricing engine, and VBA bridge all see real part names and
    dimensions.
    """
    csv_path = job_dir / "classification.csv"
    by_index = {}
    if not csv_path.exists():
        return by_index
    try:
        with csv_path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            for row in csv.DictReader(f):
                by_index[str(row.get("Index", ""))] = row
    except Exception:
        pass
    return by_index


def list_jobs() -> list:
    jobs = []
    if not config.JOBS_ROOT.exists():
        return jobs
    for job_dir in sorted(config.JOBS_ROOT.iterdir()):
        if not job_dir.is_dir():
            continue
        meta = _read_meta(job_dir)
        classification = _read_classification(job_dir)
        has_raw = (job_dir / "XT_Export_CAD_Dimensions.csv").exists()
        part_count = len(classification.get("classifications", [])) if classification else 0
        sequenced = bool(
            classification
            and classification.get("job_analysis", {}).get("sequenced_latch_lock_base")
        )
        jobs.append(
            {
                "job_id": job_dir.name,
                "display_name": meta.get("display_name", job_dir.name),
                "customer": meta.get("customer", ""),
                "notes": meta.get("notes", ""),
                "base_type": meta.get("base_type", "standard"),
                "has_raw_csv": has_raw,
                "has_classification": classification is not None,
                "part_count": part_count,
                "image_count": len(_list_assets(job_dir, "images", IMAGE_EXTS)),
                "model_count": len(_list_assets(job_dir, "models", MODEL_EXTS)),
                "sequenced_latch_lock_base": sequenced,
                "updated_at": meta.get("updated_at", ""),
            }
        )
    return jobs


def get_job(job_id: str) -> dict:
    job_dir = _job_dir(job_id)
    if not job_dir.exists():
        return None

    meta = _read_meta(job_dir)
    classification = _read_classification(job_dir) or {"job_analysis": {}, "classifications": []}
    csv_by_index = _read_classification_csv_by_index(job_dir)

    rows = []
    for item in classification.get("classifications", []):
        role = item.get("role", "")
        csv_row = csv_by_index.get(str(item.get("index", "")), {})
        rows.append(
            {
                **item,
                "role_label": role_label(role),
                "role_group": role_group(role),
                "Component": csv_row.get("Component", ""),
                "Thickness": csv_row.get("Thickness", ""),
                "Width": csv_row.get("Width", ""),
                "Length": csv_row.get("Length", ""),
                "CenterX": csv_row.get("CenterX", ""),
                "CenterY": csv_row.get("CenterY", ""),
                "CenterZ": csv_row.get("CenterZ", ""),
            }
        )

    return {
        "job_id": job_dir.name,
        "display_name": meta.get("display_name", job_dir.name),
        "customer": meta.get("customer", ""),
        "notes": meta.get("notes", ""),
        "base_type": meta.get("base_type", "standard"),
        "job_analysis": classification.get("job_analysis", {}),
        "parts": rows,
        "images": _list_assets(job_dir, "images", IMAGE_EXTS),
        "models": _list_assets(job_dir, "models", MODEL_EXTS),
        "documents": _list_assets(job_dir, "documents", DOC_EXTS),
        "has_raw_csv": (job_dir / "XT_Export_CAD_Dimensions.csv").exists(),
        "has_classification": classification is not None and bool(classification.get("classifications")),
    }


def get_asset_path(job_id: str, subfolder: str, filename: str) -> Path:
    job_dir = _job_dir(job_id)
    path = (job_dir / subfolder / filename).resolve()
    if job_dir.resolve() not in path.parents:
        return None
    return path if path.exists() else None


def create_job(job_id: str, display_name: str = "", customer: str = "") -> dict:
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    (job_dir / "images").mkdir(exist_ok=True)
    (job_dir / "models").mkdir(exist_ok=True)
    (job_dir / "documents").mkdir(exist_ok=True)
    if not (job_dir / "meta.json").exists():
        meta = {
            "display_name": display_name or job_id,
            "customer": customer,
            "notes": "",
            "base_type": "standard",
            "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        }
        (job_dir / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return get_job(job_id)


def delete_job(job_id: str) -> dict:
    """Remove a quote/job from the webapp registry (JOBS_ROOT folder + quote status)."""
    job_dir = _job_dir(job_id)
    if not job_dir.exists():
        raise FileNotFoundError(f"Job '{job_id}' not found")
    # Safety: only delete under JOBS_ROOT
    root = config.JOBS_ROOT.resolve()
    resolved = job_dir.resolve()
    if root not in resolved.parents and resolved != root:
        raise PermissionError("Refusing to delete outside jobs root")
    shutil.rmtree(resolved)

    # Also clear quote status file if present
    try:
        from . import quote_pipeline

        status_path = quote_pipeline._status_path(job_id)
        if status_path.exists():
            status_path.unlink()
        # Also try C-number variants
        for alt in (job_id.replace("-", ""), job_id.upper(), job_id.lower()):
            p = quote_pipeline._status_path(alt)
            if p.exists():
                p.unlink()
    except Exception:
        pass

    return {"deleted": True, "job_id": job_id}


def update_meta(job_id: str, **fields) -> None:
    job_dir = _job_dir(job_id)
    if not job_dir.exists():
        return
    meta = _read_meta(job_dir)
    meta.update(fields)
    meta["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    (job_dir / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")


def import_raw_csv(job_id: str, csv_path: str) -> bool:
    """Copy a local XT_Export_CAD_Dimensions.csv (e.g. from the SolidWorks
    macro's job folder) into this job. Returns False when unreadable."""
    src = Path(csv_path)
    if not src.exists() or not src.is_file():
        return False
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    (job_dir / "XT_Export_CAD_Dimensions.csv").write_bytes(src.read_bytes())
    return True


def _extract_c_number(name: str) -> str | None:
    import re

    m = re.search(r"\bC(\d{4,6})\b", name, re.I)
    return f"C{m.group(1)}" if m else None


def _workspace_roots() -> list[Path]:
    from . import config as cfg

    roots = [cfg.WORKSPACE_ROOT, cfg.JOBS_ROOT]
    roots.extend(Path(p) for p in cfg.WORKSPACE_EXTRA_ROOTS)
    seen: set[str] = set()
    out: list[Path] = []
    for r in roots:
        key = str(r)
        if key not in seen:
            seen.add(key)
            out.append(r)
    return out


def browse_workspace(path: str = "", quick: bool = True) -> dict:
    """List subfolders under a workspace path for the folder picker.

    quick=True (default): only list names + C-number from folder name — no
    per-folder glob of quote/steel sheets (those are slow on network shares).
    """
    from . import config as cfg

    base = Path(path) if path else cfg.WORKSPACE_ROOT
    if not base.exists() or not base.is_dir():
        # Fall back to first existing root
        for root in _workspace_roots():
            if root.exists():
                base = root
                break
        else:
            return {"path": str(base), "exists": False, "entries": [], "roots": [str(r) for r in _workspace_roots()]}

    entries = []
    try:
        children = list(base.iterdir())
        children.sort(key=lambda p: (not p.is_dir(), p.name.lower()))
        for child in children:
            if child.name.startswith("."):
                continue
            is_dir = child.is_dir()
            c_num = _extract_c_number(child.name) if is_dir else None
            # Fast path: skip expensive network globs when browsing month folders
            has_xt = False
            has_quote = False
            has_steel = False
            if is_dir and not quick:
                has_xt = (child / "XT_Export_CAD_Dimensions.csv").exists()
                has_quote = any(child.glob("*quote*.xls*")) or any(child.glob("*Quote*.xls*"))
                has_steel = any(child.glob("*steel*.xls*")) or any(child.glob("*J000*.xls*"))
            elif is_dir and quick:
                # Cheap single-file check only (no wildcards)
                has_xt = (child / "XT_Export_CAD_Dimensions.csv").exists()
            entries.append(
                {
                    "name": child.name,
                    "path": str(child),
                    "is_dir": is_dir,
                    "c_number": c_num,
                    "has_xt_csv": has_xt,
                    "has_quote_sheet": has_quote,
                    "has_steel_sheet": has_steel,
                    "quote_ready": is_dir and (bool(c_num) or has_xt or has_quote or has_steel),
                }
            )
    except PermissionError:
        pass

    parent = str(base.parent) if base.parent != base else None
    return {
        "path": str(base),
        "exists": True,
        "parent": parent,
        "entries": entries,
        "roots": [str(r) for r in _workspace_roots()],
    }


def _folder_looks_like_bms(folder: Path) -> bool:
    """Detect BMS / Tempcraft pot-block jobs primarily from folder/file *names*.

    Shop rule: most BMS jobs have "BMS" (or Tempcraft/Howmet/pot-block) in the
    folder or CAD filename. Do not scan arbitrary CSV/log text for SMED/holder
    tokens — that falsely tags standard Dynacast/DME jobs.
    """
    blob = folder.name.lower()
    if any(m in blob for m in ("bms", "tempcraft", "howmet", "potblock", "pot-block", "pot_block")):
        return True
    # Whole-token HTE customer prefix in folder name.
    if re.search(r"(^|[^a-z0-9])hte([^a-z0-9]|$)", blob):
        return True

    bms_name_tokens = (
        "bms",
        "tempcraft",
        "howmet",
        "potblock",
        "pot-block",
        "pot_block",
        "rfq_mb_asm",
        "mb_asm",
    )
    cad_exts = {".sldasm", ".sldprt", ".x_t", ".x_b", ".step", ".stp", ".iges", ".igs"}
    try:
        for path in folder.rglob("*"):
            if not path.is_file():
                continue
            low = path.name.lower()
            # Strongest: BMS / Tempcraft / pot-block in any file name (esp. CAD).
            if any(m in low for m in bms_name_tokens):
                return True
            # CAD files named with holder/pot block are also BMS.
            if path.suffix.lower() in cad_exts and any(
                m in low for m in ("holder block", "pot block", "id holder", "od holder", "smed")
            ):
                return True
            # Only trust the macro's own base-type log line — not free-text BOM dumps.
            if path.name.lower() in (
                "cms_base_export_log.txt",
                "cms_training_xt_log.txt",
            ) and path.stat().st_size < 2_000_000:
                try:
                    text = path.read_text(encoding="utf-8", errors="ignore")[:12000].upper()
                except Exception:
                    continue
                if "BASE TYPE: POT" in text or "BASE TYPE FORCED POT/BMS" in text:
                    return True
                if "BASE TYPE: STANDARD" in text or "BASE TYPE STANDARD FROM" in text:
                    return False
                if "BASE TYPE DEFAULT STANDARD" in text:
                    return False
    except Exception:
        pass
    return False


def _hoist_macro_deliverables(src: Path, job_dir: Path) -> dict:
    """Copy Module6121 root/base deliverables into images/ and models/.

    The macro writes `{base} ISO.jpg`, `{base} BACK ISO.jpg`, and `{base}.stl`
    to the job ROOT (and a copy under base\\). The UI only lists images/ and
    models/, so without this hoist the Quotes page shows No STL / No images.
    """
    images = job_dir / "images"
    models = job_dir / "models"
    images.mkdir(exist_ok=True)
    models.mkdir(exist_ok=True)
    copied = {"images": 0, "models": 0}

    search_roots = [src]
    base_sub = src / "base"
    if base_sub.is_dir():
        search_roots.append(base_sub)

    for root in search_roots:
        for f in root.iterdir():
            if not f.is_file():
                continue
            low = f.name.lower()
            ext = f.suffix.lower()
            if ext in IMAGE_EXTS and (
                "iso" in low or "front" in low or "back" in low or "view" in low
            ):
                dest = images / f.name
                if not dest.exists() or dest.stat().st_size != f.stat().st_size:
                    shutil.copy2(f, dest)
                    copied["images"] += 1
            elif ext in MODEL_EXTS:
                dest = models / f.name
                if not dest.exists() or dest.stat().st_size != f.stat().st_size:
                    shutil.copy2(f, dest)
                    copied["models"] += 1
    return copied


def _is_pot_block_dims(t: float, w: float, l: float, max_fp: float) -> bool:
    """Pots are thick, chunky, and clearly smaller than the mold footprint.

    Full-size A/B plates fail this (flat + large footprint).
    """
    if t < 3.0 or w <= 0 or l <= 0:
        return False
    if (l / w) > 1.7:
        return False
    fp = w * l
    if max_fp > 0 and fp >= 0.55 * max_fp:
        return False
    dim_max = max(t, w, l)
    dim_min = min(t, w, l)
    if dim_max <= 0:
        return False
    if (dim_min / dim_max) < 0.35:
        return False
    return True


def _xt_looks_like_pot_block(job_dir: Path) -> bool:
    """Geometry heuristic: 2 thin full clamps + thick holders + distinguishable pots.

    Must NEVER fire on a standard mold stack (5+ full-footprint plates).
    Prefer real pot cubes over insulation sheets alone.
    """
    xt = job_dir / "XT_Export_CAD_Dimensions.csv"
    if not xt.exists():
        return False
    try:
        with xt.open(newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.DictReader(fh))
    except Exception:
        return False
    if len(rows) < 6:
        return False

    def _f(row: dict, *keys: str) -> float:
        for k in keys:
            if k in row and str(row[k]).strip():
                try:
                    return float(str(row[k]).strip())
                except ValueError:
                    pass
        return 0.0

    parsed = []
    for r in rows:
        t = _f(r, "Thickness", "thickness")
        w = _f(r, "Width", "width")
        l = _f(r, "Length", "length")
        if t > 0 and w > 0 and l > 0:
            parsed.append((t, w, l, w * l))
    if len(parsed) < 6:
        return False

    max_fp = max(p[3] for p in parsed)
    max_w = max(p[1] for p in parsed)
    max_l = max(p[2] for p in parsed)

    # Standard mold stacks have many same-size full-footprint plates.
    full_plates = [
        p for p in parsed
        if p[1] >= max_w * 0.85 and p[2] >= max_l * 0.85 and p[0] >= 0.5
    ]
    if len(full_plates) >= 5:
        return False

    full_thin = [
        p for p in parsed
        if p[3] >= max_fp * 0.85 and 0.75 <= p[0] <= 2.5
    ]
    thick_inner = [
        p for p in parsed
        if p[0] >= 3.0 and p[3] < max_fp * 0.85 and p[3] >= max_fp * 0.15
    ]
    thin_sheets = [p for p in parsed if abs(p[0] - 0.25) <= 0.06]
    pot_like = [p for p in parsed if _is_pot_block_dims(p[0], p[1], p[2], max_fp)]

    # Prefer real pots; sheets alone without pots are not enough.
    if len(full_thin) <= 2 and len(thick_inner) >= 2 and (
        len(pot_like) >= 2 or (len(thin_sheets) >= 2 and len(pot_like) >= 1)
    ):
        return True
    return False


def import_from_folder(folder_path: str, run_quote: bool = False) -> dict:
    """Register a quote from an existing folder on disk (C-number job)."""
    src = Path(folder_path)
    if not src.exists() or not src.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder_path}")

    c_num = _extract_c_number(src.name)
    job_id = c_num or src.name[:40]
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)

    # Copy key files into the job registry
    for pattern in (
        "XT_Export_CAD_Dimensions.csv",
        "Purchased Components Quote.csv",
        "Pullcore Prices.csv",
        "classification.json",
        "classification.csv",
    ):
        for hit in src.glob(pattern):
            if hit.is_file():
                shutil.copy2(hit, job_dir / hit.name)

    for sub in ("images", "models", "documents"):
        src_sub = src / sub
        if src_sub.is_dir():
            dest_sub = job_dir / sub
            dest_sub.mkdir(exist_ok=True)
            for f in src_sub.iterdir():
                if f.is_file():
                    shutil.copy2(f, dest_sub / f.name)

    # Module6121 writes STL + ISO JPGs at job root / base\ — hoist into UI folders.
    _hoist_macro_deliverables(src, job_dir)

    # Flat files in job root -> documents/
    docs = job_dir / "documents"
    docs.mkdir(exist_ok=True)
    for f in src.iterdir():
        if not f.is_file():
            continue
        low = f.name.lower()
        if low.endswith((".xlsx", ".xls", ".pdf", ".csv")) and f.name not in {
            "XT_Export_CAD_Dimensions.csv",
            "classification.csv",
        }:
            dest = docs / f.name
            if not dest.exists():
                shutil.copy2(f, dest)

    meta = _read_meta(job_dir)
    meta.update(
        {
            "display_name": src.name,
            "customer": re.search(r"\d{6,}", src.name).group(0) if re.search(r"\d{6,}", src.name) else "",
            "source_folder": str(src),
            "c_number": c_num or "",
            "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        }
    )
    if "created_at" not in meta:
        meta["created_at"] = meta["updated_at"]

    # Force BMS when geometry/folder looks like pot-block (holders + insulation).
    if meta.get("base_type") != "bms":
        if _folder_looks_like_bms(src) or _xt_looks_like_pot_block(job_dir):
            meta["base_type"] = "bms"

    # Drop stale standard-stack classification when this is a pot-block job.
    if meta.get("base_type") == "bms":
        for stale in ("classification.json", "classification.csv"):
            p = job_dir / stale
            if p.exists():
                try:
                    p.unlink()
                except Exception:
                    pass

    (job_dir / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

    job = get_job(job_id)

    # Never run standard A/B/rail classify on BMS / pot-block jobs.
    if (
        meta.get("base_type") != "bms"
        and job.get("has_raw_csv")
        and not job.get("has_classification")
    ):
        try:
            job = classify_job(job_id, mode="rules")
        except Exception:
            pass

    if run_quote:
        from . import quote_pipeline

        qid = c_num or job_id
        create_job(qid, display_name=src.name, customer=meta.get("customer", ""))
        quote_pipeline.launch_full_quote(
            qid,
            str(src),
            {
                "subject": src.name,
                "cust_job": meta.get("customer", ""),
                "attachments": sum(1 for _ in src.rglob("*") if _.is_file()),
            },
        )
        return {
            "job_id": qid,
            "quote_id": qid,
            "quote_started": True,
            "poll_url": f"/api/quote/status/{qid}",
            "display_name": src.name,
        }

    return job


def save_upload(job_id: str, subfolder: str, filename: str, data: bytes) -> dict:
    job_dir = _job_dir(job_id)
    job_dir.mkdir(parents=True, exist_ok=True)
    target_dir = job_dir / subfolder
    target_dir.mkdir(parents=True, exist_ok=True)
    safe_name = Path(filename).name
    (target_dir / safe_name).write_bytes(data)
    return get_job(job_id)


def classify_job(job_id: str, mode: str = "rules") -> dict:
    """(Re)run the AI classifier against this job's raw XT_Export CSV.

    mode="rules" uses the deterministic geometry/shop-token rules only
    (fast, no LLM required -- this is what runs in this sandbox).
    mode="llm" additionally tries Ollama/Qwen if it is installed and on
    PATH (matches geometry_classifier/qwen_classify_xt_csv.py --long-knowledge).
    """
    job_dir = _job_dir(job_id)
    raw_csv = job_dir / "XT_Export_CAD_Dimensions.csv"
    if not raw_csv.exists():
        raise FileNotFoundError(
            f"No XT_Export_CAD_Dimensions.csv found for job '{job_id}'. "
            "Upload the raw CAD export first."
        )

    script = config.GEOMETRY_CLASSIFIER_DIR / "qwen_classify_xt_csv.py"
    if not script.exists():
        raise FileNotFoundError(f"Classifier script not found at {script}")

    args = [sys.executable, str(script), str(raw_csv), "--include-names", "--max-rows", "5000"]
    if mode == "rules":
        args.append("--rules-only")
    else:
        args.append("--long-knowledge")

    result = subprocess.run(args, capture_output=True, text=True, timeout=600)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Classifier failed")

    stem = raw_csv.stem
    out_dir = config.GEOMETRY_CLASSIFIER_DIR / "outputs"
    produced_json = out_dir / f"{stem}_qwen_classification.json"
    produced_csv = out_dir / f"{stem}_qwen_classification.csv"

    if produced_json.exists():
        (job_dir / "classification.json").write_bytes(produced_json.read_bytes())
    if produced_csv.exists():
        (job_dir / "classification.csv").write_bytes(produced_csv.read_bytes())

    meta = _read_meta(job_dir)
    meta["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    (job_dir / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

    return get_job(job_id)
