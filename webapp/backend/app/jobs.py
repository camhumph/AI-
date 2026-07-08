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
        "job_analysis": classification.get("job_analysis", {}),
        "parts": rows,
        "images": _list_assets(job_dir, "images", IMAGE_EXTS),
        "models": _list_assets(job_dir, "models", MODEL_EXTS),
        "documents": _list_assets(job_dir, "documents", DOC_EXTS),
        "has_raw_csv": (job_dir / "XT_Export_CAD_Dimensions.csv").exists(),
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
    meta = {
        "display_name": display_name or job_id,
        "customer": customer,
        "notes": "",
        "created_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
        "updated_at": time.strftime("%Y-%m-%dT%H:%M:%S"),
    }
    (job_dir / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")
    return get_job(job_id)


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
