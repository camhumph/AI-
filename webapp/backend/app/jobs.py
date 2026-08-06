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

from . import config, job_housekeeping, name_learning, plate_names
from .roles import role_label, role_group

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp"}
MODEL_EXTS = {".stl"}
DOC_EXTS = {".pdf", ".csv", ".xlsx", ".xls", ".txt", ".docx"}


class InvalidJobId(ValueError):
    """job_id was empty or unusable after sanitising."""


def _job_dir(job_id: str) -> Path:
    safe = (job_id or "").strip().replace("..", "").replace("/", "_").replace("\\", "_")
    # An empty id made `JOBS_ROOT / ""` resolve to JOBS_ROOT ITSELF, so a job
    # with a blank C-number wrote meta.json at the registry root and could
    # rmtree the whole registry. C17267 hit the front half of this: the macro
    # POSTed job-complete with a blank folder ("DONE ACTIVE CAD QUOTE. Output
    # folder:" with nothing after it), sync_completed_job resolved Path("") to
    # ".", and import_from_folder derived job_id = "".
    if not safe:
        raise InvalidJobId("job_id is empty after sanitising")
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


def _read_xt_mass_by_index(job_dir: Path) -> dict:
    """Measured mass in lb per CAD index, from XT_Export_CAD_Dimensions.csv.

    The macro gets this from SolidWorks' own CreateMassProperty (mass x 2.20462),
    falling back to volume x 0.283 when a part has no material assigned. It is a
    real measurement, so the UI should prefer it over estimating mass from STL
    mesh volume -- an STL carries no material data and no units.

    Column is `Mass_or_Vol`; older exports may spell it `Mass`.
    """
    by_index = {}
    csv_path = job_housekeeping.find_job_file(job_dir, "XT_Export_CAD_Dimensions.csv")
    if csv_path is None:
        return by_index
    try:
        with csv_path.open("r", newline="", encoding="utf-8-sig", errors="replace") as f:
            for row in csv.DictReader(f):
                raw = row.get("Mass_or_Vol", row.get("Mass", ""))
                try:
                    val = float(str(raw).strip())
                except (TypeError, ValueError):
                    continue
                if val > 0:
                    by_index[str(row.get("Index", "")).strip()] = val
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
        has_raw = job_housekeeping.job_file_exists(job_dir, "XT_Export_CAD_Dimensions.csv")
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
    mass_by_index = _read_xt_mass_by_index(job_dir)

    base_type = meta.get("base_type", "standard")

    rows = []
    for item in classification.get("classifications", []):
        role = item.get("role", "")
        idx = str(item.get("index", ""))
        csv_row = csv_by_index.get(idx, {})
        # What the shop has already told us this role is called, from a rename on
        # an earlier job. Falls straight back to the built-in label when nothing
        # has been learned. See app/name_learning.py.
        label = name_learning.label_for(role, base_type, role_label(role))
        # Measured lb from SolidWorks. classification.csv does not carry it, so
        # fall back to the raw XT export keyed on the same CAD index.
        mass = csv_row.get("Mass_or_Vol", csv_row.get("Mass", "")) or mass_by_index.get(idx, "")
        rows.append(
            {
                **item,
                "role_label": label,
                "role_group": role_group(role),
                "Component": csv_row.get("Component", ""),
                "Thickness": csv_row.get("Thickness", ""),
                "Width": csv_row.get("Width", ""),
                "Length": csv_row.get("Length", ""),
                "CenterX": csv_row.get("CenterX", ""),
                "CenterY": csv_row.get("CenterY", ""),
                "CenterZ": csv_row.get("CenterZ", ""),
                "MassLb": str(mass) if mass != "" else "",
            }
        )

    # User renames, applied last so they win over the role's default label.
    # role is deliberately NOT touched -- see plate_names.py.
    overrides = plate_names.get_overrides(job_dir)
    plate_names.apply_to_rows(rows, overrides)

    return {
        "job_id": job_dir.name,
        "display_name": meta.get("display_name", job_dir.name),
        "customer": meta.get("customer", ""),
        "notes": meta.get("notes", ""),
        "base_type": base_type,
        "job_analysis": classification.get("job_analysis", {}),
        "parts": rows,
        "plate_name_overrides": overrides,
        "images": _list_assets(job_dir, "images", IMAGE_EXTS),
        "models": _list_assets(job_dir, "models", MODEL_EXTS),
        "documents": _list_assets(job_dir, "documents", DOC_EXTS),
        "has_raw_csv": job_housekeeping.job_file_exists(job_dir, "XT_Export_CAD_Dimensions.csv"),
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


def job_dir(job_id: str) -> Path | None:
    """Public, existence-checked accessor. None when the job is not registered.

    `_job_dir` raises on a blank id and does not check existence, so callers
    outside this module would have to duplicate both checks.
    """
    try:
        d = _job_dir(job_id)
    except InvalidJobId:
        return None
    return d if d.exists() else None


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
    m = re.search(r"[-_]C(\d{4,6})\b", name, re.I) or re.search(r"\bC[- ]?(\d{4,6})\b", name, re.I)
    return f"C{m.group(1)}" if m else None


def _looks_like_quote_job(name: str) -> bool:
    """True for C-number / BMS job folders the shop can quote from New Quote."""
    if not name or name.startswith("."):
        return False
    if _extract_c_number(name):
        return True
    u = name.upper().replace(" ", "")
    if u.startswith("BMS-") or u.startswith("BMS_"):
        return True
    # e.g. 863700122-C18613 already covered; also bare long job ids with quote token
    if re.search(r"\d{6,}.*C\d{4,6}", name, re.I):
        return True
    return False


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
                has_xt = job_housekeeping.job_file_exists(child, "XT_Export_CAD_Dimensions.csv")
                has_quote = any(child.glob("*quote*.xls*")) or any(child.glob("*Quote*.xls*"))
                has_steel = any(child.glob("*steel*.xls*")) or any(child.glob("*J000*.xls*"))
            elif is_dir and quick:
                # Cheap single-file check only (no wildcards)
                has_xt = job_housekeeping.job_file_exists(child, "XT_Export_CAD_Dimensions.csv")
            quote_ready = is_dir and (
                _looks_like_quote_job(child.name) or bool(c_num) or has_xt or has_quote or has_steel
            )
            entries.append(
                {
                    "name": child.name,
                    "path": str(child),
                    "is_dir": is_dir,
                    "c_number": c_num,
                    "has_xt_csv": has_xt,
                    "has_quote_sheet": has_quote,
                    "has_steel_sheet": has_steel,
                    "quote_ready": quote_ready,
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
    """Copy Module6121 root/base/stl deliverables into images/ and models/.

    The macro writes `{base} ISO.jpg`, `{base} BACK ISO.jpg`, and `{base}.stl`
    to the job ROOT (and a copy under base\\). The UI only lists images/ and
    models/, so without this hoist the Quotes page shows No STL / No images.

    It also writes one STL per quoted plate into `stl\\` when
    EXPORT_PER_PLATE_STLS is on -- `{base}_A Plate.STL`, `{base}_Rails 1.STL`,
    and so on, all in the same corrected CMS frame as the full-assembly STL.
    Those are hoisted too, which is what fills the 3D tab's part gallery.
    """
    images = job_dir / "images"
    models = job_dir / "models"
    images.mkdir(exist_ok=True)
    models.mkdir(exist_ok=True)
    copied = {"images": 0, "models": 0, "pruned": 0}

    search_roots = [src]
    base_sub = src / "base"
    if base_sub.is_dir():
        search_roots.append(base_sub)
    # Per-plate STLs from ExportPlateStlsForComparison.
    stl_sub = src / "stl"
    if stl_sub.is_dir():
        search_roots.append(stl_sub)

    hoisted_plates = set()
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
                if root == stl_sub:
                    hoisted_plates.add(f.name)

    if stl_sub.is_dir():
        copied["pruned"] = _prune_stale_plate_models(models, hoisted_plates)
    return copied


# Names hoisted out of the macro's stl\ folder on the previous import. Kept so a
# plate that disappears between runs (a re-classify collapsing "Rails 1"/"Rails
# 2" into one "Rails", say) can be removed from models/ without ever touching an
# STL the user uploaded by hand -- only files this function put there are
# candidates for deletion.
_PLATE_MANIFEST = ".plate_stls.json"


def _prune_stale_plate_models(models: Path, current: set) -> int:
    manifest = models / _PLATE_MANIFEST
    previous = set()
    if manifest.exists():
        try:
            loaded = json.loads(manifest.read_text(encoding="utf-8"))
            if isinstance(loaded, list):
                previous = {str(n) for n in loaded}
        except Exception:
            previous = set()

    pruned = 0
    for name in previous - current:
        stale = models / name
        try:
            if stale.is_file():
                stale.unlink()
                pruned += 1
        except Exception:
            pass

    try:
        manifest.write_text(json.dumps(sorted(current), indent=2), encoding="utf-8")
    except Exception:
        pass
    return pruned


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
    xt = job_housekeeping.find_job_file(job_dir, "XT_Export_CAD_Dimensions.csv")
    if xt is None:
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
    job_id = (c_num or src.name[:40] or "").strip()
    if not job_id:
        # A blank folder_path resolves to "." and would derive an empty job_id,
        # which used to mean "the registry root". Refuse instead.
        raise InvalidJobId(
            f"Could not derive a job id from folder {folder_path!r}. "
            "Check that the macro sent a real CurrentJobFolder in its "
            "job-complete POST."
        )
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
                # Skip dotfiles: re-importing a folder that is itself a job
                # registry would otherwise drag in its .plate_stls.json and
                # prune this job's plates against a manifest it never wrote.
                if f.is_file() and not f.name.startswith("."):
                    shutil.copy2(f, dest_sub / f.name)

    # Module6121 writes STL + ISO JPGs at job root / base\ — hoist into UI folders.
    _hoist_macro_deliverables(src, job_dir)

    # Flat files in job root -> documents/
    # Always refresh quote/steel Excel copies — a stale first sync (pre-fill)
    # left empty Description / "sheet" prices on the Parts tab.
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
            refresh = (
                not dest.exists()
                or "quote" in low
                or "steel" in low
                or "j000" in low
                or "grind" in low
                or "purchased" in low
            )
            if refresh:
                try:
                    shutil.copy2(f, dest)
                except Exception:
                    pass
            # Also keep the primary quote/steel workbooks at job root for pricing.
            if any(k in low for k in ("quote", "steel", "j000", "grind")) and low.endswith(
                (".xlsx", ".xls", ".xlsm")
            ):
                try:
                    shutil.copy2(f, job_dir / f.name)
                except Exception:
                    pass

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
                "c_number": c_num or qid,
                "job_folder": src.name,
                "attachments": 0,
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


def import_folders_batch(folder_paths: list[str], run_quote: bool = True) -> dict:
    """Register multiple C-number folders and launch them as one SolidWorks batch."""
    if not folder_paths:
        return {"launched": False, "error": "No folders selected", "quote_ids": [], "jobs": []}

    from . import quote_pipeline

    items: list[dict] = []
    jobs_out: list[dict] = []
    errors: list[str] = []

    for raw in folder_paths:
        path = str(raw or "").strip()
        if not path:
            continue
        try:
            job = import_from_folder(path, run_quote=False)
            src = Path(path)
            c_num = _extract_c_number(src.name) or str(job.get("job_id") or "")
            cust = ""
            m = re.search(r"\d{6,}", src.name)
            if m:
                cust = m.group(0)
            qid = (c_num or job.get("job_id") or src.name).upper()
            create_job(qid, display_name=src.name, customer=cust)
            items.append(
                {
                    "quote_id": qid,
                    "attach_dir": str(src),
                    "c_number": c_num or qid,
                    "email_info": {
                        "subject": src.name,
                        "cust_job": cust,
                        "c_number": c_num or qid,
                        "job_folder": src.name,
                        "attachments": 0,
                    },
                }
            )
            jobs_out.append(
                {
                    "job_id": qid,
                    "quote_id": qid,
                    "folder_path": str(src),
                    "display_name": src.name,
                }
            )
        except Exception as e:
            errors.append(f"{path}: {e}")

    if not items:
        return {
            "launched": False,
            "error": "; ".join(errors) if errors else "No valid folders",
            "quote_ids": [],
            "jobs": [],
            "errors": errors,
        }

    launch: dict = {"launched": False}
    if run_quote:
        launch = quote_pipeline.launch_batch_quotes(items)

    return {
        "launched": bool(launch.get("launched")),
        "batch": True,
        "batch_count": len(items),
        "quote_ids": launch.get("quote_ids") or [i["quote_id"] for i in items],
        "c_numbers": launch.get("c_numbers") or [i["c_number"] for i in items],
        "jobs": jobs_out,
        "errors": errors,
        "poll_urls": [f"/api/quote/status/{i['quote_id']}" for i in items],
        "error": launch.get("error") or ("; ".join(errors) if errors and not launch.get("launched") else None),
    }


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
    mode="stl" runs the two-pass STL naming: Qwen picks candidate names from the
    whole CAD export, the exported per-plate meshes for those candidates are
    measured, and Qwen makes a final call against the real geometry. Slowest and
    best-evidenced -- it is the only mode that can see a plate's true stack
    thickness or which face its pockets open on.
    """
    job_dir = _job_dir(job_id)
    raw_csv = job_housekeeping.find_job_file(job_dir, "XT_Export_CAD_Dimensions.csv")
    if raw_csv is None:
        raise FileNotFoundError(
            f"No XT_Export_CAD_Dimensions.csv found for job '{job_id}'. "
            "Upload the raw CAD export first."
        )

    script = config.GEOMETRY_CLASSIFIER_DIR / "qwen_classify_xt_csv.py"
    if not script.exists():
        raise FileNotFoundError(f"Classifier script not found at {script}")

    args = [sys.executable, str(script), str(raw_csv), "--include-names", "--max-rows", "5000"]
    # Two LLM passes plus the mesh measurement do not fit in the 10 minutes the
    # single-pass modes need, and a timeout here loses the whole run.
    timeout = 600
    if mode == "rules":
        args.append("--rules-only")
    elif mode == "stl":
        args += ["--stl-two-pass", "--long-knowledge"]
        timeout = 5400
    else:
        # --no-stl keeps this mode byte-identical to what it did before the STL
        # path existed; the CLI would otherwise pick two-pass up automatically
        # whenever a job has meshes.
        args += ["--long-knowledge", "--no-stl"]

    result = subprocess.run(args, capture_output=True, text=True, timeout=timeout)
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

    # Overlay the customer BOM's own plate names where a detail number matches a
    # CAD component. The classifier reasons from bounding boxes and hole
    # topology and never sees the BOM, but the BOM usually just says which plate
    # is which -- "Stationary retainer plate", "Ejector retainer plate" -- and a
    # matched detail number is a lookup, not a guess. Never fatal: a bad BOM must
    # not cost us the geometry classification we already have.
    try:
        from . import bom_roles

        hints = bom_roles.refresh(job_dir)
        if hints.get("hints"):
            data = json.loads((job_dir / "classification.json").read_text(encoding="utf-8"))
            data = bom_roles.apply_to_classification(job_dir, data)
            data.setdefault("job_analysis", {}).setdefault("bom_role_hints", {})["bom"] = \
                hints.get("bom", "")
            (job_dir / "classification.json").write_text(
                json.dumps(data, indent=2), encoding="utf-8"
            )
    except Exception as e:
        print(f"[bom_roles] skipped for {job_id}: {type(e).__name__}: {e}")

    meta = _read_meta(job_dir)
    meta["updated_at"] = time.strftime("%Y-%m-%dT%H:%M:%S")
    (job_dir / "meta.json").write_text(json.dumps(meta, indent=2), encoding="utf-8")

    return get_job(job_id)
