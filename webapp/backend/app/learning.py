"""Actual-vs-quoted history, so the estimator learns THIS shop.

WHY THIS FILE EXISTS
    Every calibration constant in the estimator came from two parts: J8441 and
    J8410, both ID POTs, both 4140, both on the UMC-1000. They are the best data
    that was available and they are still two parts. A clamp plate in A36 is a
    different animal and nothing in the model knows that yet.

    The fix is not more hand-tuned constants. It is to close the loop: record
    what each job ACTUALLY took, compare it against what was quoted, and derive
    the corrections from the shop's own history. After thirty jobs the model is
    calibrated on thirty jobs instead of two, and it keeps improving without
    anyone editing code.

STORAGE
    A single append-only JSONL file. Deliberately not a database:

      * The whole app is a local tool with no server to administer.
      * Append-only means a bad entry can be found and corrected in a text
        editor, and the history of what was believed when is never destroyed.
      * A quoting model that silently rewrites its own training data is not
        auditable, and auditability is the entire point of this record.

    Corrections are appended as new rows with the same job_id; readers take the
    LAST row per job. Nothing is ever edited in place or deleted.
"""
from __future__ import annotations

import json
import statistics
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional

from . import config

ACTUALS_PATH = Path(
    getattr(config, "LEARNING_ACTUALS_PATH", config.DATA_DIR / "job_actuals.jsonl")
)

# Below this many samples a class keeps the measured factor. One job is an
# anecdote; mirrors MIN_SAMPLES in the frontend's shopLearning.ts.
MIN_SAMPLES = 4

# A learned factor outside this range means the DATA is wrong -- a mis-keyed
# actual, or a job where half the work was subcontracted -- not the model.
LEARN_CLAMP = (0.5, 2.0)

OP_CLASSES = (
    "faceMill",
    "rough",
    "finish",
    "semiFinish",
    "drill",
    "spot",
    "tap",
    "chamfer",
    "bore",
    "contour",
    "setup",
    "toolChange",
)


def _ensure_parent() -> None:
    ACTUALS_PATH.parent.mkdir(parents=True, exist_ok=True)


def append_actual(entry: Dict[str, Any]) -> Dict[str, Any]:
    """Append one record. Never mutates or removes an existing row."""
    _ensure_parent()
    row = dict(entry)
    row["recorded_at"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    with ACTUALS_PATH.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(row, ensure_ascii=False) + "\n")
    return row


def _read_rows() -> List[Dict[str, Any]]:
    if not ACTUALS_PATH.exists():
        return []
    out: List[Dict[str, Any]] = []
    with ACTUALS_PATH.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                out.append(json.loads(line))
            except json.JSONDecodeError:
                # A single corrupt line must not take the whole history down.
                continue
    return out


def load_actuals() -> List[Dict[str, Any]]:
    """Latest row per job_id, oldest job first.

    Appending a correction for a job supersedes the earlier row without
    destroying it -- the raw file still holds both.
    """
    latest: Dict[str, Dict[str, Any]] = {}
    for row in _read_rows():
        jid = str(row.get("job_id") or "")
        if not jid:
            continue
        latest[jid] = row
    return sorted(latest.values(), key=lambda r: str(r.get("recorded_at") or ""))


def _clamp(v: float) -> float:
    return max(LEARN_CLAMP[0], min(LEARN_CLAMP[1], v))


def _iqr_spread_pct(xs: List[float]) -> int:
    if len(xs) < 4:
        return 0
    s = sorted(xs)
    q1 = s[len(s) // 4]
    q3 = s[(3 * len(s)) // 4]
    med = statistics.median(s)
    return int(round(100 * (q3 - q1) / med)) if med > 0 else 0


def _ratios_by_class(rows: Iterable[Dict[str, Any]]) -> Dict[str, List[float]]:
    by: Dict[str, List[float]] = {}
    for r in rows:
        per = r.get("by_class") or {}
        if not isinstance(per, dict):
            continue
        for cls, v in per.items():
            if cls not in OP_CLASSES or not isinstance(v, dict):
                continue
            q = float(v.get("quoted_min") or 0)
            a = float(v.get("actual_min") or 0)
            if q <= 0 or a <= 0:
                continue
            by.setdefault(cls, []).append(a / q)
    return by


def learn_factors() -> List[Dict[str, Any]]:
    """Per-class correction factors from recorded actuals.

    Median, not mean: shop-floor actuals contain typos and jobs that got
    scrapped and re-run, and a single 10x entry would otherwise move the model.
    """
    rows = load_actuals()
    out: List[Dict[str, Any]] = []
    for cls, ratios in _ratios_by_class(rows).items():
        med = statistics.median(ratios)
        clamped = _clamp(med)
        enough = len(ratios) >= MIN_SAMPLES
        spread = _iqr_spread_pct(ratios)

        if not enough:
            note = (
                f"{len(ratios)} of {MIN_SAMPLES} jobs needed. "
                "Measured factor still in use."
            )
        elif abs(clamped - med) > 1e-9:
            note = (
                f"Median ratio {med:.2f} clamped to {clamped:.2f}. A correction that "
                "large usually means a mis-keyed actual or subcontracted work, not a "
                "model error -- check the entries before trusting it."
            )
        elif spread > 40:
            note = (
                f"Median {clamped:.2f} but the middle half of jobs spans {spread}%. "
                "The model is now unbiased on average while still being wrong job to "
                "job -- look for what separates the fast ones from the slow ones."
            )
        else:
            note = f"Median of {len(ratios)} jobs, middle half within {spread}%."

        out.append(
            {
                "cls": cls,
                "factor": round(clamped, 3) if enough else 1.0,
                "samples": len(ratios),
                "spread_pct": spread,
                "applied": enough,
                "note": note,
            }
        )
    out.sort(key=lambda d: abs(d["factor"] - 1), reverse=True)
    return out


def accuracy_report() -> Optional[Dict[str, Any]]:
    """How the estimator is doing, in the terms a shop owner cares about.

    Bias and scatter are reported separately on purpose. An estimator that is
    unbiased but scattered is not the same as one that is consistently 15% low,
    and the fixes differ: scatter needs better inputs, bias needs a factor.
    """
    rows = [
        r
        for r in load_actuals()
        if float(r.get("quoted_min") or 0) > 0 and float(r.get("actual_min") or 0) > 0
    ]
    if not rows:
        return None

    ratios = [float(r["actual_min"]) / float(r["quoted_min"]) for r in rows]
    bias = statistics.median(ratios)
    mape = statistics.median([abs(x - 1) * 100 for x in ratios])
    within = sum(1 for x in ratios if 0.8 <= x <= 1.2)
    under = sum(1 for x in ratios if x > 1.05)
    over = sum(1 for x in ratios if x < 0.95)

    if len(rows) < MIN_SAMPLES:
        summary = (
            f"{len(rows)} job{'' if len(rows) == 1 else 's'} recorded. "
            f"Needs {MIN_SAMPLES} before any of this is worth acting on."
        )
    elif 0.98 <= bias <= 1.02:
        summary = (
            f"Across {len(rows)} jobs the estimator is unbiased (median {bias:.2f}x) "
            f"with a typical miss of {mape:.0f}%. {within} of {len(rows)} landed "
            "within 20%."
        )
    else:
        direction = "LOW" if bias > 1 else "HIGH"
        tail = (
            " Quoting low is the expensive direction -- this is money already lost."
            if direction == "LOW"
            else ""
        )
        summary = (
            f"Across {len(rows)} jobs the estimator runs {direction} by "
            f"{abs(bias - 1) * 100:.0f}% (median {bias:.2f}x), typical miss "
            f"{mape:.0f}%. {under} jobs came in over quote, {over} under.{tail}"
        )

    return {
        "jobs": len(rows),
        "bias": round(bias, 3),
        "mape_pct": round(mape, 1),
        "within_twenty_pct": int(round(100 * within / len(rows))),
        "under_quoted": under,
        "over_quoted": over,
        "summary": summary,
    }
