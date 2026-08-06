"""Learn the shop's plate names from the renames the estimator types.

WHY THIS EXISTS
    Renaming a plate used to be a per-job patch. The estimator opened C18328,
    saw the app call a plate "OD Holder", typed "OD Holders", and the next pot
    job called it "OD Holder" again -- because nothing had learned anything. The
    same three keystrokes, forever, on every job.

    A rename is the highest-quality training signal this app can get. It is a
    human looking at real geometry and stating, unambiguously, what the shop
    calls that plate. It costs nothing to record and it is never ambiguous, which
    is more than can be said for most labels a model trains on.

WHAT IS LEARNED, AND WHAT IS DELIBERATELY NOT
    Learned:  role -> the label the shop wants, scoped to the base type.
              "On a BMS job, od_holder is called 'OD Holders'."

    NOT learned: the role itself. A rename changes the LABEL, never the ROLE --
              see plate_names.py, where that separation is the whole safety
              story. So when a rename looks like it is really saying "this is a
              different plate" (no word in common with the old name), this module
              records that diagnosis and FLAGS it, rather than quietly re-keying
              the plate and moving money between quote rows. A misclassification
              is a thing to show someone, not a thing to silently absorb.

APPLIED IMMEDIATELY, FROM ONE EXAMPLE
    One rename is enough. The alternative -- wait for N jobs to agree -- is the
    right call for a MEASUREMENT like a machining factor, where a single reading
    is noise (see learning.py, MIN_SAMPLES = 4). It is the wrong call for a NAME:
    the estimator is not sampling a distribution, they are telling you the
    answer. Making them tell you four times is just making them type it four
    times.

    The safety net is not a sample threshold, it is scope. A learned label only
    ever replaces the label THIS APP WOULD HAVE GENERATED for that role. A plate
    the workbook names something else entirely is left alone. So the worst case
    of a bad learned label is that a default the shop already had to correct by
    hand shows up wrong again -- exactly where it was before -- and one more
    rename overwrites it.

STORAGE
    Append-only JSONL beside the machining actuals, same reasoning as
    learning.py: no server to administer, a bad row can be found and fixed in a
    text editor, and a model that silently rewrites its own training data cannot
    be audited. Readers take the LAST row per (base_type, role).
"""
from __future__ import annotations

import json
import re
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, List, Optional

from . import config

LEARNED_PATH = Path(
    getattr(config, "PLATE_NAME_LEARNING_PATH", config.DATA_DIR / "plate_name_learning.jsonl")
)

# Roles that identify nothing in particular, so a label learned against one would
# apply to every unmapped row at once. Same set plate_names.GENERIC_ROLES uses.
GENERIC_ROLES = {"", "steel_plate", "purchased_component", "other", "ignore", "hardware_other"}

# Base types are kept apart because the same word means different plates in each.
# A "TCP" on a pot base is not the plate a standard base calls a top clamp.
BASE_TYPES = ("standard", "bms")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _norm(s: Any) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(s or "").lower()).strip()


def _tokens(s: Any) -> List[str]:
    return [t for t in _norm(s).split(" ") if t]


def _base_type(value: Any) -> str:
    v = str(value or "").strip().lower()
    return "bms" if v in ("bms", "pot", "pot_block") else "standard"


# --------------------------------------------------------------------------
# WHY WAS IT RENAMED?
#
# The reason is not decoration. Each one points at a different thing to fix, and
# the difference between them is the difference between "adjust a label" and
# "the classifier put the wrong role on this plate".
# --------------------------------------------------------------------------
def diagnose(old: str, new: str) -> Dict[str, str]:
    """Classify a rename into a reason code, a sentence, and what to do about it."""
    o, n = str(old or "").strip(), str(new or "").strip()
    on, nn = _norm(o), _norm(n)
    ot, nt = _tokens(o), _tokens(n)

    def out(code: str, why: str, fix: str, role_suspect: bool = False) -> Dict[str, str]:
        return {"code": code, "why": why, "fix": fix, "role_suspect": role_suspect}

    if not on or not nn:
        return out("unknown", "Not enough of the old name survived to compare.", "")

    if o != n and on == nn:
        # Same letters after normalising. If they also match once case is taken
        # out, the ONLY difference was case; otherwise a quote or a hyphen moved.
        if o.lower() == n.lower():
            return out(
                "capitalisation",
                f"Same words, different capitalisation: {o!r} -> {n!r}.",
                "The shop writes this name in a fixed case. Learned as written.",
            )
        return out(
            "punctuation",
            f"Same words, different punctuation or spacing: {o!r} -> {n!r}.",
            "The generated label is punctuated the way the app wants it, not the "
            "way the steel sheet reads. Learned as written.",
        )

    if nn == on + "s" or nn == on + "es" or on == nn + "s" or on == nn + "es":
        plural = len(nn) > len(on)
        return out(
            "plural" if plural else "singular",
            f"{'Pluralised' if plural else 'Made singular'}: {o!r} -> {n!r}.",
            "The label counted the plate wrong -- this slot holds "
            f"{'more than one' if plural else 'exactly one'} on this kind of base. "
            "Learned, so the count is right from the start next time.",
        )

    # An acronym in either direction: "Top Clamp Plate" <-> "TCP".
    def initials(tokens: List[str]) -> str:
        return "".join(t[0] for t in tokens if t)

    if len(nt) == 1 and len(ot) > 1 and nn.replace(" ", "") == initials(ot):
        return out(
            "abbreviation",
            f"Shortened to its initials: {o!r} -> {n!r}.",
            "The shop's sheets use the abbreviation. Learned, so the full name is "
            "not written out again.",
        )
    if len(ot) == 1 and len(nt) > 1 and on.replace(" ", "") == initials(nt):
        return out(
            "expansion",
            f"Abbreviation written out: {o!r} -> {n!r}.",
            "The shop wants this one spelled out. Learned.",
        )

    # "PLATE" IS NOT EVIDENCE THAT TWO NAMES MEAN THE SAME PLATE.
    #
    # Overlap is what separates "they adjusted the wording" from "they said this
    # is a different plate". Nearly every name in a mold base ends in the same
    # handful of nouns, so counting them as shared makes every pair look related:
    # "B Plate" -> "Manifold Plate" came back as a rewording when it is the
    # clearest possible signal that the classifier put the wrong role on a plate.
    _STOP = {"plate", "plt", "the", "a", "of", "and"}
    oset = {t for t in ot if t not in _STOP}
    nset = {t for t in nt if t not in _STOP}
    shared = oset & nset
    if shared and oset < nset:
        added = " ".join(t for t in nt if t in nset - oset)
        return out(
            "qualifier_added",
            f"Kept the name and added {added!r}: {o!r} -> {n!r}.",
            "The generated label was not specific enough to tell this plate from "
            "its neighbour. Learned with the qualifier.",
        )
    if shared and nset < oset:
        dropped = " ".join(t for t in ot if t in oset - nset)
        return out(
            "qualifier_dropped",
            f"Trimmed {dropped!r} off the name: {o!r} -> {n!r}.",
            "The generated label carried a word the shop does not use. Learned "
            "without it.",
        )

    # Both names were nothing but stopwords ("Plate" -> "Plt"). No overlap to
    # measure, and certainly no evidence of a different plate.
    if not oset and not nset:
        return out(
            "reworded",
            f"Reworded: {o!r} -> {n!r}.",
            "The shop words this one differently. Learned as written.",
        )

    if not shared:
        # THE ONE THAT IS NOT A LABEL PROBLEM.
        return out(
            "different_plate",
            f"{n!r} shares no word with {o!r}. This does not read as a relabel -- it "
            "reads as the plate having been identified as the wrong kind of plate.",
            "The LABEL is learned, but the plate keeps its role, its quote row and "
            "its price, so if the role really is wrong the money is still wrong. "
            "Check the role on the Parts tab and re-run classification if it is.",
            role_suspect=True,
        )

    return out(
        "reworded",
        f"Reworded: {o!r} -> {n!r}.",
        "The shop words this one differently. Learned as written.",
    )


# --------------------------------------------------------------------------
# The store
# --------------------------------------------------------------------------
def _ensure_parent() -> None:
    LEARNED_PATH.parent.mkdir(parents=True, exist_ok=True)


def _read_rows() -> List[Dict[str, Any]]:
    if not LEARNED_PATH.exists():
        return []
    out: List[Dict[str, Any]] = []
    try:
        with LEARNED_PATH.open("r", encoding="utf-8") as fh:
            for line in fh:
                line = line.strip()
                if not line:
                    continue
                try:
                    out.append(json.loads(line))
                except json.JSONDecodeError:
                    # One corrupt line must not cost the whole history.
                    continue
    except OSError:
        return []
    return out


def record_rename(
    *,
    job_id: str,
    base_type: str,
    key: str,
    role: str,
    old_name: str,
    new_name: str,
    component: str = "",
    dims: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """Log one rename and return what was learned from it. Never raises.

    Returns the entry, including `learned` -- true when this rename became the
    default label for that role. A rename on a generic role is still logged (it
    is evidence) but teaches nothing, because the role it would be filed under
    identifies nothing.
    """
    role = (role or "").strip()
    bt = _base_type(base_type)
    reason = diagnose(old_name, new_name)

    # TWO DIFFERENT QUESTIONS, AND CONFLATING THEM BROKE forget().
    #
    #   teaches  -- does this row participate in the fold at all? Yes whenever
    #               the role is specific enough to file under.
    #   learned  -- did it INSTALL a global label? Only when it carries a name.
    #
    # A clearing row (blank new name, written by forget()) teaches but does not
    # install: it exists precisely to remove the label a previous row installed.
    # With one flag doing both jobs the clearing row was skipped by
    # learned_labels() and the forgotten label came straight back.
    teaches = bool(role) and role not in GENERIC_ROLES
    has_name = bool(str(new_name).strip())
    teachable = teaches and has_name

    entry: Dict[str, Any] = {
        "recorded_at": _now(),
        "job_id": str(job_id or ""),
        "base_type": bt,
        "key": str(key or ""),
        "role": role,
        "old_name": str(old_name or ""),
        "new_name": str(new_name or ""),
        "component": str(component or ""),
        "dims": dims or {},
        "reason": reason["code"],
        "why": reason["why"],
        "fix": reason["fix"],
        "role_suspect": reason["role_suspect"],
        "teaches": teaches,
        "learned": teachable,
    }
    if not teaches:
        entry["not_learned_because"] = (
            f"Role {role!r} identifies nothing specific, so a label learned against it "
            "would rename every unmapped row on every future job. This job keeps the "
            "name; nothing global changed."
        )
    elif not has_name:
        entry["cleared"] = True

    try:
        _ensure_parent()
        with LEARNED_PATH.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(entry, ensure_ascii=False) + "\n")
    except OSError as e:
        # Losing the training signal must never lose the rename.
        entry["learned"] = False
        entry["not_learned_because"] = f"Could not write {LEARNED_PATH.name}: {e}"
    return entry


# Cached on the file's size and mtime, because label_for() is called once per
# part and a real mold base has hundreds of solids -- re-reading and re-folding
# the whole log for each one turned a page load into hundreds of file reads.
# Keyed on the stat rather than a timer so an edit to the .jsonl by hand, which
# is a supported way to fix a bad row, is picked up on the very next call.
_LABEL_CACHE: Dict[str, Any] = {"stamp": None, "labels": {}}


def _stamp() -> Optional[tuple]:
    try:
        st = LEARNED_PATH.stat()
    except OSError:
        return None
    return (st.st_mtime_ns, st.st_size)


def learned_labels() -> Dict[str, Dict[str, Any]]:
    """{"bms:od_holder": {...}} -- the current label per base type and role.

    Last write wins, with the count of how many times the shop has said it and
    which jobs those were. An entry whose new name is empty CLEARS the learned
    label, which is how a mistake gets taken back.
    """
    stamp = _stamp()
    if stamp is not None and _LABEL_CACHE["stamp"] == stamp:
        return _LABEL_CACHE["labels"]

    out: Dict[str, Dict[str, Any]] = {}
    for row in _read_rows():
        # `teaches` on rows this version wrote; older rows only had `learned`.
        if not row.get("teaches", row.get("learned")):
            continue
        role = str(row.get("role") or "").strip()
        if not role or role in GENERIC_ROLES:
            continue
        key = f"{_base_type(row.get('base_type'))}:{role}"
        label = str(row.get("new_name") or "").strip()
        if not label:
            out.pop(key, None)
            continue
        prev = out.get(key) or {}
        jobs = list(prev.get("jobs") or [])
        jid = str(row.get("job_id") or "")
        if jid and jid not in jobs:
            jobs.append(jid)
        out[key] = {
            "base_type": _base_type(row.get("base_type")),
            "role": role,
            "label": label,
            "replaces": str(row.get("old_name") or ""),
            "reason": row.get("reason") or "",
            "why": row.get("why") or "",
            "fix": row.get("fix") or "",
            "role_suspect": bool(row.get("role_suspect")),
            # Not a confidence gate -- see the module docstring. Shown so the
            # estimator can see whether a label is one person's one-off or the
            # settled house style.
            "samples": int(prev.get("samples") or 0) + 1,
            "jobs": jobs,
            "updated_at": row.get("recorded_at") or "",
        }

    _LABEL_CACHE["stamp"] = stamp
    _LABEL_CACHE["labels"] = out
    return out


def label_for(role: str, base_type: str, default: str = "") -> str:
    """The learned label for this role, or `default`.

    Cheap enough to call per row: the file holds one line per rename ever typed,
    which is tens of lines, not thousands.
    """
    role = (role or "").strip()
    if not role or role in GENERIC_ROLES:
        return default
    hit = learned_labels().get(f"{_base_type(base_type)}:{role}")
    return (hit or {}).get("label") or default


def apply_to_generated(role: str, base_type: str, current: str, generated: str) -> str:
    """Swap in the learned label ONLY where the app generated the name itself.

    THIS IS THE WHOLE SAFETY MARGIN, so it is worth being exact about.

    `generated` is what this app would call the role with nothing learned --
    roles.ROLE_LABELS, or sheet_pricing's canonical BMS label. `current` is what
    the row says right now, which on a steel row is the text the macro wrote into
    the workbook.

    The learned label replaces `current` only when the two match. A plate the
    workbook calls something the app did not invent is the shop having already
    said what it wants for THAT plate, and a global preference has no business
    overruling it.
    """
    cur = str(current or "").strip()
    gen = str(generated or "").strip()
    if not cur or (gen and _norm(cur) != _norm(gen)):
        return cur or gen
    learned = label_for(role, base_type, "")
    return learned or cur or gen


def history(limit: int = 200) -> List[Dict[str, Any]]:
    """Newest first. The audit trail behind every learned label."""
    rows = _read_rows()
    rows.reverse()
    return rows[: max(0, int(limit))]


def review_queue() -> List[Dict[str, Any]]:
    """Renames that look like a misclassification rather than a relabel.

    Surfaced separately because these are the ones where the LABEL is now right
    and the PRICE may still be wrong -- the role, and therefore the quote row,
    did not move. See the `different_plate` branch of diagnose().
    """
    seen = set()
    out: List[Dict[str, Any]] = []
    for row in history(limit=1000):
        if not row.get("role_suspect"):
            continue
        key = (row.get("job_id"), row.get("role"), row.get("new_name"))
        if key in seen:
            continue
        seen.add(key)
        out.append(row)
    return out


def forget(base_type: str, role: str, job_id: str = "") -> Dict[str, Any]:
    """Take back a learned label by appending a clearing row.

    Append, not delete: the history of what was believed when is the point of an
    append-only store, and a learned label that turned out to be wrong is one of
    the more interesting things it can hold.
    """
    return record_rename(
        job_id=job_id or "(cleared by hand)",
        base_type=base_type,
        key=f"role:{role}",
        role=role,
        old_name=label_for(role, base_type, ""),
        new_name="",
    )
