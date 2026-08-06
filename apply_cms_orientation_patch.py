#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
apply_cms_orientation_patch.py   (rev 2)
==============================================================================
Applies the "rotate the whole assembly as ONE rigid set" + "faces after BOM
matching" patch to the CMS XT export macro.

WHAT CHANGED IN REV 2
---------------------
Rev 1 rotated every LEAF PART component, which shears the mold apart:
Component2.Transform2 on a child of a sub-assembly is relative to its PARENT,
so rotating leaves AND their parents applies R twice to anything nested.
Rev 2 rotates TOP-LEVEL components only -- children ride along rigidly inside
their parent. It also suppresses mates first (so the solver cannot undo the
rotation) and re-fixes every component afterwards.

Rev 2 also stops using exact multi-line text anchors. It parses the module into
procedures and edits inside the right one, so comment/whitespace differences
between your builds no longer break the patch.

EDITS
-----
  0. declarations : add  Private gMoldGeometryIsSquareAndUpright As Boolean
                    (only if it is not declared already)
  1. ProcessOneJob : add  gMoldGeometryIsSquareAndUpright = False
  2. StraightenAssemblyFromPlateTransform : per-component loop
                    -> RotateWholeAssemblyAsOneRigidSetByMatrix
  3. ApplyMoldOrientationMatrix : same loop replacement
  4. ProcessOneJob : the *Top block -> FinalizeFacesAfterPhysicalRotation
  5. append every new helper procedure before "END OF MODULE"

USAGE
-----
  Find out which of your .bas files can actually take the patch:
      python apply_cms_orientation_patch.py --scan C:\\CMS_AI
      python apply_cms_orientation_patch.py --scan C:\\CMS_AI --recursive

  Patch one:
      python apply_cms_orientation_patch.py "C:\\path\\to\\gemini1.bas"

  Overwrite the original instead of writing a _PATCHED copy
  (a timestamped .bak byte-copy is always made first):
      python apply_cms_orientation_patch.py "C:\\path\\to\\gemini1.bas" --in-place

  Apply whatever edits are possible on an older build and report the rest:
      python apply_cms_orientation_patch.py "...\\old.bas" --allow-partial
==============================================================================
"""

import argparse
import datetime
import os
import re
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
NEW_FUNCS_DEFAULT = os.path.join(HERE, "CMS_Orientation_Patch_NewFunctions.bas")

FLAG = "gMoldGeometryIsSquareAndUpright"
ROTATE_FN = "RotateWholeAssemblyAsOneRigidSetByMatrix"
FACES_FN = "FinalizeFacesAfterPhysicalRotation"

# rev-1 procedure names -- if these are present the file carries the bad patch
REV1_ROTATE = "RotateEveryLeafPartComponentTogetherByMatrix"
REV1_NAMES = [
    REV1_ROTATE,
    "SelectEveryLeafPartComponent",
    "ApplyTransformToEveryLeafPartComponent",
    "ApplyTransformToSelectedComponents",
    "FloatSelectedComponents",
]

END_OF_MODULE_RE = re.compile(
    r"^' =+\n' END OF MODULE\n' =+\n?", re.MULTILINE)

PROC_START_RE = re.compile(
    r"^[ \t]*(?:(?:Private|Public|Friend)[ \t]+)?(?:Static[ \t]+)?"
    r"(Function|Sub)[ \t]+(\w+)", re.IGNORECASE)
PROC_END_RE = re.compile(r"^[ \t]*End[ \t]+(Function|Sub)[ \t]*$", re.IGNORECASE)

LOOP_RE = re.compile(
    r"^([ \t]*)For[ \t]+(\w+)[ \t]*=[ \t]*0[ \t]+To[ \t]+UBound\(vComps\)[ \t]*\n"
    r"[ \t]*If[ \t]+Not[ \t]+vComps\(\2\)[ \t]+Is[ \t]+Nothing[ \t]+Then[ \t]*\n"
    r"[ \t]*If[ \t]+StraightenOneComponentByMatrix\([ \t]*vComps\(\2\)[ \t]*,"
    r"[ \t]*R[ \t]*,[ \t]*swMathUtil[ \t]*\)[ \t]+Then[ \t]*\n"
    r"[ \t]*moved[ \t]*=[ \t]*moved[ \t]*\+[ \t]*1[ \t]*\n"
    r"[ \t]*Else[ \t]*\n"
    r"[ \t]*failed[ \t]*=[ \t]*failed[ \t]*\+[ \t]*1[ \t]*\n"
    r"[ \t]*End[ \t]+If[ \t]*\n"
    r"[ \t]*End[ \t]+If[ \t]*\n"
    r"[ \t]*Next[ \t]+\2[ \t]*\n",
    re.MULTILINE | re.IGNORECASE)


# ---------------------------------------------------------------------------
# small helpers
# ---------------------------------------------------------------------------
def die(msg):
    print("FAILED: %s" % msg, file=sys.stderr)
    sys.exit(1)


def read_text_lf(path):
    with open(path, "rb") as fh:
        raw = fh.read()
    for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    else:
        die("could not decode %s with any known encoding." % path)
    return text.replace("\r\n", "\n").replace("\r", "\n"), enc


def write_text_crlf(path, text, enc):
    if enc == "utf-8-sig":
        enc = "utf-8"
    with open(path, "wb") as fh:
        fh.write(text.replace("\n", "\r\n").encode(enc, errors="replace"))


def strip_code(s):
    """Remove string literals and trailing comment so keyword tests are safe."""
    out, in_str = [], False
    for ch in s:
        if ch == '"':
            in_str = not in_str
            continue
        if ch == "'" and not in_str:
            break
        if not in_str:
            out.append(ch)
        else:
            out.append(" ")
    return "".join(out)


def code_only(block):
    """Same text with every comment and string literal blanked out.

    Every name lookup MUST go through this. The macro's comments mention the
    very identifiers being searched for -- e.g.
        ' EnsureCmsTopOrientationFromMatchedTcpBcp here: its snap-to-nearest
    -- and matching raw text captures the prose instead of the code.
    """
    return "\n".join(strip_code(ln) for ln in block.split("\n"))


VBA_RESERVED = {
    "if", "then", "else", "elseif", "end", "and", "or", "not", "true", "false",
    "nothing", "here", "call", "set", "dim", "as", "is", "to", "for", "next",
    "do", "loop", "while", "with", "exit", "function", "sub", "select", "case",
}


def is_local_object_var(proc_body, name):
    """True if `name` is plausibly a variable of this procedure."""
    if not name or name.lower() in VBA_RESERVED:
        return False
    code = code_only(proc_body)
    pat = r"(?im)^[ \t]*(?:Dim|Set|Static)[ \t]+%s\b" % re.escape(name)
    return bool(re.search(pat, code))


def pick_model_var(text, proc):
    """Fallback: the procedure's own assembly-model variable."""
    code = code_only(text[proc["start"]:proc["end"]])
    for cand in ("swModel", "model", "swAssy", "assyModel"):
        if re.search(r"(?im)^[ \t]*(?:Dim|Set)[ \t]+%s\b" % cand, code):
            return cand
    m = re.search(r"(?im)^[ \t]*Set[ \t]+(\w+)[ \t]*=[ \t]*swApp\.ActiveDoc", code)
    if m:
        return m.group(1)
    return first_param_name(text, proc) or "swModel"


def find_procs(text):
    """Return [{name, kind, start, end}] with char offsets, end exclusive."""
    procs, lines, pos = [], text.split("\n"), 0
    open_proc = None
    for line in lines:
        line_start, pos = pos, pos + len(line) + 1
        bare = strip_code(line)
        if open_proc is None:
            m = PROC_START_RE.match(bare)
            # "End Function" also matches nothing here; declarations like
            # "Dim x As Object" do not start with Function/Sub
            if m and not PROC_END_RE.match(bare):
                open_proc = {"kind": m.group(1), "name": m.group(2),
                             "start": line_start}
        else:
            m = PROC_END_RE.match(bare)
            if m and m.group(1).lower() == open_proc["kind"].lower():
                open_proc["end"] = pos
                procs.append(open_proc)
                open_proc = None
    if open_proc is not None:
        open_proc["end"] = len(text)
        procs.append(open_proc)
    return procs


def get_proc(procs, name):
    for p in procs:
        if p["name"].lower() == name.lower():
            return p
    return None


def first_param_name(text, proc):
    """Name of a procedure's first parameter, e.g. 'model' or 'swModel'."""
    head = text[proc["start"]:proc["start"] + 600]
    i = head.find("(")
    if i < 0:
        return None
    depth, j = 0, i
    while j < len(head):
        if head[j] == "(":
            depth += 1
        elif head[j] == ")":
            depth -= 1
            if depth == 0:
                break
        j += 1
    params = head[i + 1:j]
    first = params.split(",")[0]
    first = re.sub(r"(?i)\b(ByVal|ByRef|Optional|ParamArray)\b", " ", first)
    m = re.search(r"(\w+)", first)
    return m.group(1) if m else None


def param_names(text, proc):
    head = text[proc["start"]:proc["start"] + 600]
    i = head.find("(")
    if i < 0:
        return set()
    depth, j = 0, i
    while j < len(head):
        if head[j] == "(":
            depth += 1
        elif head[j] == ")":
            depth -= 1
            if depth == 0:
                break
        j += 1
    out = set()
    for part in head[i + 1:j].split(","):
        part = re.sub(r"(?i)\b(ByVal|ByRef|Optional|ParamArray)\b", " ", part)
        m = re.search(r"(\w+)", part)
        if m:
            out.add(m.group(1))
    return out


def verify_patch(text):
    """Every identifier the patch wrote into a call must actually exist.

    This is the check that would have caught the 'here' bug: the patcher had
    written  EnsureCmsTopOrientationFromMatchedTcpBcp here, ...  because it
    scraped the name out of a comment, and VBA only complained at compile time.
    """
    problems = []
    call_re = re.compile(r"(?:%s[ \t]+(\w+)|%s\([ \t]*(\w+)[ \t]*,)"
                         % (re.escape(FACES_FN), re.escape(ROTATE_FN)))
    for p in find_procs(text):
        code = code_only(text[p["start"]:p["end"]])
        params = param_names(text, p)
        for m in call_re.finditer(code):
            var = m.group(1) or m.group(2)
            if var in params:
                continue
            if re.search(r"(?im)^[ \t]*(?:Dim|Set|Static)[ \t]+%s\b"
                         % re.escape(var), code):
                continue
            problems.append(
                "in %s: '%s' is passed to a patch call but is not a parameter "
                "or a Dim/Set variable of that procedure" % (p["name"], var))
    return problems


def logical_lines(block):
    """[(index_of_first_physical_line, joined_text)] with ' _' continuations merged."""
    raw, out, i = block.split("\n"), [], 0
    while i < len(raw):
        start, acc = i, raw[i]
        while acc.rstrip().endswith("_") and i + 1 < len(raw):
            i += 1
            acc = acc.rstrip()[:-1] + " " + raw[i].strip()
        out.append((start, acc))
        i += 1
    return out


def find_if_block(block, opener_regex):
    """Locate an If...End If block. Returns (first_line, last_line, indent) or None."""
    ll = logical_lines(block)
    start_idx = None
    indent = ""
    for k, (lineno, textline) in enumerate(ll):
        bare = strip_code(textline)
        if start_idx is None:
            if opener_regex.search(bare):
                start_idx = k
                indent = re.match(r"[ \t]*", textline).group(0)
                depth = 1
            continue
        if re.match(r"(?i)^[ \t]*If\b.*\bThen[ \t]*$", bare):
            depth += 1
        elif re.match(r"(?i)^[ \t]*End[ \t]+If[ \t]*$", bare):
            depth -= 1
            if depth == 0:
                last_lineno = lineno
                # the End If may itself have been a continuation start; use the
                # physical line of the next logical entry minus one
                if k + 1 < len(ll):
                    last_lineno = ll[k + 1][0] - 1
                else:
                    last_lineno = len(block.split("\n")) - 1
                return ll[start_idx][0], last_lineno, indent
    return None


# ---------------------------------------------------------------------------
# the edits
# ---------------------------------------------------------------------------
def edit_declare_flag(text):
    """E0 - make sure the flag is declared in the declarations section."""
    if re.search(r"(?im)^[ \t]*(?:Private|Public|Dim)[ \t]+%s[ \t]+As[ \t]+Boolean"
                 % FLAG, text):
        return text, "already declared"

    m = re.search(r"(?im)^([ \t]*)(?:Private|Public|Dim)[ \t]+"
                  r"FinalStlCoordFrameReady[ \t]+As[ \t]+Boolean[ \t]*\n", text)
    if not m:
        procs = find_procs(text)
        if not procs:
            return text, None
        ins = procs[0]["start"]
        block = "Private %s As Boolean\n\n" % FLAG
        return text[:ins] + block + text[ins:], "declared before first procedure"

    ins = m.end()
    block = "%sPrivate %s As Boolean\n" % (m.group(1), FLAG)
    return text[:ins] + block + text[ins:], "declared after FinalStlCoordFrameReady"


def edit_reset_flag(text):
    """E1 - reset the flag at the top of ProcessOneJob."""
    proc = get_proc(find_procs(text), "ProcessOneJob")
    if proc is None:
        return text, None
    body = text[proc["start"]:proc["end"]]

    if re.search(r"(?im)^[ \t]*%s[ \t]*=[ \t]*False" % FLAG, code_only(body)):
        return text, "already reset"

    m = re.search(r"(?im)^([ \t]*)FinalStlCoordFrameReady[ \t]*=[ \t]*False[ \t]*\n",
                  body)
    if not m:
        m = re.search(r"(?im)^([ \t]*)LastJobFailReason[ \t]*=[ \t]*\"\"[ \t]*\n", body)
    if not m:
        return text, None

    ins = proc["start"] + m.end()
    line = "%s%s = False\n" % (m.group(1), FLAG)
    return text[:ins] + line + text[ins:], "reset added"


def edit_rotation_loop(text, proc_name):
    """E2 / E3 - swap the per-component loop for the one rigid rotation."""
    proc = get_proc(find_procs(text), proc_name)
    if proc is None:
        return text, None
    body = text[proc["start"]:proc["end"]]

    if ROTATE_FN in code_only(body):
        return text, "already patched"

    m = LOOP_RE.search(body)
    if not m:
        return text, None

    var = first_param_name(text, proc)
    if not var or var.lower() in VBA_RESERVED:
        var = "model"
    ind = m.group(1)
    new = (
        "%s' PATCH: rotate the WHOLE assembly as one rigid set.\n"
        "%s' TOP-LEVEL components only -- Transform2 on a nested child is\n"
        "%s' relative to its parent, so rotating both applies R twice and\n"
        "%s' shears the mold apart. Children ride along inside their parent.\n"
        "%sIf %s(%s, R) Then\n"
        "%s    moved = 1\n"
        "%s    failed = 0\n"
        "%sElse\n"
        "%s    moved = 0\n"
        "%s    failed = 1\n"
        "%sEnd If\n"
        % (ind, ind, ind, ind, ind, ROTATE_FN, var,
           ind, ind, ind, ind, ind, ind)
    )
    lo = proc["start"] + m.start()
    hi = proc["start"] + m.end()
    return text[:lo] + new + text[hi:], "loop replaced (rotates %s)" % var


def edit_faces_block(text):
    """E4 - face/view correction after BOM matching."""
    proc = get_proc(find_procs(text), "ProcessOneJob")
    if proc is None:
        return text, None
    body = text[proc["start"]:proc["end"]]

    if FACES_FN in code_only(body):
        return text, "already patched"

    opener = re.compile(r"(?i)^[ \t]*If[ \t]+%s[ \t]+Then[ \t]*$" % FLAG)
    found = find_if_block(body, opener)
    if not found:
        return text, None
    first, last, ind = found

    lines = body.split("\n")
    block = "\n".join(lines[first:last + 1])

    # IMPORTANT: read the variable name from CODE only.
    # The original block carries the comment
    #     ' EnsureCmsTopOrientationFromMatchedTcpBcp here: its snap-to-nearest
    # and matching against raw text captures "here" instead of the model
    # variable, which then compiles as "Variable not defined".
    code = code_only(block)
    if "EnsureCmsTopOrientationFromMatchedTcpBcp" not in code:
        return text, None

    m = re.search(r"EnsureCmsTopOrientationFromMatchedTcpBcp[ \t]+(\w+)[ \t]*,", code)
    var = m.group(1) if m else None
    if not is_local_object_var(body, var):
        var = pick_model_var(text, proc)

    m2 = re.search(r"EnsureCmsTopOrientationFromMatchedTcpBcp[ \t]+\w+[ \t]*,"
                   r"[ \t]*(\w+)", code)
    persist = m2.group(1) if m2 else \
        "PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE"

    new_lines = [
        "%s' PATCH: the physical rotation already happened BEFORE BOM matching," % ind,
        "%s' so every ExportRows(i).CadPartIndex the matcher stored is still" % ind,
        "%s' valid. All that is left here is face/view correction:" % ind,
        "%s'     *Top   = TCP side" % ind,
        "%s'     *Front = pots closer than holders" % ind,
        "%sIf %s Then" % (ind, FLAG),
        "%s    %s %s" % (ind, FACES_FN, var),
        "%sElse" % ind,
        "%s    EnsureCmsTopOrientationFromMatchedTcpBcp %s, %s" % (ind, var, persist),
        "%sEnd If" % ind,
    ]
    lines[first:last + 1] = new_lines
    new_body = "\n".join(lines)
    return (text[:proc["start"]] + new_body + text[proc["end"]:],
            "block replaced (model var %s)" % var)


def remove_comment_region_containing(text, needle):
    """Delete the comment-only region that contains `needle`."""
    lines = text.split("\n")
    hit = None
    for i, ln in enumerate(lines):
        if needle in ln and ln.lstrip().startswith("'"):
            hit = i
            break
    if hit is None:
        return text, False

    def is_fluff(s):
        t = s.strip()
        return t == "" or t.startswith("'")

    # Never expand into the module's own "END OF MODULE" banner -- that banner
    # is comment lines too, and eating it would lose the append anchor.
    barrier = len(lines)
    for i, ln in enumerate(lines):
        if "END OF MODULE" in ln:
            barrier = max(0, i - 1)       # also spare the '=== rule above it
            break

    if hit >= barrier:
        return text, False

    lo = hit
    while lo - 1 >= 0 and is_fluff(lines[lo - 1]):
        lo -= 1
    hi = hit
    while hi + 1 < barrier and is_fluff(lines[hi + 1]):
        hi += 1
    del lines[lo:hi + 1]
    return "\n".join(lines), True


def strip_old_patch(text, new_funcs):
    """Remove a previously applied patch (rev 1 or rev 2) so it can be re-applied.

    Only ever deletes procedures the patch itself introduced -- never anything
    that belongs to the original macro.
    """
    removable = {n.lower() for n in REV1_NAMES}
    removable |= {p["name"].lower() for p in find_procs(new_funcs)}

    targets = [p for p in find_procs(text) if p["name"].lower() in removable]
    if not targets:
        return text, None

    removed = []
    for p in sorted(targets, key=lambda x: x["start"], reverse=True):
        start = p["start"]
        # swallow the contiguous comment/blank block sitting directly above
        head = text[:start]
        hl = head.split("\n")
        cut = len(hl) - 1                      # hl[-1] == "" (head ends in \n)
        while cut - 1 >= 0:
            t = hl[cut - 1].strip()
            if t == "" or t.startswith("'"):
                cut -= 1
            else:
                break
        start = len(("\n".join(hl[:cut]) + "\n") if cut > 0 else "")
        text = text[:start] + text[p["end"]:]
        removed.append(p["name"])

    for needle in ("END OF CMS ORIENTATION PATCH", "CMS ORIENTATION PATCH"):
        while True:
            text, hit = remove_comment_region_containing(text, needle)
            if not hit:
                break

    # rev-1 call sites -> rev-2 name
    n_calls = text.count(REV1_ROTATE)
    if n_calls:
        text = text.replace(REV1_ROTATE, ROTATE_FN)

    text = re.sub(r"\n{4,}", "\n\n\n", text)

    note = "removed %d old patch procedure(s)" % len(removed)
    if n_calls:
        note += ", renamed %d call site(s) to %s" % (n_calls, ROTATE_FN)
    return text, note


def edit_append_functions(text, new_funcs):
    """E5 - append the new procedures."""
    existing = {p["name"].lower() for p in find_procs(text)}
    incoming = [p["name"] for p in find_procs(new_funcs)]

    clash = [n for n in incoming if n.lower() in existing]
    if clash:
        return text, "SKIPPED - already defined in this module: %s" % ", ".join(clash)

    body = new_funcs.strip("\n") + "\n"
    m = END_OF_MODULE_RE.search(text)
    if m:
        out = text[:m.start()] + body + "\n" + text[m.start():]
        return out, "%d procedure(s) inserted before END OF MODULE" % len(incoming)
    return (text.rstrip("\n") + "\n\n" + body,
            "%d procedure(s) appended at end of file" % len(incoming))


# ---------------------------------------------------------------------------
# scan mode
# ---------------------------------------------------------------------------
def scan(folder, recursive):
    paths = []
    if recursive:
        for root, _dirs, files in os.walk(folder):
            paths += [os.path.join(root, f) for f in files
                      if f.lower().endswith(".bas")]
    else:
        paths = [os.path.join(folder, f) for f in sorted(os.listdir(folder))
                 if f.lower().endswith(".bas")]

    if not paths:
        print("No .bas files found in %s" % folder)
        return

    print("Scanning %d .bas file(s) in %s\n" % (len(paths), folder))
    hdr = "%-58s %7s %5s %5s %5s %5s  %s" % (
        "file", "lines", "PoJ", "Str", "Ori", "Face", "verdict")
    print(hdr)
    print("-" * len(hdr))

    ready = []
    for p in paths:
        try:
            text, _enc = read_text_lf(p)
        except SystemExit:
            continue
        procs = find_procs(text)
        names = {x["name"].lower() for x in procs}

        has_poj = "processonejob" in names
        has_str = "straightenassemblyfromplatetransform" in names
        has_ori = "applymoldorientationmatrix" in names

        has_face = False
        if has_poj:
            proc = get_proc(procs, "ProcessOneJob")
            body = text[proc["start"]:proc["end"]]
            opener = re.compile(r"(?i)^[ \t]*If[ \t]+%s[ \t]+Then[ \t]*$" % FLAG)
            fb = find_if_block(body, opener)
            has_face = bool(fb) and "EnsureCmsTopOrientationFromMatchedTcpBcp" in body

        rev1 = [n for n in REV1_NAMES if n.lower() in names]
        already = ROTATE_FN.lower() in names

        if already:
            verdict = "already patched (rev 2)"
        elif rev1:
            verdict = "HAS REV 1 - remove %s etc. first" % rev1[0]
        elif has_poj and has_str and has_ori and has_face:
            verdict = "PATCHABLE"
            ready.append(p)
        else:
            missing = []
            if not has_poj:
                missing.append("ProcessOneJob")
            if not has_str:
                missing.append("StraightenAssemblyFromPlateTransform")
            if not has_ori:
                missing.append("ApplyMoldOrientationMatrix")
            if not has_face:
                missing.append("the %s block" % FLAG)
            verdict = "older build - missing " + ", ".join(missing)

        print("%-58s %7d %5s %5s %5s %5s  %s" % (
            os.path.basename(p)[:58], text.count("\n") + 1,
            "y" if has_poj else "-", "y" if has_str else "-",
            "y" if has_ori else "-", "y" if has_face else "-", verdict))

    print()
    if ready:
        print("Patch one of these:")
        for p in ready:
            print('  python apply_cms_orientation_patch.py "%s"' % p)
    else:
        print("Nothing here is patchable. The build you want is the one that")
        print("already contains StraightenAssemblyFromPlateTransform AND")
        print("ApplyMoldOrientationMatrix AND %s." % FLAG)
        print("Export the live module out of the SolidWorks VBA editor")
        print("(right-click the module > Export File...) and scan that folder.")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Apply the CMS orientation patch (rev 2) to the export macro.")
    ap.add_argument("source", nargs="?", help="path to the .bas module")
    ap.add_argument("--scan", metavar="DIR",
                    help="report which .bas files in DIR can take the patch")
    ap.add_argument("--recursive", action="store_true",
                    help="with --scan, walk sub-folders too")
    ap.add_argument("--new-funcs", default=NEW_FUNCS_DEFAULT,
                    help="path to CMS_Orientation_Patch_NewFunctions.bas")
    ap.add_argument("--out", help="output path (default <source>_PATCHED.bas)")
    ap.add_argument("--in-place", action="store_true",
                    help="overwrite the source (a .bak is made first)")
    ap.add_argument("--allow-partial", action="store_true",
                    help="write the file even if some edits could not be applied")
    args = ap.parse_args()

    if args.scan:
        if not os.path.isdir(args.scan):
            die("not a folder: %s" % args.scan)
        scan(args.scan, args.recursive)
        return

    if not args.source:
        die("give a .bas path, or use --scan to find one:\n"
            "       python apply_cms_orientation_patch.py --scan C:\\CMS_AI")
    if not os.path.isfile(args.source):
        die("source file not found: %s" % args.source)
    if not os.path.isfile(args.new_funcs):
        die("new-functions file not found: %s\n"
            "       It should sit next to this script." % args.new_funcs)

    text, enc = read_text_lf(args.source)
    new_funcs, _ = read_text_lf(args.new_funcs)

    print("Patching: %s" % args.source)
    print("  %d lines, encoding %s\n" % (text.count("\n") + 1, enc))

    # An earlier patch (rev 1's leaf-part rotation, or an earlier rev 2) is
    # stripped out first so this run can put the current code in cleanly.
    text, r = strip_old_patch(text, new_funcs)
    if r:
        print("  OK  [upgrade] %s\n" % r)

    results = []

    text, r = edit_declare_flag(text);              results.append(("0 declare flag", r))
    text, r = edit_reset_flag(text);                results.append(("1 reset flag in ProcessOneJob", r))
    text, r = edit_rotation_loop(text, "StraightenAssemblyFromPlateTransform")
    results.append(("2 StraightenAssemblyFromPlateTransform loop", r))
    text, r = edit_rotation_loop(text, "ApplyMoldOrientationMatrix")
    results.append(("3 ApplyMoldOrientationMatrix loop", r))
    text, r = edit_faces_block(text);               results.append(("4 faces after BOM match", r))
    text, r = edit_append_functions(text, new_funcs)
    results.append(("5 append new procedures", r))

    failed = []
    for label, res in results:
        if res is None:
            print("  --  [%s] NOT FOUND" % label)
            failed.append(label)
        elif isinstance(res, str) and res.startswith("SKIPPED"):
            print("  --  [%s] %s" % (label, res))
            failed.append(label)
        else:
            print("  OK  [%s] %s" % (label, res))

    problems = verify_patch(text)
    if problems:
        print()
        die("verification failed, nothing written:\n       %s"
            % "\n       ".join(problems))
    print("  OK  [verify] every patched call references a real variable")

    if failed and not args.allow_partial:
        print()
        die("%d edit(s) could not be applied, so nothing was written:\n       %s\n\n"
            "       This module is probably an older build. Run:\n"
            "         python apply_cms_orientation_patch.py --scan <folder>\n"
            "       to find the build that has all the anchors, or pass\n"
            "       --allow-partial to write what did apply."
            % (len(failed), "\n       ".join(failed)))

    if args.in_place:
        stamp = datetime.datetime.now().strftime("%Y%m%d-%H%M%S")
        backup = "%s.%s.bak" % (args.source, stamp)
        shutil.copy2(args.source, backup)
        print("\nBackup written to:\n  %s" % backup)
        out_path = args.source
    else:
        out_path = args.out or (os.path.splitext(args.source)[0] + "_PATCHED.bas")

    write_text_crlf(out_path, text, enc)

    print("\nPatched file written to:\n  %s" % out_path)
    if failed:
        print("\nWARNING: written with %d edit(s) missing (--allow-partial)."
              % len(failed))
    print("\nNext: in the SolidWorks VBA editor remove the old module,")
    print("File > Import File... the patched .bas, then Debug > Compile.")


if __name__ == "__main__":
    main()
