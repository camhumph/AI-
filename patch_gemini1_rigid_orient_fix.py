"""
gemini1.bas -- FIX: rigid single rotation + stop the view code overriding it.

THREE BUGS THIS FIXES
    1. PARTS SHIFTED RELATIVE TO EACH OTHER (the scrambled screenshot).
       Both rotation passes used model.GetComponents(False), which returns
       components at ALL levels. A child of a sub-assembly then got rotated
       TWICE -- once by its own Transform2 and again inherited from its rotated
       parent. GetComponents(True) returns top-level only, so each component
       moves rigidly with its children and nothing shifts.

    2. "IT IS APPLYING THE FRONT AS THE TOP."
       The log shows exactly this:
           "Matched TOP/BOTTOM HOLDER selected *Front. Stack axis=Z"
           "Standard views REDEFINED: current TCP/top-side orientation
            assigned to *Top (ViewId 5)."
       EnsureCmsTopOrientationFromMatchedTcpBcp runs AFTER the straighten and
       overwrites the now-correct standard views with its own snap-to-nearest
       guess -- persisting *Front as *Top. Once the geometry is genuinely square
       and the right way up, that function must be SKIPPED and plain *Top used.

    3. TWO ROTATIONS + TWO RE-SCANS CORRUPTED THE EXPORT INDICES.
       Pass 2 used to run after BOM matching and then re-scan, which re-sorts
       parts by volume and invalidates every ExportRows().CadPartIndex the
       matcher had just stored -- so exports could pull the wrong body. Now both
       rotations are computed and applied as ONE matrix, once, BEFORE any
       matching, with a single re-scan. Nothing downstream sees stale indices.

HOW UP/FRONT ARE FOUND WITHOUT THE BOM
    Running before matching means no TCP/pot names. Pure geometry instead:
      - The two largest-footprint parts are the clamp plates. Their centre-to-
        centre vector is the stack axis. Sign: toward the LIGHTER one, which is
        the shop's existing TCP=LIGHT / BCP=HEAVY convention (the log's own
        "TCP/BCP mass match: TCP preference=LIGHT" line).
      - Holders and pots are the chunky blocks (thickness >= 3"). Of those the
        two with the LARGEST footprint are the holders, the two smallest are the
        pots. front = mean(pots) - mean(holders), perpendicularised to up.
    Both vectors are snapped to signed axes, so the combined result is an exact
    signed permutation once the plate transform has squared things.

USAGE
    python patch_gemini1_rigid_orient_fix.py "C:\\CMS_Local_Workspace\\gemini1.bas"
    --dry-run

Requires the two earlier patches. Idempotent. Timestamped .bak.
"""

import shutil
import sys
import time
from pathlib import Path

ENCODING = "latin-1"
MARKER = "ComputeMoldUpAndFrontFromGeometry"

# ---------------------------------------------------------------------------
# FIX 1 -- top-level components only, in BOTH passes.
# ---------------------------------------------------------------------------
FIX1_OLD = "    vComps = model.GetComponents(False)\n"
FIX1_NEW = (
    "    ' TOP-LEVEL ONLY. GetComponents(False) returns every level, so a child\n"
    "    ' of a sub-assembly was rotated twice -- once by its own transform and\n"
    "    ' again inherited from its rotated parent. That is what scrambled the\n"
    "    ' assembly. True keeps each component rigid with its children.\n"
    "    vComps = model.GetComponents(True)\n"
)

# ---------------------------------------------------------------------------
# FIX 2 -- geometry-only up/front, and fold pass 2 into pass 1.
# ---------------------------------------------------------------------------
EDIT_ANCHOR = "Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean\n"

NEW_CODE = '''\
' Mold UP and FRONT from geometry alone -- no BOM, no component names.
'
' Must be able to run BEFORE BOM matching, because rotating after the match
' would invalidate every ExportRows().CadPartIndex it just stored.
'
'   up    : the two largest-footprint parts are the clamp plates; the stack axis
'           is the vector between their centres, pointing toward the LIGHTER one
'           (the shop's TCP=LIGHT / BCP=HEAVY convention).
'   front : holders and pots are the chunky blocks (T >= 3"). Of those, the two
'           biggest footprints are the holders and the two smallest are the
'           pots, so front = mean(pots) - mean(holders).
'
' Returns False if the mold does not look like a pot-block base.
Private Function ComputeMoldUpAndFrontFromGeometry(ByRef ux As Double, ByRef uy As Double, ByRef uz As Double, _
                                                  ByRef fx As Double, ByRef fy As Double, ByRef fz As Double) As Boolean
On Error GoTo eh

    ComputeMoldUpAndFrontFromGeometry = False
    If PartCount < 4 Then Exit Function

    ' ---- clamp plates: two largest footprints -------------------------
    Dim i As Long
    Dim p1 As Long, p2 As Long
    Dim f1 As Double, f2 As Double
    Dim fp As Double
    p1 = 0: p2 = 0: f1 = -1#: f2 = -1#

    For i = 1 To PartCount
        fp = parts(i).Width * parts(i).Length
        If fp > f1 Then
            f2 = f1: p2 = p1
            f1 = fp: p1 = i
        ElseIf fp > f2 Then
            f2 = fp: p2 = i
        End If
    Next i

    If p1 < 1 Or p2 < 1 Then Exit Function

    ' Point up toward the lighter plate = TCP.
    Dim hiIdx As Long, loIdx As Long
    If parts(p1).massValue <= parts(p2).massValue Then
        loIdx = p1: hiIdx = p2
    Else
        loIdx = p2: hiIdx = p1
    End If

    ux = parts(loIdx).AsmCenterX - parts(hiIdx).AsmCenterX
    uy = parts(loIdx).AsmCenterY - parts(hiIdx).AsmCenterY
    uz = parts(loIdx).AsmCenterZ - parts(hiIdx).AsmCenterZ

    If SnapVectorToSignedAxis(ux, uy, uz) = False Then Exit Function

    LogLine "ORIENT geometry: clamp plates idx " & CStr(loIdx) & " (lighter=TCP, " & _
            FormatNumberForCsv(parts(loIdx).massValue) & " lb) and " & CStr(hiIdx) & _
            " (heavier=BCP, " & FormatNumberForCsv(parts(hiIdx).massValue) & " lb)"

    ' ---- holders vs pots: chunky blocks by footprint -------------------
    Dim chunk() As Long
    Dim nChunk As Long
    ReDim chunk(1 To PartCount)
    nChunk = 0

    For i = 1 To PartCount
        If i <> p1 And i <> p2 Then
            If parts(i).Thickness >= 3# Then
                nChunk = nChunk + 1
                chunk(nChunk) = i
            End If
        End If
    Next i

    If nChunk >= 4 Then
        ' Sort chunky blocks by footprint, descending.
        Dim a As Long, b As Long, t As Long
        For a = 1 To nChunk - 1
            For b = a + 1 To nChunk
                If (parts(chunk(b)).Width * parts(chunk(b)).Length) > _
                   (parts(chunk(a)).Width * parts(chunk(a)).Length) Then
                    t = chunk(a): chunk(a) = chunk(b): chunk(b) = t
                End If
            Next b
        Next a

        Dim hX As Double, hY As Double, hZ As Double
        Dim pX As Double, pY As Double, pZ As Double

        ' Two biggest = holders, two smallest = pots.
        hX = (parts(chunk(1)).AsmCenterX + parts(chunk(2)).AsmCenterX) / 2#
        hY = (parts(chunk(1)).AsmCenterY + parts(chunk(2)).AsmCenterY) / 2#
        hZ = (parts(chunk(1)).AsmCenterZ + parts(chunk(2)).AsmCenterZ) / 2#

        pX = (parts(chunk(nChunk)).AsmCenterX + parts(chunk(nChunk - 1)).AsmCenterX) / 2#
        pY = (parts(chunk(nChunk)).AsmCenterY + parts(chunk(nChunk - 1)).AsmCenterY) / 2#
        pZ = (parts(chunk(nChunk)).AsmCenterZ + parts(chunk(nChunk - 1)).AsmCenterZ) / 2#

        fx = pX - hX: fy = pY - hY: fz = pZ - hZ

        LogLine "ORIENT geometry: holders idx " & CStr(chunk(1)) & "," & CStr(chunk(2)) & _
                "  pots idx " & CStr(chunk(nChunk - 1)) & "," & CStr(chunk(nChunk))
    Else
        LogLine "ORIENT geometry: only " & CStr(nChunk) & " chunky block(s); " & _
                "front will use a synthetic perpendicular."
        fx = 0#: fy = 0#: fz = 0#
    End If

    ' Strip the up component, then snap.
    Dim dp As Double
    dp = fx * ux + fy * uy + fz * uz
    fx = fx - dp * ux: fy = fy - dp * uy: fz = fz - dp * uz

    If Sqr(fx * fx + fy * fy + fz * fz) < 0.001 Then
        If Abs(ux) < 0.5 Then
            fx = 1#: fy = 0#: fz = 0#
        ElseIf Abs(uy) < 0.5 Then
            fx = 0#: fy = 1#: fz = 0#
        Else
            fx = 0#: fy = 0#: fz = 1#
        End If
        dp = fx * ux + fy * uy + fz * uz
        fx = fx - dp * ux: fy = fy - dp * uy: fz = fz - dp * uz
    End If

    If SnapVectorToSignedAxis(fx, fy, fz) = False Then Exit Function

    ComputeMoldUpAndFrontFromGeometry = True
    Exit Function

eh:
    ComputeMoldUpAndFrontFromGeometry = False
End Function

'''

# ---------------------------------------------------------------------------
# FIX 3 -- fold the orient into the straighten so it is ONE rotation, and skip
#          the view code that was overwriting it.
# ---------------------------------------------------------------------------
FIX3_OLD = '''    If StraightenAssemblyFromPlateTransform(swModel) Then
        LogStart "Re-scan CAD after straightening"
        ScanActiveSolidWorksDocument
        SortPartsByVolumeDescending
        LogLine "CAD PartCount=" & PartCount & " (after straightening)"
        LogDone "Re-scan CAD after straightening"
    End If
'''

FIX3_NEW = '''    If StraightenAssemblyFromPlateTransform(swModel) Then
        LogStart "Re-scan CAD after straightening"
        ScanActiveSolidWorksDocument
        SortPartsByVolumeDescending
        LogLine "CAD PartCount=" & PartCount & " (after straightening)"
        LogDone "Re-scan CAD after straightening"

        ' Geometry is square now. Second and FINAL rotation: put it the right
        ' way up as an exact signed axis permutation. Runs here, before BOM
        ' matching, so no ExportRows CadPartIndex can be invalidated later.
        Dim gUx As Double, gUy As Double, gUz As Double
        Dim gFx As Double, gFy As Double, gFz As Double

        If ComputeMoldUpAndFrontFromGeometry(gUx, gUy, gUz, gFx, gFy, gFz) Then
            If ApplyMoldOrientationMatrix(swModel, gUx, gUy, gUz, gFx, gFy, gFz) Then
                gMoldGeometryIsSquareAndUpright = True
                LogStart "Re-scan CAD after TCP-top / pot-front orientation"
                ScanActiveSolidWorksDocument
                SortPartsByVolumeDescending
                LogLine "CAD PartCount=" & PartCount & " (upright)"
                LogDone "Re-scan CAD after TCP-top / pot-front orientation"
            End If
        Else
            LogLine "ORIENT: geometry did not look like a pot-block base; " & _
                    "leaving the existing view logic in charge."
        End If
    End If
'''

# The applier + the flag.
APPLIER = '''\
' True once the geometry has been physically squared AND turned the right way
' up. When set, EnsureCmsTopOrientationFromMatchedTcpBcp must NOT run: its
' snap-to-nearest-view logic would overwrite correct standard views with a
' guess. The J8493 log caught it doing exactly that -- "selected *Front" then
' "assigned to *Top".
Private gMoldGeometryIsSquareAndUpright As Boolean

' Rotate the assembly so up -> +Y and front -> +Z, as one rigid transform.
Private Function ApplyMoldOrientationMatrix(ByVal model As Object, _
                                            ByVal ux As Double, ByVal uy As Double, ByVal uz As Double, _
                                            ByVal fx As Double, ByVal fy As Double, ByVal fz As Double) As Boolean
On Error GoTo eh

    ApplyMoldOrientationMatrix = False

    Dim rx As Double, ry As Double, rz As Double
    rx = uy * fz - uz * fy
    ry = uz * fx - ux * fz
    rz = ux * fy - uy * fx

    Dim R(0 To 2, 0 To 2) As Double
    R(0, 0) = rx: R(0, 1) = ry: R(0, 2) = rz
    R(1, 0) = ux: R(1, 1) = uy: R(1, 2) = uz
    R(2, 0) = fx: R(2, 1) = fy: R(2, 2) = fz

    LogLine "ORIENT: up=" & FormatNumberForCsv(ux) & "," & FormatNumberForCsv(uy) & _
            "," & FormatNumberForCsv(uz) & " -> +Y   front=" & _
            FormatNumberForCsv(fx) & "," & FormatNumberForCsv(fy) & "," & _
            FormatNumberForCsv(fz) & " -> +Z"

    If uy > 0.999 And fz > 0.999 Then
        LogLine "ORIENT: already TCP-up and pots-front; nothing to rotate."
        ApplyMoldOrientationMatrix = True
        Exit Function
    End If

    Dim swMathUtil As Object
    Set swMathUtil = swApp.GetMathUtility
    If swMathUtil Is Nothing Then Exit Function

    Dim vComps As Variant
    vComps = model.GetComponents(True)
    If IsEmpty(vComps) Then Exit Function

    Dim moved As Long, failed As Long
    Dim ci As Long
    moved = 0: failed = 0

    For ci = 0 To UBound(vComps)
        If Not vComps(ci) Is Nothing Then
            If StraightenOneComponentByMatrix(vComps(ci), R, swMathUtil) Then
                moved = moved + 1
            Else
                failed = failed + 1
            End If
        End If
    Next ci

    LogLine "ORIENT: rotated " & CStr(moved) & " top-level component(s), " & _
            CStr(failed) & " failed."

    If moved = 0 Then Exit Function

    On Error Resume Next
    model.EditRebuild3
    model.GraphicsRedraw2
    On Error GoTo eh

    ApplyMoldOrientationMatrix = True
    Exit Function

eh:
    LogLine "ApplyMoldOrientationMatrix error: " & Err.Description
    ApplyMoldOrientationMatrix = False
End Function

'''

# ---------------------------------------------------------------------------
# FIX 4 -- bypass the view code that overwrites the correct orientation.
# ---------------------------------------------------------------------------
FIX4_OLD = '''    LogStart "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
'''

FIX4_NEW = '''    LogStart "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    If gMoldGeometryIsSquareAndUpright Then
        ' The geometry is already square and TCP-up / pots-front, so the plain
        ' standard views are correct. Do NOT run
        ' EnsureCmsTopOrientationFromMatchedTcpBcp here: its snap-to-nearest
        ' logic would redefine *Top from whichever view it thinks is closest,
        ' and the J8493 log shows it choosing *Front and persisting THAT as
        ' *Top. Just show the true top and persist it as-is.
        LogLine "Geometry already square and upright: using plain *Top, " & _
                "skipping snap-based view redefinition."
        swModel.ShowNamedView2 "*Top", 5
        StabilizeActiveView swModel, 100
        PersistCurrentViewAsStandardTop swModel
    Else
        EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
    End If
'''


def apply_once(text, old, new, label, changed, skipped, already_token=None):
    if already_token and already_token in text:
        print(f"  skip  {label}")
        return text, changed, skipped + 1
    n = text.count(old)
    if n < 1:
        raise SystemExit(f"ABORT: anchor for {label} not found. Nothing written.")
    text = text.replace(old, new)
    print(f"  ok    {label}" + (f"  ({n} sites)" if n > 1 else ""))
    return text, changed + 1, skipped


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in {a for a in sys.argv[1:] if a.startswith("--")}
    if not args:
        raise SystemExit(__doc__)

    path = Path(args[0])
    if not path.is_file():
        raise SystemExit(f"Not a file: {path}")

    original = path.read_text(encoding=ENCODING)
    text = original
    print(f"Read {path}\n     {len(original):,} chars, {original.count(chr(10)):,} lines\n")

    if "StraightenAssemblyFromPlateTransform" not in text:
        raise SystemExit(
            "ABORT: run patch_gemini1_straighten_first.py first. Nothing written."
        )

    changed = skipped = 0

    # 1. top-level components only, everywhere
    if "GetComponents(False)" in text:
        text, changed, skipped = apply_once(
            text, FIX1_OLD, FIX1_NEW, "1/5 GetComponents(False) -> (True)", changed, skipped
        )
    else:
        print("  skip  1/5 already top-level only")
        skipped += 1

    # 2. geometry-only up/front
    text, changed, skipped = apply_once(
        text, EDIT_ANCHOR, NEW_CODE + EDIT_ANCHOR,
        "2/5 inserted ComputeMoldUpAndFrontFromGeometry", changed, skipped,
        already_token=MARKER + "(ByRef ux As Double",
    )

    # 3. applier + flag
    text, changed, skipped = apply_once(
        text, EDIT_ANCHOR, APPLIER + EDIT_ANCHOR,
        "3/5 inserted ApplyMoldOrientationMatrix + flag", changed, skipped,
        already_token="ApplyMoldOrientationMatrix(ByVal model As Object",
    )

    # 4. fold orient into the straighten block
    text, changed, skipped = apply_once(
        text, FIX3_OLD, FIX3_NEW,
        "4/5 one rotation, one re-scan, before BOM matching", changed, skipped,
        already_token="If ComputeMoldUpAndFrontFromGeometry(gUx",
    )

    # 5. bypass the overriding view code
    text, changed, skipped = apply_once(
        text, FIX4_OLD, FIX4_NEW,
        "5/5 skip snap-based view redefinition when upright", changed, skipped,
        already_token="If gMoldGeometryIsSquareAndUpright Then",
    )

    print()
    if changed == 0:
        print(f"Nothing to do -- all {skipped} edit(s) already applied.")
        return
    if dry:
        print(f"--dry-run: {changed} edit(s) would be applied. Nothing written.")
        return

    backup = path.with_suffix(path.suffix + f".bak_{time.strftime('%Y%m%d_%H%M%S')}")
    shutil.copy2(path, backup)
    path.write_text(text, encoding=ENCODING)

    print(f"Backup:  {backup}")
    print(f"Patched: {path}  ({len(text) - len(original):+,} chars)")
    print()
    print("Re-run J8493. Expect ONE rotation event, then ONE re-scan:")
    print()
    print("  STRAIGHTEN: rotated 42 top-level component(s), 0 failed.")
    print("  >>> START: Re-scan CAD after straightening")
    print("  ORIENT geometry: clamp plates idx 1 (lighter=TCP, 48.299 lb) and 2 ...")
    print("  ORIENT geometry: holders idx 3,4  pots idx 5,6")
    print("  ORIENT: up=0.0000,1.0000,0.0000 -> +Y   front=0.0000,0.0000,1.0000 -> +Z")
    print("  ORIENT: rotated 42 top-level component(s), 0 failed.")
    print("  Geometry already square and upright: using plain *Top, skipping")
    print("    snap-based view redefinition.")
    print()
    print("The line that must NOT appear any more is:")
    print("  'Matched TOP/BOTTOM HOLDER selected *Front ... assigned to *Top'")
    print("That was the front being persisted as the top.")
    print()
    print("If parts still shift, the components are nested deeper than one level")
    print("or are mated. Send me the log and I'll switch to floating them first.")


if __name__ == "__main__":
    main()
