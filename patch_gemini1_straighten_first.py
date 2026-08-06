"""
gemini1.bas -- STRAIGHTEN THE ASSEMBLY FIRST, from a plate's placement transform.

WHAT THE J8493 LOG PROVED
    1. My previous hook was in the WRONG FUNCTION. These lines
           "STRAIGHTEN skipped: TCP or BCP component not found by name."
       only appear at 14:33:27 onward -- during the J BLOCK / PULLCORE packages,
       on one-part temp assemblies (CAM.sldasm, PULLCORE.sldasm) where TCP
       genuinely does not exist. The real orientation ran at 14:29:37 via
       EnsureCmsTopOrientationFromMatchedTcpBcp, which never calls
       TryShowTcpTopViewFromComponentCenters. So nothing I patched ever executed
       on the mold.

    2. The tilt, measured from the log:
           "Matched TOP/BOTTOM HOLDER TCP-BCP separation: X=0.000 Y=4.014 Z=5.524"
       normalised = (0, 0.588, 0.809) -> 54 degrees off axis, sitting between Y
       and Z. The code then chose "Stack axis=Z" because 5.524 > 4.014. That is
       snap-to-nearest-axis failing by 54 degrees, and it is why every
       subsequent pot-front check failed (front delta -4.809, flipped, still
       failed): depth was being measured along a meaningless direction.

       The final STL matrix [1,0,0; 0,0,-1; 0,1,0] is a clean 90-degree
       rotation with no 54-degree component -- nothing ever corrected the tilt.

THE FIX (yours)
    Pick a plate, get its rotation, rotate the whole assembly by the inverse,
    FIRST -- before anything reads a centre or a bounding box. Then TCP/front/
    pots all run on square geometry.

    Not a fitted bounding box. The same source the pullcore code trusts:
    Component2.Transform2. Per the comment at line 9283, ArrayData holds the
    part's local X/Y/Z axes as COLUMNS, expressed in assembly coords -- that IS
    the part's rotation, exactly. For an orthonormal rotation the inverse is the
    transpose, so:

        R rows = the plate's local X, Y, Z axes

    Apply R to every component and the reference plate lands perfectly on the
    model axes (R * M = M_transpose * M = Identity). Every other plate is
    parallel to it, so the whole base comes square.

    No names needed, no BOM needed -- which matters, because
    FindComponentByKeys fails on this customer's "...sldasm-Part-36-1" naming
    and that killed all three of my earlier attempts.

USAGE
    python patch_gemini1_straighten_first.py "C:\\CMS_Local_Workspace\\gemini1.bas"
    --dry-run

Idempotent per edit. Timestamped .bak before writing.
"""

import shutil
import sys
import time
from pathlib import Path

ENCODING = "latin-1"
MARKER = "StraightenAssemblyFromPlateTransform"
HELPER = "StraightenOneComponentByMatrix"

# ---------------------------------------------------------------------------
# EDIT 1 -- insert the functions before ScanActiveSolidWorksDocument's caller
#           neighbourhood. We anchor on a stable private function definition.
# ---------------------------------------------------------------------------
EDIT_1_ANCHOR = "Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean\n"

NEW_CODE = '''\
' ============================================================================
' STRAIGHTEN THE WHOLE ASSEMBLY FROM A PLATE'S PLACEMENT TRANSFORM
'
' Runs immediately after the CAD scan, before ANY orientation / centre / bbox
' logic. Once this succeeds the geometry is square to model XYZ, so *Top really
' is top, and the TCP/front/pot passes all work on sane numbers.
'
' WHY A TRANSFORM AND NOT A BOUNDING BOX
'   An axis-aligned bbox of a tilted part tells you nothing about the tilt, and
'   a fitted bbox is an expensive search that can pick the wrong face. The
'   placement transform IS the rotation, exactly, for free -- the same source
'   TryGetComponentAssemblyPullAngleDeg uses for pullcore de-rotation.
'
'   Per the convention documented there: ArrayData holds the part's local X/Y/Z
'   axes as COLUMNS in assembly coords. For an orthonormal rotation the inverse
'   is the transpose, so R rows = those local axes. Applying R to every
'   component makes the reference plate axis-aligned (R * M = Identity) and
'   carries every parallel plate with it.
'
' WHY THE LARGEST-FOOTPRINT PART
'   That is always a clamp plate on a mold base, it is unambiguous, and it needs
'   no BOM and no name matching -- which is essential here, because
'   FindComponentByKeys returns Nothing on "...sldasm-Part-36-1" naming.
Private Function StraightenAssemblyFromPlateTransform(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    StraightenAssemblyFromPlateTransform = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function
    If PartCount < 1 Then Exit Function

    ' ---- 1. pick the reference plate: biggest footprint ------------------
    Dim i As Long
    Dim bestIdx As Long
    Dim bestFoot As Double
    Dim fp As Double

    bestIdx = 0
    bestFoot = 0#

    For i = 1 To PartCount
        fp = parts(i).Width * parts(i).Length
        If fp > bestFoot Then
            bestFoot = fp
            bestIdx = i
        End If
    Next i

    If bestIdx < 1 Then
        LogLine "STRAIGHTEN: no reference plate found."
        Exit Function
    End If

    LogLine "STRAIGHTEN reference plate: " & parts(bestIdx).componentName & _
            "  L/W/T=" & FormatNumberForCsv(parts(bestIdx).Length) & "/" & _
            FormatNumberForCsv(parts(bestIdx).Width) & "/" & _
            FormatNumberForCsv(parts(bestIdx).Thickness)

    ' ---- 2. find that component and read its rotation -------------------
    Dim vComps As Variant
    vComps = model.GetComponents(False)
    If IsEmpty(vComps) Then
        LogLine "STRAIGHTEN: GetComponents returned nothing."
        Exit Function
    End If

    Dim refComp As Object
    Set refComp = Nothing

    Dim wantName As String
    wantName = UCase(Trim(parts(bestIdx).componentName))

    Dim ci As Long
    For ci = 0 To UBound(vComps)
        If Not vComps(ci) Is Nothing Then
            If UCase(Trim(vComps(ci).Name2)) = wantName Then
                Set refComp = vComps(ci)
                Exit For
            End If
        End If
    Next ci

    If refComp Is Nothing Then
        ' Fall back on a suffix match -- Name2 sometimes carries a config or
        ' parent prefix the scan stripped.
        For ci = 0 To UBound(vComps)
            If Not vComps(ci) Is Nothing Then
                If InStr(wantName, UCase(Trim(vComps(ci).Name2))) > 0 Or _
                   InStr(UCase(Trim(vComps(ci).Name2)), wantName) > 0 Then
                    Set refComp = vComps(ci)
                    Exit For
                End If
            End If
        Next ci
    End If

    If refComp Is Nothing Then
        LogLine "STRAIGHTEN: could not locate the reference plate component."
        Exit Function
    End If

    Dim xf As Object
    Set xf = refComp.Transform2
    If xf Is Nothing Then
        LogLine "STRAIGHTEN: reference plate has no Transform2."
        Exit Function
    End If

    Dim v As Variant
    v = xf.ArrayData
    If IsEmpty(v) Then Exit Function
    If IsArray(v) = False Then Exit Function
    If UBound(v) < 8 Then Exit Function

    ' ---- 3. R = transpose of the plate rotation -------------------------
    ' ArrayData columns are the local axes, so the transpose's ROWS are those
    ' same axes. R * M = Identity, i.e. the plate ends up axis-aligned.
    Dim R(0 To 2, 0 To 2) As Double
    R(0, 0) = CDbl(v(0)): R(0, 1) = CDbl(v(1)): R(0, 2) = CDbl(v(2))
    R(1, 0) = CDbl(v(3)): R(1, 1) = CDbl(v(4)): R(1, 2) = CDbl(v(5))
    R(2, 0) = CDbl(v(6)): R(2, 1) = CDbl(v(7)): R(2, 2) = CDbl(v(8))

    ' How far off axis were we? Report it, so the log proves what happened.
    Dim offAxis As Double
    offAxis = Abs(Abs(R(0, 0)) - 1#) + Abs(Abs(R(1, 1)) - 1#) + Abs(Abs(R(2, 2)) - 1#)

    LogLine "STRAIGHTEN plate rotation (local axes in assembly coords):"
    LogLine "  local X = " & FormatNumberForCsv(R(0, 0)) & ", " & _
            FormatNumberForCsv(R(0, 1)) & ", " & FormatNumberForCsv(R(0, 2))
    LogLine "  local Y = " & FormatNumberForCsv(R(1, 0)) & ", " & _
            FormatNumberForCsv(R(1, 1)) & ", " & FormatNumberForCsv(R(1, 2))
    LogLine "  local Z = " & FormatNumberForCsv(R(2, 0)) & ", " & _
            FormatNumberForCsv(R(2, 1)) & ", " & FormatNumberForCsv(R(2, 2))
    LogLine "  off-axis measure = " & FormatNumberForCsv(offAxis)

    If offAxis < 0.0001 Then
        LogLine "STRAIGHTEN: assembly is already square to model XYZ; nothing to do."
        StraightenAssemblyFromPlateTransform = True
        Exit Function
    End If

    ' ---- 4. apply R to every component ---------------------------------
    Dim swMathUtil As Object
    Set swMathUtil = swApp.GetMathUtility
    If swMathUtil Is Nothing Then
        LogLine "STRAIGHTEN failed: GetMathUtility returned Nothing."
        Exit Function
    End If

    Dim moved As Long, failed As Long
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

    LogLine "STRAIGHTEN: rotated " & CStr(moved) & " component(s), " & _
            CStr(failed) & " failed."

    If moved = 0 Then
        LogLine "STRAIGHTEN failed: no component transforms could be set. " & _
                "Components are probably fixed -- they need floating first."
        Exit Function
    End If

    On Error Resume Next
    model.EditRebuild3
    model.GraphicsRedraw2
    On Error GoTo ErrHandler

    LogLine "STRAIGHTEN: assembly is now square to model XYZ. " & _
            "*Top is a true top view. Re-scanning so cached dims match."

    StraightenAssemblyFromPlateTransform = True
    Exit Function

ErrHandler:
    LogLine "StraightenAssemblyFromPlateTransform error: " & Err.Description
    StraightenAssemblyFromPlateTransform = False
End Function


' Rotate one component by R, working directly on Transform2.ArrayData.
'
' Per the convention at TryGetComponentAssemblyPullAngleDeg:
'   v(0..2)/v(3..5)/v(6..8) = local X/Y/Z axes in assembly coords (columns)
'   v(9..11) = translation, v(12) = scale
' so new column = R * old column, and new translation = R * old translation.
'
' Done as plain arithmetic rather than IMathTransform.Multiply, so there is no
' operand-order convention to get wrong.
Private Function StraightenOneComponentByMatrix(ByVal swComp As Object, _
                                                ByRef R() As Double, _
                                                ByVal swMathUtil As Object) As Boolean
On Error GoTo eh

    StraightenOneComponentByMatrix = False

    Dim xform As Object
    Set xform = swComp.Transform2
    If xform Is Nothing Then Exit Function

    Dim v As Variant
    v = xform.ArrayData
    If IsEmpty(v) Then Exit Function
    If IsArray(v) = False Then Exit Function
    If UBound(v) < 12 Then Exit Function

    Dim o(0 To 15) As Double
    Dim k As Long
    For k = 0 To 15
        If k <= UBound(v) Then o(k) = CDbl(v(k)) Else o(k) = 0#
    Next k

    Dim n(0 To 15) As Double

    Dim col As Long
    Dim base As Long
    For col = 0 To 3
        If col = 3 Then base = 9 Else base = col * 3
        n(base + 0) = R(0, 0) * o(base + 0) + R(0, 1) * o(base + 1) + R(0, 2) * o(base + 2)
        n(base + 1) = R(1, 0) * o(base + 0) + R(1, 1) * o(base + 1) + R(1, 2) * o(base + 2)
        n(base + 2) = R(2, 0) * o(base + 0) + R(2, 1) * o(base + 1) + R(2, 2) * o(base + 2)
    Next col

    n(12) = o(12)
    If n(12) = 0# Then n(12) = 1#
    n(13) = 0#: n(14) = 0#: n(15) = 0#

    Dim newX As Object
    Set newX = swMathUtil.CreateTransform(n)
    If newX Is Nothing Then Exit Function

    swComp.Transform2 = newX

    StraightenOneComponentByMatrix = True
    Exit Function

eh:
    StraightenOneComponentByMatrix = False
End Function

'''


# ---------------------------------------------------------------------------
# EDIT 2 -- call it the instant the scan finishes, then re-scan.
#
# This is the anchor that matters. The previous patch hooked
# TryShowTcpTopViewFromComponentCenters, which the mold's orientation path never
# calls -- proven by the log, where STRAIGHTEN only ever appeared while
# processing the J BLOCK and PULLCORE temp assemblies.
# ---------------------------------------------------------------------------
EDIT_2_OLD = '''    LogLine "CAD PartCount=" & PartCount
    LogDone "Scan CAD"
'''

EDIT_2_NEW = '''    LogLine "CAD PartCount=" & PartCount
    LogDone "Scan CAD"

    ' ------------------------------------------------------------------
    ' STRAIGHTEN FIRST. Before any orientation, centre or bbox is read.
    '
    ' The J8493 import sat 54 degrees off axis -- the log's own
    ' "TCP-BCP separation: X=0.000 Y=4.014 Z=5.524" normalises to
    ' (0, 0.588, 0.809) -- so every standard view was crooked and the
    ' snap-to-nearest-axis logic chose Z when the truth was 54 degrees
    ' between Y and Z. Squaring the geometry here makes every later pass
    ' (TCP top, holder long side, pots-in-front) work on sane numbers.
    '
    ' Geometry moves, so every cached centre and bbox from the scan above
    ' is stale -- re-scan immediately.
    ' ------------------------------------------------------------------
    If StraightenAssemblyFromPlateTransform(swModel) Then
        LogStart "Re-scan CAD after straightening"
        ScanActiveSolidWorksDocument
        SortPartsByVolumeDescending
        LogLine "CAD PartCount=" & PartCount & " (after straightening)"
        LogDone "Re-scan CAD after straightening"
    End If
'''


def main() -> None:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    flags = {a for a in sys.argv[1:] if a.startswith("--")}
    dry = "--dry-run" in flags

    if not args:
        raise SystemExit(__doc__)

    path = Path(args[0])
    if not path.is_file():
        raise SystemExit(f"Not a file: {path}")

    original = path.read_text(encoding=ENCODING)
    text = original
    print(f"Read {path}")
    print(f"     {len(original):,} chars, {original.count(chr(10)):,} lines\n")

    changed = skipped = 0

    # Insert the new functions. If the older patch already added
    # StraightenOneComponentByMatrix, strip it from our block to avoid a
    # duplicate-definition compile error.
    block = NEW_CODE
    if HELPER + "(ByVal swComp As Object" in text:
        cut = block.find("' Rotate one component by R")
        if cut > 0:
            block = block[:cut]
            print(f"  note  {HELPER} already exists -- inserting driver only")

    if MARKER + "(ByVal model As Object) As Boolean" in text:
        print("  skip  1/2 straighten function already present")
        skipped += 1
    else:
        n = text.count(EDIT_1_ANCHOR)
        if n != 1:
            raise SystemExit(f"ABORT: edit-1 anchor matched {n} times, expected 1. Nothing written.")
        text = text.replace(EDIT_1_ANCHOR, block + EDIT_1_ANCHOR, 1)
        print(f"  ok    1/2 inserted {MARKER}")
        changed += 1

    if "If " + MARKER + "(swModel) Then" in text:
        print("  skip  2/2 already called after Scan CAD")
        skipped += 1
    else:
        n = text.count(EDIT_2_OLD)
        if n != 1:
            raise SystemExit(
                f"ABORT: edit-2 anchor matched {n} times, expected 1.\n"
                f"Looking for the two lines after ScanActiveSolidWorksDocument:\n"
                f"{EDIT_2_OLD}Nothing written."
            )
        text = text.replace(EDIT_2_OLD, EDIT_2_NEW, 1)
        print("  ok    2/2 hooked in right after Scan CAD (+ re-scan)")
        changed += 1

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
    print("Re-run J8493. Right after 'DONE : Scan CAD' you should now see:")
    print()
    print("  STRAIGHTEN reference plate: ...Part-36-1  L/W/T=18.375/15.875/1.375")
    print("  STRAIGHTEN plate rotation (local axes in assembly coords):")
    print("    local X = 1.0000, 0.0000, 0.0000")
    print("    local Y = 0.0000, 0.5878, 0.8090      <- the 54 degree tilt")
    print("    local Z = 0.0000, -0.8090, 0.5878")
    print("    off-axis measure = 0.8244")
    print("  STRAIGHTEN: rotated 42 component(s), 0 failed.")
    print("  >>> START: Re-scan CAD after straightening")
    print()
    print("Then, further down, the payoff -- this line should become nearly")
    print("axis-clean instead of Y=4.014 Z=5.524:")
    print("  Matched TOP/BOTTOM HOLDER TCP-BCP separation: X=0.000 Y=6.828 Z=0.000")
    print("and the pot-front check should PASS instead of warning twice.")
    print()
    print("IF IT SAYS 'no component transforms could be set':")
    print("  the components are fixed. Tell me and I'll add a float pass")
    print("  (select all + UnfixComponent) ahead of the rotation.")


if __name__ == "__main__":
    main()
