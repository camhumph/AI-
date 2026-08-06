"""
Apply STRAIGHTEN-WHOLE-ASSEMBLY to gemini1.bas.

THE IDEA (yours)
    The pullcore logic does not fight the view -- it MEASURES the tilt and
    applies the inverse to the output (PULLCORE_STRAIGHTEN_SIGN *
    pcDetectedDeg, line ~8549). Do the same thing, but for the whole assembly
    and in 3D: measure the mold's true axes, then rigidly rotate every
    component so those axes land on the model's XYZ.

    Once the assembly is physically square, *Top IS top, *Front IS front, and
    every downstream output -- DXF, ISO JPG, STL, X_T -- is correct with no
    per-output correction. That is strictly better than my Orientation3
    attempt, which only fixed the camera and left the geometry crooked.

THE MATH
    Your own comment at line 9283 documents the convention:
        Transform2.ArrayData = 3x3 rotation in the first 9 entries, COLUMN-major
        (cols are the part's local X,Y,Z axes in assembly coords),
        v(9..11) = translation, v(12) = scale.

    So for local point p:  assembly = M*p + t,  M columns = v(0:3), v(3:6), v(6:9).

    To rotate the whole assembly by R about the origin:
        new column i = R * old column i
        new t        = R * old t

    Pure arithmetic on ArrayData -- no IMathTransform.Multiply, so there is no
    multiply-order convention to get wrong.

    R maps assembly -> mold frame. Its ROWS are the mold's axes in assembly
    coords:  row0 = right, row1 = up, row2 = front.
    SolidWorks *Top looks down -Y and *Front looks down -Z, so mapping
    up -> +Y and front -> +Z is what makes the standard views correct.

USAGE
    python patch_gemini1_straighten_assembly.py "C:\\CMS_Local_Workspace\\gemini1.bas"
    --dry-run    report only, write nothing

Idempotent per edit. Timestamped .bak before writing.
"""

import shutil
import sys
import time
from pathlib import Path

ENCODING = "latin-1"
MARKER = "StraightenWholeAssemblyToTrueAxes"

EDIT_1_ANCHOR = (
    "Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean\n"
)

NEW_FUNCTION = '''\
' ============================================================================
' STRAIGHTEN THE WHOLE ASSEMBLY TO ITS TRUE AXES
'
' Same trick as the pullcore DXF straighten (PULLCORE_STRAIGHTEN_SIGN *
' detectedAngle), but in 3D and applied to the geometry instead of the output:
' measure the mold's real axes, then rigidly rotate every component so those
' axes land on the model's XYZ.
'
' WHY THIS AND NOT A VIEW FIX
'   The imported X_T is not axis-aligned, so ALL SIX standard views are crooked
'   -- they are defined by the model planes and the part is not square to them.
'   Rotating the camera (Orientation3) fixes one picture and leaves the geometry
'   tilted, so the DXF, the ISO JPGs and the STL each need their own correction.
'   Rotating the GEOMETRY once fixes every output at the same time, and makes
'   CaptureFinalStandardViewsForStlCoordinateSystem a no-op instead of a patch.
'
' AFTER THIS RUNS
'   up    (TCP->BCP)        lands on model +Y
'   front (holders->pots)   lands on model +Z
'   right                   lands on model +X
'   so *Top is a true top view and *Front is a true front view.
'
' Returns True if the assembly was rotated. Returns False having changed
' nothing if the axes cannot be measured, so callers can fall back.
Private Function StraightenWholeAssemblyToTrueAxes(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    StraightenWholeAssemblyToTrueAxes = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    ' ---- 1. UP, from the two outer clamp plates ------------------------
    Dim tcpComp As Object, bcpComp As Object
    Set tcpComp = FindComponentByKeys(model, TCP_TOP_ORIENTATION_KEYS)
    Set bcpComp = FindComponentByKeys(model, BCP_BOTTOM_ORIENTATION_KEYS)

    If tcpComp Is Nothing Or bcpComp Is Nothing Then
        LogLine "STRAIGHTEN skipped: TCP or BCP component not found by name."
        Exit Function
    End If

    Dim tX As Double, tY As Double, tZ As Double
    Dim bX As Double, bY As Double, bZ As Double
    If TryGetComponentCenterInches(tcpComp, tX, tY, tZ) = False Then
        LogLine "STRAIGHTEN skipped: could not read TCP center."
        Exit Function
    End If
    If TryGetComponentCenterInches(bcpComp, bX, bY, bZ) = False Then
        LogLine "STRAIGHTEN skipped: could not read BCP center."
        Exit Function
    End If

    Dim ux As Double, uy As Double, uz As Double, uLen As Double
    ux = tX - bX: uy = tY - bY: uz = tZ - bZ
    uLen = Sqr(ux * ux + uy * uy + uz * uz)
    If uLen < 0.001 Then
        LogLine "STRAIGHTEN skipped: TCP and BCP centers coincide."
        Exit Function
    End If
    ux = ux / uLen: uy = uy / uLen: uz = uz / uLen

    ' ---- 2. FRONT, from holders -> pots, perpendicularised to UP -------
    ' Pots sit in front of the holders, so that vector is the front direction.
    ' If the pots cannot be found we fall back to any perpendicular; the
    ' existing front pass then rotates about the (now correct) axis.
    Dim fx As Double, fy As Double, fz As Double
    Dim haveFront As Boolean
    haveFront = False

    Dim idhComp As Object, odhComp As Object
    Set idhComp = FindComponentByKeys(model, ID_HOLDER_KEYS)
    Set odhComp = FindComponentByKeys(model, OD_HOLDER_KEYS)

    Dim potX As Double, potY As Double, potZ As Double
    Dim hldX As Double, hldY As Double, hldZ As Double

    If TryGetTrueAxisPotAndHolderCenters(model, potX, potY, potZ, hldX, hldY, hldZ) Then
        fx = potX - hldX: fy = potY - hldY: fz = potZ - hldZ
        haveFront = True
    End If

    If haveFront Then
        ' Gram-Schmidt against up.
        Dim dpf As Double
        dpf = fx * ux + fy * uy + fz * uz
        fx = fx - dpf * ux: fy = fy - dpf * uy: fz = fz - dpf * uz
        Dim fLen As Double
        fLen = Sqr(fx * fx + fy * fy + fz * fz)
        If fLen < 0.001 Then
            haveFront = False
        Else
            fx = fx / fLen: fy = fy / fLen: fz = fz / fLen
        End If
    End If

    If haveFront = False Then
        ' Any stable perpendicular: start from the world axis least parallel to up.
        Dim ax As Double, ay As Double, az As Double
        ax = Abs(ux): ay = Abs(uy): az = Abs(uz)
        If ax <= ay And ax <= az Then
            fx = 1#: fy = 0#: fz = 0#
        ElseIf ay <= ax And ay <= az Then
            fx = 0#: fy = 1#: fz = 0#
        Else
            fx = 0#: fy = 0#: fz = 1#
        End If
        Dim dpf2 As Double
        dpf2 = fx * ux + fy * uy + fz * uz
        fx = fx - dpf2 * ux: fy = fy - dpf2 * uy: fz = fz - dpf2 * uz
        Dim fLen2 As Double
        fLen2 = Sqr(fx * fx + fy * fy + fz * fz)
        If fLen2 < 0.000001 Then
            LogLine "STRAIGHTEN skipped: could not build a perpendicular to up."
            Exit Function
        End If
        fx = fx / fLen2: fy = fy / fLen2: fz = fz / fLen2
        LogLine "STRAIGHTEN: pot/holder front unavailable; using a synthetic " & _
                "perpendicular. The existing front pass will still rotate " & _
                "pots ahead of holders."
    End If

    ' ---- 3. RIGHT = up x front (right-handed: right x up = front) ------
    Dim rx As Double, ry As Double, rz As Double
    rx = uy * fz - uz * fy
    ry = uz * fx - ux * fz
    rz = ux * fy - uy * fx

    ' ---- 4. R rows = right / up / front --------------------------------
    Dim R(0 To 2, 0 To 2) As Double
    R(0, 0) = rx: R(0, 1) = ry: R(0, 2) = rz
    R(1, 0) = ux: R(1, 1) = uy: R(1, 2) = uz
    R(2, 0) = fx: R(2, 1) = fy: R(2, 2) = fz

    ' Already square? Then leave the model completely alone.
    Dim offAxis As Double
    offAxis = Abs(Abs(uy) - 1#) + Abs(Abs(fz) - 1#) + Abs(Abs(rx) - 1#)
    If offAxis < 0.0001 Then
        LogLine "STRAIGHTEN: assembly is already axis-aligned; nothing to do."
        StraightenWholeAssemblyToTrueAxes = True
        Exit Function
    End If

    LogLine "STRAIGHTEN: rotating whole assembly onto its true axes."
    LogLine "  up    (TCP-BCP)      = " & FormatNumberForCsv(ux) & ", " & _
            FormatNumberForCsv(uy) & ", " & FormatNumberForCsv(uz) & "   -> +Y"
    LogLine "  front (holders-pots) = " & FormatNumberForCsv(fx) & ", " & _
            FormatNumberForCsv(fy) & ", " & FormatNumberForCsv(fz) & "   -> +Z"
    LogLine "  right                = " & FormatNumberForCsv(rx) & ", " & _
            FormatNumberForCsv(ry) & ", " & FormatNumberForCsv(rz) & "   -> +X"

    ' ---- 5. apply R to every component --------------------------------
    Dim swMathUtil As Object
    Set swMathUtil = swApp.GetMathUtility
    If swMathUtil Is Nothing Then
        LogLine "STRAIGHTEN failed: GetMathUtility returned Nothing."
        Exit Function
    End If

    Dim vComps As Variant
    vComps = model.GetComponents(False)
    If IsEmpty(vComps) Then
        LogLine "STRAIGHTEN failed: GetComponents returned nothing."
        Exit Function
    End If

    Dim moved As Long, failed As Long
    moved = 0: failed = 0

    Dim ci As Long
    For ci = 0 To UBound(vComps)

        Dim c As Object
        Set c = vComps(ci)

        If Not c Is Nothing Then
            If StraightenOneComponentByMatrix(c, R, swMathUtil) Then
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
                "Components may be fixed -- see note in the patch header."
        Exit Function
    End If

    On Error Resume Next
    model.EditRebuild3
    model.GraphicsRedraw2
    On Error GoTo ErrHandler

    ' Re-scan: every cached center / bbox in parts() is now stale.
    LogLine "STRAIGHTEN: assembly geometry is now square to model XYZ. " & _
            "*Top and *Front are true. Re-scan CAD before using cached dims."

    StraightenWholeAssemblyToTrueAxes = True
    Exit Function

ErrHandler:
    LogLine "StraightenWholeAssemblyToTrueAxes error: " & Err.Description
    StraightenWholeAssemblyToTrueAxes = False
End Function


' Rotate one component by R, working directly on Transform2.ArrayData.
'
' Per the convention documented at TryGetComponentAssemblyPullAngleDeg:
'   v(0..2) = local X axis in assembly coords   (column 0)
'   v(3..5) = local Y axis                      (column 1)
'   v(6..8) = local Z axis                      (column 2)
'   v(9..11) = translation, v(12) = scale
' so new column = R * old column, and new translation = R * old translation.
'
' Doing the arithmetic here rather than via IMathTransform.Multiply avoids any
' multiply-order ambiguity.
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

    ' Three rotation columns, then the translation, all through R.
    Dim base As Long
    Dim col As Long
    For col = 0 To 3
        base = col * 3
        If col = 3 Then base = 9          ' translation lives at 9..11
        n(base + 0) = R(0, 0) * o(base + 0) + R(0, 1) * o(base + 1) + R(0, 2) * o(base + 2)
        n(base + 1) = R(1, 0) * o(base + 0) + R(1, 1) * o(base + 1) + R(1, 2) * o(base + 2)
        n(base + 2) = R(2, 0) * o(base + 0) + R(2, 1) * o(base + 1) + R(2, 2) * o(base + 2)
    Next col

    n(12) = o(12)                          ' preserve scale
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


' Average pot-block centre and average holder centre, for the front vector.
' Falls back to False if either side cannot be identified by name.
Private Function TryGetTrueAxisPotAndHolderCenters(ByVal model As Object, _
                                                   ByRef potX As Double, ByRef potY As Double, ByRef potZ As Double, _
                                                   ByRef hldX As Double, ByRef hldY As Double, ByRef hldZ As Double) As Boolean
On Error GoTo eh

    TryGetTrueAxisPotAndHolderCenters = False

    Dim idh As Object, odh As Object
    Set idh = FindComponentByKeys(model, ID_HOLDER_KEYS)
    Set odh = FindComponentByKeys(model, OD_HOLDER_KEYS)
    If idh Is Nothing Or odh Is Nothing Then Exit Function

    Dim ax As Double, ay As Double, az As Double
    Dim bx As Double, by As Double, bz As Double
    If TryGetComponentCenterInches(idh, ax, ay, az) = False Then Exit Function
    If TryGetComponentCenterInches(odh, bx, by, bz) = False Then Exit Function

    hldX = (ax + bx) / 2#
    hldY = (ay + by) / 2#
    hldZ = (az + bz) / 2#

    ' Pot blocks: reuse the export/BOM-matched indexes the front pass already
    ' relies on, via the same quote names StandardPlateName produces.
    Dim iIdPot As Long, iOdPot As Long
    iIdPot = FindCadPartIndexByQuoteOrKeys("ID POT BLOCK", "")
    iOdPot = FindCadPartIndexByQuoteOrKeys("OD POT BLOCK", "")

    If iIdPot < 1 Or iOdPot < 1 Then Exit Function
    If iIdPot > PartCount Or iOdPot > PartCount Then Exit Function
    If parts(iIdPot).hasAsmCenter = False Then Exit Function
    If parts(iOdPot).hasAsmCenter = False Then Exit Function

    potX = (parts(iIdPot).AsmCenterX + parts(iOdPot).AsmCenterX) / 2#
    potY = (parts(iIdPot).AsmCenterY + parts(iOdPot).AsmCenterY) / 2#
    potZ = (parts(iIdPot).AsmCenterZ + parts(iOdPot).AsmCenterZ) / 2#

    TryGetTrueAxisPotAndHolderCenters = True
    Exit Function

eh:
    TryGetTrueAxisPotAndHolderCenters = False
End Function

'''

EDIT_2_OLD = '''    If TryOrientTcpUpByViewProjection(model) Then
        TryShowTcpTopViewFromComponentCenters = True
        Exit Function
    End If
'''

EDIT_2_NEW = '''    ' Straighten the GEOMETRY first. Once the assembly is square to model XYZ
    ' the plain standard views are correct, so just show *Top and stop -- no
    ' camera trickery, and the DXF / JPG / STL all inherit the fix.
    If StraightenWholeAssemblyToTrueAxes(model) Then
        model.ShowNamedView2 "*Top", 5
        StabilizeActiveView model, 100
        LogLine "TCP-top orientation: assembly straightened, using true *Top."
        TryShowTcpTopViewFromComponentCenters = True
        Exit Function
    End If

    If TryOrientTcpUpByViewProjection(model) Then
        TryShowTcpTopViewFromComponentCenters = True
        Exit Function
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

    if MARKER + "(ByVal model As Object) As Boolean" in text:
        print("  skip  1/2 functions already present")
        skipped += 1
    else:
        n = text.count(EDIT_1_ANCHOR)
        if n != 1:
            raise SystemExit(f"ABORT: edit-1 anchor matched {n} times, expected 1. Nothing written.")
        text = text.replace(EDIT_1_ANCHOR, NEW_FUNCTION + EDIT_1_ANCHOR, 1)
        print("  ok    1/2 inserted StraightenWholeAssemblyToTrueAxes (+2 helpers)")
        changed += 1

    if "If " + MARKER + "(model) Then" in text:
        print("  skip  2/2 caller already rewired")
        skipped += 1
    else:
        n = text.count(EDIT_2_OLD)
        if n != 1:
            raise SystemExit(f"ABORT: edit-2 anchor matched {n} times, expected 1. Nothing written.")
        text = text.replace(EDIT_2_OLD, EDIT_2_NEW, 1)
        print("  ok    2/2 rewired TryShowTcpTopViewFromComponentCenters")
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
    print("Log lines to look for:")
    print("  STRAIGHTEN: rotating whole assembly onto its true axes.")
    print("    up    (TCP-BCP)      = 0.0000, 0.9986, -0.0523   -> +Y")
    print("  STRAIGHTEN: rotated 27 component(s), 0 failed.")
    print()
    print("TWO THINGS THAT CAN GO WRONG, both logged explicitly:")
    print()
    print("1. 'STRAIGHTEN failed: no component transforms could be set.'")
    print("   Fixed components can refuse Transform2. Fix: float them first with")
    print("   swAssy.UnfixComponent after selecting all, or set the top-level")
    print("   components floating before this runs. Tell me and I'll add it.")
    print()
    print("2. 'STRAIGHTEN skipped: TCP or BCP component not found by name.'")
    print("   J8494's parts are named '...sldasm-Part-24-1' with no TCP/SMED")
    print("   token, so FindComponentByKeys may come back empty. Then the axis")
    print("   has to come from the BOM-matched indices instead of names.")
    print()
    print("IMPORTANT: this rotates geometry, so every cached centre and bbox in")
    print("parts() is stale afterwards. It runs before the export passes, but if")
    print("you see dims that disagree with the drawing, re-scan after straightening.")


if __name__ == "__main__":
    main()
