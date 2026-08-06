"""
gemini1.bas -- ORIENT AFTER STRAIGHTENING (pass 2).

Pass 1 (already working) squares the geometry to model XYZ using a plate's
placement transform. This is pass 2: put the mold the RIGHT WAY UP now that it
is square.

WHY THIS IS EASY NOW
    Before pass 1 the mold sat 54 degrees off axis, so no exact fix existed.
    After pass 1 the mold's own axes ARE model axes -- just possibly the wrong
    ones. So the remaining correction is a SIGNED AXIS PERMUTATION: one of 24
    exact 90-degree orientations. No fitted anything, no tolerance fudging.

WHAT IT ENFORCES
    up    = TCP  -> BCP direction   lands on model +Y
    front = holders -> pots         lands on model +Z
    right = up x front              lands on model +X

    SolidWorks *Top looks down -Y and *Front looks down -Z, so that mapping
    makes *Top a true top-of-TCP view and *Front look at the pot side.

WHY IT RUNS HERE
    Line 1229, immediately before EnsureCmsTopOrientationFromMatchedTcpBcp.
    That is after BOM matching, so TCP / holders / pots are known -- the
    chicken-and-egg that blocked doing this during pass 1. Pass 1 must already
    have run, or the vectors are not near-axis and the snap is meaningless.

    Once this succeeds the downstream view code has nothing left to correct:
    it should log "Front definition: pots are already closer to front" instead
    of flipping twice and warning.

USAGE
    python patch_gemini1_orient_after_straighten.py "C:\\CMS_Local_Workspace\\gemini1.bas"
    --dry-run

Requires patch_gemini1_straighten_first.py to have been applied first (it
reuses StraightenOneComponentByMatrix). Idempotent. Timestamped .bak.
"""

import shutil
import sys
import time
from pathlib import Path

ENCODING = "latin-1"
MARKER = "OrientMoldToTcpTopAndPotFront"
HELPER = "StraightenOneComponentByMatrix"

EDIT_1_ANCHOR = "Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean\n"

NEW_CODE = '''\
' ============================================================================
' PASS 2 -- PUT THE STRAIGHTENED MOLD THE RIGHT WAY UP
'
' Pass 1 (StraightenAssemblyFromPlateTransform) made the geometry square to
' model XYZ. It did not decide WHICH axis is up. This does:
'
'     up    = TCP -> BCP           ->  model +Y
'     front = holders -> pots      ->  model +Z
'     right = up x front           ->  model +X
'
' Because the geometry is already square, both measured vectors land within a
' fraction of a degree of a model axis, so this is a SIGNED PERMUTATION -- one
' of 24 exact 90-degree orientations. Snapping to the dominant axis keeps it
' exact and stops any residual skew accumulating.
'
' SolidWorks *Top looks down -Y and *Front down -Z, so this mapping makes *Top
' the true top of the TCP and *Front the pot side.
Private Function OrientMoldToTcpTopAndPotFront(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    OrientMoldToTcpTopAndPotFront = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function
    If PartCount < 1 Then Exit Function

    ' ---- 1. UP: from BCP toward TCP -------------------------------------
    Dim iTcp As Long, iBcp As Long
    iTcp = FindCadPartIndexByQuoteOrKeys("TCP", TCP_TOP_ORIENTATION_KEYS)
    iBcp = FindCadPartIndexByQuoteOrKeys("BCP", BCP_BOTTOM_ORIENTATION_KEYS)

    If iTcp < 1 Or iBcp < 1 Or iTcp > PartCount Or iBcp > PartCount Then
        LogLine "ORIENT pass 2 skipped: TCP or BCP index not resolved " & _
                "(TCP=" & CStr(iTcp) & " BCP=" & CStr(iBcp) & ")."
        Exit Function
    End If

    Dim ux As Double, uy As Double, uz As Double
    ux = parts(iTcp).AsmCenterX - parts(iBcp).AsmCenterX
    uy = parts(iTcp).AsmCenterY - parts(iBcp).AsmCenterY
    uz = parts(iTcp).AsmCenterZ - parts(iBcp).AsmCenterZ

    If SnapVectorToSignedAxis(ux, uy, uz) = False Then
        LogLine "ORIENT pass 2 skipped: TCP-BCP vector is degenerate."
        Exit Function
    End If

    ' ---- 2. FRONT: from holders toward pots -----------------------------
    Dim iIdh As Long, iOdh As Long, iIdp As Long, iOdp As Long
    iIdh = FindCadPartIndexByQuoteOrKeys("ID HOLDER", ID_HOLDER_KEYS)
    iOdh = FindCadPartIndexByQuoteOrKeys("OD HOLDER", OD_HOLDER_KEYS)
    iIdp = FindCadPartIndexByQuoteOrKeys("ID POT BLOCK", "")
    iOdp = FindCadPartIndexByQuoteOrKeys("OD POT BLOCK", "")

    Dim haveFront As Boolean
    Dim fx As Double, fy As Double, fz As Double
    haveFront = False

    If iIdh >= 1 And iOdh >= 1 And iIdp >= 1 And iOdp >= 1 Then
        If iIdh <= PartCount And iOdh <= PartCount And _
           iIdp <= PartCount And iOdp <= PartCount Then

            Dim hX As Double, hY As Double, hZ As Double
            Dim pX As Double, pY As Double, pZ As Double

            hX = (parts(iIdh).AsmCenterX + parts(iOdh).AsmCenterX) / 2#
            hY = (parts(iIdh).AsmCenterY + parts(iOdh).AsmCenterY) / 2#
            hZ = (parts(iIdh).AsmCenterZ + parts(iOdh).AsmCenterZ) / 2#

            pX = (parts(iIdp).AsmCenterX + parts(iOdp).AsmCenterX) / 2#
            pY = (parts(iIdp).AsmCenterY + parts(iOdp).AsmCenterY) / 2#
            pZ = (parts(iIdp).AsmCenterZ + parts(iOdp).AsmCenterZ) / 2#

            fx = pX - hX: fy = pY - hY: fz = pZ - hZ

            ' Strip anything along up, so front is a clean perpendicular, then
            ' snap. The pot/holder offset along the stack axis is irrelevant.
            Dim dp As Double
            dp = fx * ux + fy * uy + fz * uz
            fx = fx - dp * ux: fy = fy - dp * uy: fz = fz - dp * uz

            If SnapVectorToSignedAxis(fx, fy, fz) Then haveFront = True
        End If
    End If

    If haveFront = False Then
        LogLine "ORIENT pass 2: pot/holder front unavailable; " & _
                "using a synthetic perpendicular. TCP-top is still enforced."
        If Abs(ux) < 0.5 Then
            fx = 1#: fy = 0#: fz = 0#
        ElseIf Abs(uy) < 0.5 Then
            fx = 0#: fy = 1#: fz = 0#
        Else
            fx = 0#: fy = 0#: fz = 1#
        End If
        Dim dp2 As Double
        dp2 = fx * ux + fy * uy + fz * uz
        fx = fx - dp2 * ux: fy = fy - dp2 * uy: fz = fz - dp2 * uz
        If SnapVectorToSignedAxis(fx, fy, fz) = False Then Exit Function
    End If

    ' ---- 3. right = up x front (right-handed: right x up = front) -------
    Dim rx As Double, ry As Double, rz As Double
    rx = uy * fz - uz * fy
    ry = uz * fx - ux * fz
    rz = ux * fy - uy * fx

    ' ---- 4. R rows = right / up / front --------------------------------
    Dim R(0 To 2, 0 To 2) As Double
    R(0, 0) = rx: R(0, 1) = ry: R(0, 2) = rz
    R(1, 0) = ux: R(1, 1) = uy: R(1, 2) = uz
    R(2, 0) = fx: R(2, 1) = fy: R(2, 2) = fz

    LogLine "ORIENT pass 2 (TCP top, pots front):"
    LogLine "  up    TCP<-BCP  = " & FormatNumberForCsv(ux) & ", " & _
            FormatNumberForCsv(uy) & ", " & FormatNumberForCsv(uz) & "   -> +Y"
    LogLine "  front holders->pots = " & FormatNumberForCsv(fx) & ", " & _
            FormatNumberForCsv(fy) & ", " & FormatNumberForCsv(fz) & "   -> +Z"
    LogLine "  right           = " & FormatNumberForCsv(rx) & ", " & _
            FormatNumberForCsv(ry) & ", " & FormatNumberForCsv(rz) & "   -> +X"

    If uy > 0.999 And fz > 0.999 And Abs(rx) > 0.999 Then
        LogLine "ORIENT pass 2: already TCP-up and pots-front; nothing to do."
        OrientMoldToTcpTopAndPotFront = True
        Exit Function
    End If

    ' ---- 5. apply to every component -----------------------------------
    Dim swMathUtil As Object
    Set swMathUtil = swApp.GetMathUtility
    If swMathUtil Is Nothing Then
        LogLine "ORIENT pass 2 failed: GetMathUtility returned Nothing."
        Exit Function
    End If

    Dim vComps As Variant
    vComps = model.GetComponents(False)
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

    LogLine "ORIENT pass 2: rotated " & CStr(moved) & " component(s), " & _
            CStr(failed) & " failed."

    If moved = 0 Then
        LogLine "ORIENT pass 2 failed: no component transforms could be set."
        Exit Function
    End If

    On Error Resume Next
    model.EditRebuild3
    model.GraphicsRedraw2
    On Error GoTo ErrHandler

    LogLine "ORIENT pass 2 done. TCP is +Y (top), pots are +Z (front). " & _
            "*Top and *Front are now correct with no flipping needed."

    OrientMoldToTcpTopAndPotFront = True
    Exit Function

ErrHandler:
    LogLine "OrientMoldToTcpTopAndPotFront error: " & Err.Description
    OrientMoldToTcpTopAndPotFront = False
End Function


' Normalise a vector, then snap it to the nearest signed model axis.
'
' Only valid AFTER pass 1 has squared the geometry -- at that point the real
' vector is within a fraction of a degree of an axis, so snapping keeps the
' result an exact signed permutation instead of letting residual skew through.
' Returns False if the vector is too short to have a direction.
Private Function SnapVectorToSignedAxis(ByRef x As Double, _
                                        ByRef y As Double, _
                                        ByRef z As Double) As Boolean
On Error GoTo eh

    SnapVectorToSignedAxis = False

    Dim L As Double
    L = Sqr(x * x + y * y + z * z)
    If L < 0.000001 Then Exit Function

    x = x / L: y = y / L: z = z / L

    Dim ax As Double, ay As Double, az As Double
    ax = Abs(x): ay = Abs(y): az = Abs(z)

    If ax >= ay And ax >= az Then
        If x >= 0# Then
            x = 1#: y = 0#: z = 0#
        Else
            x = -1#: y = 0#: z = 0#
        End If
    ElseIf ay >= ax And ay >= az Then
        If y >= 0# Then
            x = 0#: y = 1#: z = 0#
        Else
            x = 0#: y = -1#: z = 0#
        End If
    Else
        If z >= 0# Then
            x = 0#: y = 0#: z = 1#
        Else
            x = 0#: y = 0#: z = -1#
        End If
    End If

    SnapVectorToSignedAxis = True
    Exit Function

eh:
    SnapVectorToSignedAxis = False
End Function

'''

# ---------------------------------------------------------------------------
# EDIT 2 -- run pass 2 immediately before the existing TCP-top view step.
# ---------------------------------------------------------------------------
EDIT_2_OLD = '''    LogStart "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
'''

EDIT_2_NEW = '''    ' ------------------------------------------------------------------
    ' PASS 2: now that pass 1 has squared the geometry, put it the right way
    ' up -- TCP to +Y, pots to +Z -- as an exact signed axis permutation.
    ' BOM matching has run by this point, so TCP / holders / pots are known.
    '
    ' Geometry moves, so re-scan before the view code reads any centre.
    ' ------------------------------------------------------------------
    If OrientMoldToTcpTopAndPotFront(swModel) Then
        LogStart "Re-scan CAD after TCP-top / pot-front orientation"
        ScanActiveSolidWorksDocument
        SortPartsByVolumeDescending
        LogLine "CAD PartCount=" & PartCount & " (after pass 2)"
        LogDone "Re-scan CAD after TCP-top / pot-front orientation"
    End If

    LogStart "Set TCP-top orientation from matched TCP/BCP, then save BASE"

    EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
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

    if HELPER + "(ByVal swComp As Object" not in text:
        raise SystemExit(
            f"ABORT: {HELPER} not found.\n"
            "Run patch_gemini1_straighten_first.py first -- pass 2 reuses its\n"
            "component-rotation helper. Nothing written."
        )

    changed = skipped = 0

    if MARKER + "(ByVal model As Object) As Boolean" in text:
        print("  skip  1/2 pass-2 functions already present")
        skipped += 1
    else:
        n = text.count(EDIT_1_ANCHOR)
        if n != 1:
            raise SystemExit(f"ABORT: edit-1 anchor matched {n} times, expected 1. Nothing written.")
        text = text.replace(EDIT_1_ANCHOR, NEW_CODE + EDIT_1_ANCHOR, 1)
        print(f"  ok    1/2 inserted {MARKER} + SnapVectorToSignedAxis")
        changed += 1

    if "If " + MARKER + "(swModel) Then" in text:
        print("  skip  2/2 already called before the TCP-top step")
        skipped += 1
    else:
        n = text.count(EDIT_2_OLD)
        if n != 1:
            raise SystemExit(f"ABORT: edit-2 anchor matched {n} times, expected 1. Nothing written.")
        text = text.replace(EDIT_2_OLD, EDIT_2_NEW, 1)
        print("  ok    2/2 hooked in before EnsureCmsTopOrientationFromMatchedTcpBcp")
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
    print("Re-run J8493. Expect, after the pass-1 straighten block:")
    print()
    print("  ORIENT pass 2 (TCP top, pots front):")
    print("    up    TCP<-BCP      = 0.0000, 1.0000, 0.0000   -> +Y")
    print("    front holders->pots = 0.0000, 0.0000, 1.0000   -> +Z")
    print("    right               = 1.0000, 0.0000, 0.0000   -> +X")
    print("  ORIENT pass 2: rotated 42 component(s), 0 failed.")
    print()
    print("Then the payoff, further down -- these should now PASS first time")
    print("instead of flipping twice and warning:")
    print("  Front definition: pots are already closer to front.")
    print("  Final *Front verification OK: pot blocks are closer to front than holders.")
    print()
    print("If you see 'ORIENT pass 2 skipped: TCP or BCP index not resolved',")
    print("FindCadPartIndexByQuoteOrKeys could not find them by quote name or")
    print("key. Send me that line and I'll pull the indices straight from")
    print("ExportRows instead, where the log already proves they are matched.")


if __name__ == "__main__":
    main()
