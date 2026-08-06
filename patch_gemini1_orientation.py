"""
Apply the TRUE-AXIS ORIENTATION fix to gemini1.bas.

WHY
    The imported X_T is not axis-aligned, so all six standard views (*Top,
    *Bottom, *Front, *Right...) are crooked -- they are defined by the model's
    XYZ planes and the part is not square to them. OrientTcpTopFromCenters makes
    it worse by snapping to whichever standard view the TCP->BCP vector is MOST
    aligned with, which rounds a 30-degree error straight into the drawing.

    This patch builds the rotation matrix from the geometry and assigns it to
    IModelView.Orientation3 directly. Exact at any import angle.

USAGE
    python patch_gemini1_orientation.py "C:\\CMS_Local_Workspace\\gemini1.bas"

    --dry-run    report what would change, write nothing

IDEMPOTENT
    Each edit is checked independently. Already-applied edits are skipped, not
    treated as an error -- that is what tripped the earlier patch script when it
    refused on a whole-file "already contains" test.

Writes a timestamped .bak before touching anything.
"""

import shutil
import sys
import time
from pathlib import Path

# latin-1 round-trips every byte 0..255, so nothing outside the patched regions
# can change in a cp1252 VBA file.
ENCODING = "latin-1"

MARKER = "TryOrientTcpUpByTrueGeometricAxis"


# ---------------------------------------------------------------------------
# EDIT 1 -- insert the new function immediately before
#           TryOrientTcpUpByViewProjection.
# ---------------------------------------------------------------------------
EDIT_1_ANCHOR = (
    "Private Function TryOrientTcpUpByViewProjection(ByVal model As Object) As Boolean\n"
)

NEW_FUNCTION = '''\
' Orient the view so the mold's real TCP->BCP axis points straight at the
' camera, regardless of how the imported geometry is rotated in model space.
'
' THE PROBLEM THIS SOLVES
'   The imported X_T is not axis-aligned. The mold sits at a compound angle to
'   the model's XYZ planes, so *Top, *Front, *Right -- ALL SIX standard views --
'   are crooked. Switching *Bottom -> *Top or nudging CMS_TOP_ROTATE_Z_STEPS
'   cannot help, because every one of those views is defined by the model axes
'   and the part is not square to them.
'
'   OrientTcpTopFromCenters compounds it: it measures the TCP->BCP vector, asks
'   which axis it is MOST aligned with, and snaps to that standard view --
'       "Matched TOP/BOTTOM HOLDER selected *Bottom. Stack axis=Y"
'   On a skewed import "most aligned" can be 30 degrees off. That snap IS the
'   crookedness.
'
' THE FIX
'   Build the matrix from geometry instead of picking a view.
'   IModelView.Orientation3 is a settable 9-element rotation matrix (this macro
'   already READS it in five places for the STL post-rotation). Rows are the
'   view axes in model space: row0 = screen right, row1 = screen up,
'   row2 = out of screen toward the viewer. So a true top view is just:
'       up = normalize(TCP center - BCP center), row2 = up, row1 = any perp,
'       row0 = row1 x row2.
'
' WHY ONLY THE UP AXIS MATTERS
'   The front pass already works -- the J8494 log says "Final *Front
'   verification OK: pot blocks are closer to front than holders."
'   DefineStandardFrontFromHolderAndPotCom and friends operate in ACTIVE VIEW
'   coordinates and spin about the view axis, so given a correct axis they put
'   the pots in front on their own. This function deliberately leaves them alone.
'
' Returns True if the view was set. Returns False having changed nothing if the
' TCP/BCP components or centers cannot be read, so the caller falls back.
Private Function TryOrientTcpUpByTrueGeometricAxis(ByVal model As Object) As Boolean
On Error GoTo ErrHandler

    TryOrientTcpUpByTrueGeometricAxis = False

    If model Is Nothing Then Exit Function
    If model.GetType <> swDocASSEMBLY Then Exit Function

    ' --- 1. the two outer plates give us the true stack axis --------------
    Dim tcpComp As Object
    Dim bcpComp As Object

    Set tcpComp = FindComponentByKeys(model, TCP_TOP_ORIENTATION_KEYS)
    Set bcpComp = FindComponentByKeys(model, BCP_BOTTOM_ORIENTATION_KEYS)

    If tcpComp Is Nothing Or bcpComp Is Nothing Then
        LogLine "True-axis orientation skipped: TCP or BCP component not found."
        Exit Function
    End If

    Dim tcpX As Double, tcpY As Double, tcpZ As Double
    Dim bcpX As Double, bcpY As Double, bcpZ As Double

    If TryGetComponentCenterInches(tcpComp, tcpX, tcpY, tcpZ) = False Then
        LogLine "True-axis orientation skipped: could not read TCP center."
        Exit Function
    End If

    If TryGetComponentCenterInches(bcpComp, bcpX, bcpY, bcpZ) = False Then
        LogLine "True-axis orientation skipped: could not read BCP center."
        Exit Function
    End If

    ' up points FROM the bottom clamp TOWARD the top clamp, so looking down it
    ' puts the TCP nearest the camera.
    Dim ux As Double, uy As Double, uz As Double
    ux = tcpX - bcpX
    uy = tcpY - bcpY
    uz = tcpZ - bcpZ

    Dim uLen As Double
    uLen = Sqr(ux * ux + uy * uy + uz * uz)

    If uLen < 0.001 Then
        LogLine "True-axis orientation skipped: TCP and BCP centers coincide."
        Exit Function
    End If

    ux = ux / uLen
    uy = uy / uLen
    uz = uz / uLen

    ' --- 2. any perpendicular will do for screen-up -----------------------
    ' The front pass rotates about the view axis afterwards, so this only has to
    ' be perpendicular and deterministic. Start from whichever world axis is
    ' LEAST parallel to up, for the most stable cross product.
    Dim sx As Double, sy As Double, sz As Double
    Dim ax As Double, ay As Double, az As Double

    ax = Abs(ux): ay = Abs(uy): az = Abs(uz)

    If ax <= ay And ax <= az Then
        sx = 1#: sy = 0#: sz = 0#
    ElseIf ay <= ax And ay <= az Then
        sx = 0#: sy = 1#: sz = 0#
    Else
        sx = 0#: sy = 0#: sz = 1#
    End If

    ' Gram-Schmidt: strip the component along up, leaving a true perpendicular.
    Dim dp As Double
    dp = sx * ux + sy * uy + sz * uz

    sx = sx - dp * ux
    sy = sy - dp * uy
    sz = sz - dp * uz

    Dim sLen As Double
    sLen = Sqr(sx * sx + sy * sy + sz * sz)

    If sLen < 0.000001 Then
        LogLine "True-axis orientation skipped: could not build a perpendicular."
        Exit Function
    End If

    sx = sx / sLen
    sy = sy / sLen
    sz = sz / sLen

    ' --- 3. right = screenUp x outOfScreen, completing a right-handed basis
    Dim rx As Double, ry As Double, rz As Double
    rx = sy * uz - sz * uy
    ry = sz * ux - sx * uz
    rz = sx * uy - sy * ux

    ' --- 4. hand the matrix straight to the view --------------------------
    Dim swViewTA As Object
    Set swViewTA = model.ActiveView

    If swViewTA Is Nothing Then
        LogLine "True-axis orientation failed: ActiveView is Nothing."
        Exit Function
    End If

    ' The app runs invisible, so graphics updates must be forced on or the view
    ' assignment silently no-ops -- same reason
    ' CaptureFinalStandardViewsForStlCoordinateSystem does this.
    On Error Resume Next
    swViewTA.EnableGraphicsUpdate = True
    On Error GoTo ErrHandler

    Dim mTA(0 To 8) As Double
    mTA(0) = rx: mTA(1) = ry: mTA(2) = rz     ' row 0 - screen right
    mTA(3) = sx: mTA(4) = sy: mTA(5) = sz     ' row 1 - screen up
    mTA(6) = ux: mTA(7) = uy: mTA(8) = uz     ' row 2 - out of screen (stack axis)

    swViewTA.Orientation3 = mTA

    On Error Resume Next
    model.GraphicsRedraw2
    On Error GoTo ErrHandler

    StabilizeActiveView model, 100

    LogLine "TRUE-AXIS orientation applied. Imported geometry was not " & _
            "axis-aligned, so no standard view could be square to it."
    LogLine "  stack axis (TCP-BCP) = " & _
            FormatNumberForCsv(ux) & ", " & _
            FormatNumberForCsv(uy) & ", " & _
            FormatNumberForCsv(uz)
    LogLine "  screen up            = " & _
            FormatNumberForCsv(sx) & ", " & _
            FormatNumberForCsv(sy) & ", " & _
            FormatNumberForCsv(sz)
    LogLine "  screen right         = " & _
            FormatNumberForCsv(rx) & ", " & _
            FormatNumberForCsv(ry) & ", " & _
            FormatNumberForCsv(rz)
    LogLine "  Front pass will now rotate about this axis to put pots ahead " & _
            "of holders."

    TryOrientTcpUpByTrueGeometricAxis = True
    Exit Function

ErrHandler:
    LogLine "TryOrientTcpUpByTrueGeometricAxis error: " & Err.Description
    TryOrientTcpUpByTrueGeometricAxis = False
End Function

'''


# ---------------------------------------------------------------------------
# EDIT 2 -- make it the FIRST attempt inside
#           TryShowTcpTopViewFromComponentCenters, demoting the old
#           snap-to-nearest-standard-view method to fallback.
#
# The `TryShowTcpTopViewFromComponentCenters = True` line inside the block makes
# this anchor unique to that one function.
# ---------------------------------------------------------------------------
EDIT_2_OLD = '''    If TryOrientTcpUpByViewProjection(model) Then
        TryShowTcpTopViewFromComponentCenters = True
        Exit Function
    End If
'''

EDIT_2_NEW = '''    ' Exact geometric axis FIRST. TryOrientTcpUpByViewProjection below snaps to
    ' whichever standard view is closest, which cannot be square to a skewed
    ' import -- it is now only the fallback.
    If TryOrientTcpUpByTrueGeometricAxis(model) Then
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
    print(f"     {len(original):,} chars, {original.count(chr(10)):,} lines")
    print()

    changed = 0
    skipped = 0

    # ---- edit 1: insert the function --------------------------------------
    if MARKER + "(ByVal model As Object) As Boolean" in text:
        print("  skip  1/2 function already present")
        skipped += 1
    else:
        n = text.count(EDIT_1_ANCHOR)
        if n != 1:
            raise SystemExit(
                f"ABORT: anchor for edit 1 matched {n} times, expected 1.\n"
                f"  anchor: {EDIT_1_ANCHOR.strip()}\n"
                f"Nothing was written."
            )
        text = text.replace(EDIT_1_ANCHOR, NEW_FUNCTION + EDIT_1_ANCHOR, 1)
        print("  ok    1/2 inserted TryOrientTcpUpByTrueGeometricAxis")
        changed += 1

    # ---- edit 2: rewire the caller ----------------------------------------
    if "If " + MARKER + "(model) Then" in text:
        print("  skip  2/2 caller already rewired")
        skipped += 1
    else:
        n = text.count(EDIT_2_OLD)
        if n != 1:
            raise SystemExit(
                f"ABORT: anchor for edit 2 matched {n} times, expected 1.\n"
                f"Nothing was written."
            )
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
    print(f"Patched: {path}")
    print(f"         {len(text):,} chars, {text.count(chr(10)):,} lines "
          f"({len(text) - len(original):+,} chars)")
    print()
    print("Paste the file into the SolidWorks VBA editor and re-run J8494.")
    print("Look for this in the log:")
    print()
    print("  TRUE-AXIS orientation applied. Imported geometry was not axis-aligned...")
    print("    stack axis (TCP-BCP) = 0.0000, 0.9986, -0.0523")
    print()
    print("A stack axis like that is the smoking gun -- 3 degrees off Y, which is")
    print("exactly the crookedness no standard view could correct.")
    print()
    print("If instead you see:")
    print("  True-axis orientation skipped: TCP or BCP component not found")
    print("then this job's parts carry no TCP/SMED token (J8494's are named")
    print("'...sldasm-Part-24-1'). Add the customer's spelling to")
    print("TCP_TOP_ORIENTATION_KEYS / BCP_BOTTOM_ORIENTATION_KEYS, or tell me and")
    print("I will wire the BOM-matched-index version instead of name lookup.")


if __name__ == "__main__":
    main()
