' ============================================================================
' gemini1.bas -- TRUE-AXIS ORIENTATION
'
' THE PROBLEM
'   The imported X_T is not axis-aligned. The mold sits at a compound angle to
'   the model's XYZ planes, so *Top, *Front, *Right -- ALL SIX standard views --
'   are crooked. Switching *Bottom -> *Top or nudging CMS_TOP_ROTATE_Z_STEPS
'   cannot fix that, because every one of those views is defined by the model
'   axes and the part simply is not square to them.
'
'   The existing code makes this worse by design. OrientTcpTopFromCenters
'   (line ~2200) measures the TCP->BCP vector, decides which axis it MOST
'   aligns with, and then snaps to that standard view:
'
'       "Matched TOP/BOTTOM HOLDER selected *Bottom. Stack axis=Y, TCP at high end=False"
'
'   On a skewed import "most aligned" can be 30 degrees off, and you get exactly
'   the crooked result in the screenshot.
'
' THE FIX
'   Stop snapping to a standard view. BUILD the rotation matrix from the
'   geometry and assign it to the view directly.
'
'   IModelView.Orientation3 is a settable 9-element rotation matrix (the macro
'   already READS it in five places for the STL post-rotation, e.g. line 5815).
'   Rows are the view axes expressed in model space:
'
'       row 0 = screen right
'       row 1 = screen up
'       row 2 = out of the screen, toward the viewer
'
'   So a true top view of the TCP is simply:
'
'       up   = normalize(TCP center - BCP center)     <- the real stack axis
'       row2 = up                                     <- look straight down it
'       row1 = any unit vector perpendicular to up
'       row0 = row1 x row2                            <- completes right-handed
'
'   That is exact at any import angle. No axis guessing, no snapping.
'
' WHY THIS IS ENOUGH
'   Only the UP axis has to be exact here. Once *Top is square to the mold, the
'   front rotation you already have takes over and it already works -- the J8494
'   log confirms it:
'
'       "Final *Front verification OK: pot blocks are closer to front than holders."
'
'   DefineStandardFrontFromHolderAndPotCom / EnsurePotBlocksCloserThanHolders
'   InActiveView / EnforcePotBlocksCloserAfterFrontPersist all operate in ACTIVE
'   VIEW coordinates and spin about the view axis. Feed them a correct axis and
'   they land the pots in front of the holders on their own. That is why this
'   patch deliberately does NOT touch the front logic.
'
' ============================================================================
' HOW TO APPLY -- 2 steps
' ============================================================================
'
' STEP 1  Paste the whole function below into gemini1.bas at module level.
'         Anywhere is fine; next to TryOrientTcpUpByViewProjection (~line 2147)
'         keeps it with its relatives.
'
' STEP 2  Wire it in as the FIRST attempt, so the old snap-to-standard-view
'         method becomes the fallback rather than the primary.
'
'         In TryShowTcpTopViewFromComponentCenters (~line 2126), FIND:
'
'             If TryOrientTcpUpByViewProjection(model) Then
'                 TryShowTcpTopViewFromComponentCenters = True
'                 Exit Function
'             End If
'
'         REPLACE WITH:
'
'             ' Exact geometric axis first. Only falls back to the old
'             ' snap-to-nearest-standard-view method if this cannot run.
'             If TryOrientTcpUpByTrueGeometricAxis(model) Then
'                 TryShowTcpTopViewFromComponentCenters = True
'                 Exit Function
'             End If
'
'             If TryOrientTcpUpByViewProjection(model) Then
'                 TryShowTcpTopViewFromComponentCenters = True
'                 Exit Function
'             End If
'
' Everything downstream is unchanged: PersistCurrentViewAsStandardTop captures
' whatever is on screen, so it will now capture the true top.
' ============================================================================


' Orient the view so the mold's real TCP->BCP axis points straight at the
' camera, regardless of how the imported geometry is rotated in model space.
'
' Returns True if the view was set. Returns False (having changed nothing) if
' the TCP/BCP components or their centers cannot be read, so the caller can
' fall back to the old method.
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
    ' The front pass rotates about the view axis afterwards, so this only has
    ' to be perpendicular and deterministic. Start from whichever world axis is
    ' LEAST parallel to up (most numerically stable cross product).
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

    ' --- 4. hand the matrix straight to the view -------------------------
    Dim swView As Object
    Set swView = model.ActiveView

    If swView Is Nothing Then
        LogLine "True-axis orientation failed: ActiveView is Nothing."
        Exit Function
    End If

    ' The app runs invisible, so graphics updates must be forced on or the view
    ' assignment silently no-ops -- same reason
    ' CaptureFinalStandardViewsForStlCoordinateSystem does this.
    On Error Resume Next
    swView.EnableGraphicsUpdate = True
    On Error GoTo ErrHandler

    Dim m(0 To 8) As Double
    m(0) = rx: m(1) = ry: m(2) = rz     ' row 0 - screen right
    m(3) = sx: m(4) = sy: m(5) = sz     ' row 1 - screen up
    m(6) = ux: m(7) = uy: m(8) = uz     ' row 2 - out of screen (the stack axis)

    swView.Orientation3 = m

    model.GraphicsRedraw2
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


' ============================================================================
' WHAT TO EXPECT IN THE LOG
'
'   TRUE-AXIS orientation applied. Imported geometry was not axis-aligned, ...
'     stack axis (TCP-BCP) = 0.0000, 0.9986, -0.0523
'     screen up            = 1.0000, 0.0000, 0.0000
'     screen right         = 0.0000, -0.0523, -0.9986
'   ...
'   Standard views REDEFINED: current TCP/top-side orientation assigned to *Top
'   ...
'   Final *Front verification OK: pot blocks are closer to front than holders.
'
' A stack axis like "0, 0.9986, -0.0523" is the smoking gun -- that is 3 degrees
' off the Y axis, which is precisely the crookedness no standard view could
' correct and the old snap method rounded away.
'
' If you see "True-axis orientation skipped: TCP or BCP component not found",
' the component names on this job do not match TCP_TOP_ORIENTATION_KEYS (line
' 101) or BCP_BOTTOM_ORIENTATION_KEYS. Add the customer's spelling to those
' pipe-delimited lists -- J8494's parts are named
' "...sldasm-Part-24-1" style, with no TCP/SMED token at all, in which case the
' function correctly declines and you need the BOM-matched index route instead.
'
' NOTE ON CMS_TOP_ROTATE_Z_STEPS (line 115): leave it at 0. It exists to nudge a
' snapped standard view, and with a true axis there is nothing to nudge. If the
' drawing comes out square but 90 degrees round, that is the front pass, not this.
' ============================================================================
