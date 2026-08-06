' ============================================================
' CMS ORIENTATION PATCH - NEW FUNCTIONS  (rev 2)
' ------------------------------------------------------------
' Paste this entire block into the SAME module as the CMS XT
' export macro (these are Private, so they must live in that
' module, not a separate one).
'
' WHY REV 2
' ---------
' Rev 1 rotated every LEAF PART component. That is wrong.
' Component2.Transform2 for a child of a sub-assembly is relative
' to its PARENT, not to the top-level assembly. Rotating leaves
' AND their parents applies R twice to anything nested -- R*R --
' which shears the mold into a scrambled pile. The module's own
' comment already said this:
'
'   ' TOP-LEVEL ONLY. GetComponents(False) returns every level, so
'   ' a child of a sub-assembly was rotated twice ...
'
' Rev 2 does what "select all and rotate" actually means:
'
'   1. suppress the mate group  (otherwise the solver drags
'      components back and scrambles them on rebuild)
'   2. UNFIX every top-level component  (fixed components silently
'      refuse Transform2, so some move and some don't -> mixed up)
'   3. apply ONE rigid transform, about the assembly centre, to
'      every TOP-LEVEL component only. Children ride along inside
'      their parent, untouched. One movement, everything together.
'   4. REFIX every top-level component so nothing can drift.
'
' Requires (already present in the macro):
'   LogLine, FormatNumberForCsv, StabilizeActiveView,
'   UnsuppressAllAssemblyComponents, ShowAllAssemblyComponents,
'   PersistCurrentViewAsStandardTop, DefineStandardFrontFromHolderAndPotCom,
'   swDocPART, swDocASSEMBLY, CMS_TOP_VIEW_NAME,
'   AUTO_DEFINE_FRONT_FROM_HOLDER_POT_COM, swApp
' ============================================================


' ------------------------------------------------------------
' Settings.
' These are functions, not Private Const, because VBA only accepts
' module-level Const in the declarations section at the very top of
' the module -- and this block gets appended at the bottom.
' ------------------------------------------------------------

' Suppress the mate group before rotating. Leave True: with mates live,
' the solver undoes the rotation component by component on rebuild.
Private Function OPT_SUPPRESS_MATES_BEFORE_ROTATION() As Boolean
    OPT_SUPPRESS_MATES_BEFORE_ROTATION = True
End Function

' Fix every top-level component again after rotating, so nothing drifts.
Private Function OPT_REFIX_COMPONENTS_AFTER_ROTATION() As Boolean
    OPT_REFIX_COMPONENTS_AFTER_ROTATION = True
End Function


' ------------------------------------------------------------
' Rotate the WHOLE assembly as one rigid set.
' Replaces the old per-component loop in
' StraightenAssemblyFromPlateTransform and ApplyMoldOrientationMatrix.
' ------------------------------------------------------------
Private Function RotateWholeAssemblyAsOneRigidSetByMatrix(ByVal assyModel As Object, _
                                                         ByRef R() As Double) As Boolean
On Error GoTo ErrHandler

    RotateWholeAssemblyAsOneRigidSetByMatrix = False

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim swMathUtil As Object
    Set swMathUtil = swApp.GetMathUtility
    If swMathUtil Is Nothing Then
        LogLine "RotateWholeAssembly failed: GetMathUtility returned Nothing."
        Exit Function
    End If

    UnsuppressAllAssemblyComponents assyModel
    ShowAllAssemblyComponents assyModel

    ' ---- TOP LEVEL ONLY -------------------------------------------------
    ' GetComponents(True) = top level only. Each top-level component moves
    ' rigidly WITH its children, so a sub-assembly's parts keep their
    ' relative positions exactly. Never use GetComponents(False) here:
    ' Transform2 on a nested child is relative to its parent, so rotating
    ' both applies R twice and shears the mold apart.
    Dim vTop As Variant
    vTop = assyModel.GetComponents(True)
    If IsEmpty(vTop) Then Exit Function
    If IsArray(vTop) = False Then Exit Function

    LogLine "RotateWholeAssembly: " & CStr(UBound(vTop) + 1) & _
            " top-level component(s) will move as one rigid set."

    ' ---- 1. kill the mates so the solver cannot fight us ----------------
    If OPT_SUPPRESS_MATES_BEFORE_ROTATION() Then
        Dim nMates As Long
        nMates = SuppressAllAssemblyMates(assyModel)
        If nMates > 0 Then
            LogLine "RotateWholeAssembly: suppressed " & CStr(nMates) & _
                    " mate group(s) so the solver cannot undo the rotation."
        End If
    End If

    ' ---- 2. unfix everything (the Ctrl+A / float step) ------------------
    Dim nSel As Long
    nSel = SelectAllTopLevelComponents(assyModel, vTop)
    If nSel < 1 Then
        LogLine "RotateWholeAssembly failed: could not select any top-level component."
        assyModel.ClearSelection2 True
        Exit Function
    End If

    SetFixedStateOnSelectedComponents assyModel, False      ' float / unfix
    assyModel.ClearSelection2 True

    ' ---- 3. ONE rigid transform, about the assembly centre -------------
    Dim xform As Object
    Set xform = BuildSolidWorksTransformAboutAssemblyCenter(assyModel, R, swMathUtil)
    If xform Is Nothing Then
        LogLine "RotateWholeAssembly failed: transform was Nothing."
        Exit Function
    End If

    Dim i As Long
    Dim c As Object
    Dim moved As Long
    Dim failed As Long
    moved = 0
    failed = 0

    For i = 0 To UBound(vTop)
        Set c = vTop(i)
        If Not c Is Nothing Then
            If c.IsSuppressed = False Then
                If ApplyRigidTransformToOneComponent(c, xform) Then
                    moved = moved + 1
                Else
                    failed = failed + 1
                    LogLine "RotateWholeAssembly: transform refused by " & c.Name2
                End If
            End If
        End If
    Next i

    LogLine "RotateWholeAssembly: moved " & CStr(moved) & _
            " top-level component(s), " & CStr(failed) & " refused."

    If moved = 0 Then
        LogLine "RotateWholeAssembly failed: no component accepted the transform."
        Exit Function
    End If

    On Error Resume Next
    assyModel.EditRebuild3
    On Error GoTo ErrHandler

    ' ---- 4. refix so nothing can drift ---------------------------------
    If OPT_REFIX_COMPONENTS_AFTER_ROTATION() Then
        If SelectAllTopLevelComponents(assyModel, vTop) > 0 Then
            SetFixedStateOnSelectedComponents assyModel, True    ' fix
        End If
        assyModel.ClearSelection2 True
    End If

    On Error Resume Next
    assyModel.EditRebuild3
    assyModel.ForceRebuild3 False
    assyModel.GraphicsRedraw2
    On Error GoTo ErrHandler

    LogLine "RotateWholeAssembly: done. Whole mold rotated as one rigid set."
    RotateWholeAssemblyAsOneRigidSetByMatrix = True
    Exit Function

ErrHandler:
    LogLine "RotateWholeAssemblyAsOneRigidSetByMatrix error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    RotateWholeAssemblyAsOneRigidSetByMatrix = False
End Function


' ------------------------------------------------------------
' Select every top-level component. This is the Ctrl+A equivalent.
' Sub-assemblies are selected as whole containers, NOT drilled into.
' ------------------------------------------------------------
Private Function SelectAllTopLevelComponents(ByVal assyModel As Object, _
                                             ByRef vTop As Variant) As Long
On Error GoTo ErrHandler

    SelectAllTopLevelComponents = 0

    If assyModel Is Nothing Then Exit Function
    If IsArray(vTop) = False Then Exit Function

    assyModel.ClearSelection2 True

    Dim i As Long
    Dim c As Object
    Dim n As Long
    n = 0

    For i = 0 To UBound(vTop)
        Set c = vTop(i)
        If Not c Is Nothing Then
            If c.IsSuppressed = False Then
                If c.Select4(True, Nothing, False) Then n = n + 1
            End If
        End If
    Next i

    SelectAllTopLevelComponents = n
    Exit Function

ErrHandler:
    LogLine "SelectAllTopLevelComponents error: " & Err.Description
    SelectAllTopLevelComponents = 0
End Function


' ------------------------------------------------------------
' Fix (True) or float/unfix (False) whatever is currently selected.
' ------------------------------------------------------------
Private Sub SetFixedStateOnSelectedComponents(ByVal assyModel As Object, _
                                              ByVal makeFixed As Boolean)
On Error GoTo ErrHandler

    If assyModel Is Nothing Then Exit Sub
    If assyModel.GetType <> swDocASSEMBLY Then Exit Sub

    On Error Resume Next
    Err.Clear

    If makeFixed Then
        assyModel.FixComponent
    Else
        assyModel.UnfixComponent
    End If

    If Err.Number <> 0 Then
        LogLine "SetFixedStateOnSelectedComponents: " & _
                IIf(makeFixed, "FixComponent", "UnfixComponent") & _
                " failed: " & Err.Description
        Err.Clear
    Else
        LogLine "SetFixedStateOnSelectedComponents: selection " & _
                IIf(makeFixed, "fixed.", "floated/unfixed.")
    End If

    On Error GoTo ErrHandler
    Exit Sub

ErrHandler:
    LogLine "SetFixedStateOnSelectedComponents error: " & Err.Description
End Sub


' ------------------------------------------------------------
' Suppress the assembly's mate group(s).
'
' This is the other reason a rotated assembly comes out scrambled:
' the transforms get set correctly, then EditRebuild3 runs the mate
' solver, which pulls every mated component back toward geometry
' that has moved. Suppressing the mate group makes the rotation stick.
' Returns the number of mate groups suppressed.
' ------------------------------------------------------------
Private Function SuppressAllAssemblyMates(ByVal assyModel As Object) As Long
On Error GoTo ErrHandler

    SuppressAllAssemblyMates = 0

    If assyModel Is Nothing Then Exit Function
    If assyModel.GetType <> swDocASSEMBLY Then Exit Function

    Dim feat As Object
    Dim tn As String
    Dim n As Long
    n = 0

    Set feat = assyModel.FirstFeature

    Do While Not feat Is Nothing
        tn = ""
        On Error Resume Next
        tn = feat.GetTypeName2
        On Error GoTo ErrHandler

        If tn = "MateGroup" Then
            On Error Resume Next
            Err.Clear
            assyModel.ClearSelection2 True
            feat.Select2 False, -1
            assyModel.EditSuppress2
            If Err.Number = 0 Then n = n + 1
            Err.Clear
            On Error GoTo ErrHandler
        End If

        Set feat = feat.GetNextFeature
    Loop

    assyModel.ClearSelection2 True
    SuppressAllAssemblyMates = n
    Exit Function

ErrHandler:
    LogLine "SuppressAllAssemblyMates error: " & Err.Description
    On Error Resume Next
    assyModel.ClearSelection2 True
    SuppressAllAssemblyMates = n
End Function


' ------------------------------------------------------------
' Build ONE rigid rotation transform about the assembly centre.
' Rotating about the origin can swing the whole mold off into space.
' ------------------------------------------------------------
Private Function BuildSolidWorksTransformAboutAssemblyCenter(ByVal assyModel As Object, _
                                                             ByRef R() As Double, _
                                                             ByVal swMathUtil As Object) As Object
On Error GoTo ErrHandler

    Set BuildSolidWorksTransformAboutAssemblyCenter = Nothing

    If assyModel Is Nothing Then Exit Function
    If swMathUtil Is Nothing Then Exit Function

    Dim arr(0 To 15) As Double

    ' SolidWorks transform convention used throughout this macro:
    '   x' = x*m0 + y*m3 + z*m6 + m9
    '   y' = x*m1 + y*m4 + z*m7 + m10
    '   z' = x*m2 + y*m5 + z*m8 + m11
    '
    ' R rows = new output axes.
    arr(0) = R(0, 0)
    arr(1) = R(1, 0)
    arr(2) = R(2, 0)

    arr(3) = R(0, 1)
    arr(4) = R(1, 1)
    arr(5) = R(2, 1)

    arr(6) = R(0, 2)
    arr(7) = R(1, 2)
    arr(8) = R(2, 2)

    Dim cx As Double
    Dim cy As Double
    Dim cz As Double

    If TryGetAssemblyCenterMeters(assyModel, cx, cy, cz) Then
        ' Rotate about centre:
        '   p' = R * (p - C) + C
        '   translation = C - R*C
        arr(9) = cx - ((R(0, 0) * cx) + (R(0, 1) * cy) + (R(0, 2) * cz))
        arr(10) = cy - ((R(1, 0) * cx) + (R(1, 1) * cy) + (R(1, 2) * cz))
        arr(11) = cz - ((R(2, 0) * cx) + (R(2, 1) * cy) + (R(2, 2) * cz))

        LogLine "BuildTransform: rotating about assembly centre metres " & _
                FormatNumberForCsv(cx) & "/" & _
                FormatNumberForCsv(cy) & "/" & _
                FormatNumberForCsv(cz)
    Else
        arr(9) = 0#
        arr(10) = 0#
        arr(11) = 0#
        LogLine "BuildTransform WARNING: assembly centre unavailable; rotating about origin."
    End If

    arr(12) = 1#
    arr(13) = 0#
    arr(14) = 0#
    arr(15) = 0#

    Set BuildSolidWorksTransformAboutAssemblyCenter = swMathUtil.CreateTransform(arr)
    Exit Function

ErrHandler:
    LogLine "BuildSolidWorksTransformAboutAssemblyCenter error: " & Err.Description
    Set BuildSolidWorksTransformAboutAssemblyCenter = Nothing
End Function


' Assembly centre in METRES, in assembly coordinates.
'
' NOTE: ModelDoc2/AssemblyDoc does not reliably expose GetBox for an assembly
' (the two-argument GetBox used elsewhere in this macro is the PartDoc form).
' The dependable path is to union the assembly-space boxes of the leaf part
' components -- Component2.GetBox(False, False) is already proven in this
' module. This only MEASURES, it never moves anything, so walking all levels
' here is safe. The ModelDoc GetBox forms are kept as fallbacks.
Private Function TryGetAssemblyCenterMeters(ByVal assyModel As Object, _
                                            ByRef cx As Double, _
                                            ByRef cy As Double, _
                                            ByRef cz As Double) As Boolean
On Error GoTo ErrHandler

    TryGetAssemblyCenterMeters = False

    If assyModel Is Nothing Then Exit Function

    ' ---- 1. union of leaf part component boxes (preferred) -------------
    Dim vComps As Variant
    Dim i As Long
    Dim c As Object
    Dim m As Object
    Dim vb As Variant
    Dim haveBox As Boolean

    Dim xLo As Double, yLo As Double, zLo As Double
    Dim xHi As Double, yHi As Double, zHi As Double

    haveBox = False

    On Error Resume Next
    vComps = assyModel.GetComponents(False)
    On Error GoTo ErrHandler

    If IsArray(vComps) Then
        For i = 0 To UBound(vComps)
            Set c = vComps(i)
            If Not c Is Nothing Then
                If c.IsSuppressed = False Then
                    Set m = c.GetModelDoc2
                    If Not m Is Nothing Then
                        If m.GetType = swDocPART Then
                            vb = Empty
                            On Error Resume Next
                            vb = c.GetBox(False, False)
                            On Error GoTo ErrHandler

                            If IsArray(vb) Then
                                If UBound(vb) >= 5 Then
                                    If haveBox = False Then
                                        xLo = CDbl(vb(0)): yLo = CDbl(vb(1)): zLo = CDbl(vb(2))
                                        xHi = CDbl(vb(3)): yHi = CDbl(vb(4)): zHi = CDbl(vb(5))
                                        haveBox = True
                                    Else
                                        If CDbl(vb(0)) < xLo Then xLo = CDbl(vb(0))
                                        If CDbl(vb(1)) < yLo Then yLo = CDbl(vb(1))
                                        If CDbl(vb(2)) < zLo Then zLo = CDbl(vb(2))
                                        If CDbl(vb(3)) > xHi Then xHi = CDbl(vb(3))
                                        If CDbl(vb(4)) > yHi Then yHi = CDbl(vb(4))
                                        If CDbl(vb(5)) > zHi Then zHi = CDbl(vb(5))
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next i
    End If

    If haveBox Then
        cx = (xLo + xHi) / 2#
        cy = (yLo + yHi) / 2#
        cz = (zLo + zHi) / 2#
        TryGetAssemblyCenterMeters = True
        Exit Function
    End If

    ' ---- 2. fallbacks: ModelDoc-level box ------------------------------
    Dim vBox As Variant

    vBox = Empty
    On Error Resume Next
    vBox = assyModel.GetBox(0)
    On Error GoTo ErrHandler

    If IsArray(vBox) = False Then
        vBox = Empty
        On Error Resume Next
        vBox = assyModel.GetBox(False, False)
        On Error GoTo ErrHandler
    End If

    If IsArray(vBox) = False Then Exit Function
    If UBound(vBox) < 5 Then Exit Function

    cx = (CDbl(vBox(0)) + CDbl(vBox(3))) / 2#
    cy = (CDbl(vBox(1)) + CDbl(vBox(4))) / 2#
    cz = (CDbl(vBox(2)) + CDbl(vBox(5))) / 2#

    TryGetAssemblyCenterMeters = True
    Exit Function

ErrHandler:
    LogLine "TryGetAssemblyCenterMeters error: " & Err.Description
    TryGetAssemblyCenterMeters = False
End Function


' new = global * old, composed explicitly.
'
' NOTE: IMathTransform.Multiply is deliberately NOT used here. Its operand
' order differs between SolidWorks versions, and getting it backwards
' silently scrambles the assembly. This does the same 4x4 composition that
' the proven StraightenOneComponentByMatrix already does, with the same
' index convention:
'   x' = x*m0 + y*m3 + z*m6 + m9
'   y' = x*m1 + y*m4 + z*m7 + m10
'   z' = x*m2 + y*m5 + z*m8 + m11
Private Function ApplyRigidTransformToOneComponent(ByVal swComp As Object, _
                                                   ByVal globalXform As Object) As Boolean
On Error GoTo ErrHandler

    ApplyRigidTransformToOneComponent = False

    If swComp Is Nothing Then Exit Function
    If globalXform Is Nothing Then Exit Function

    Dim swMathUtil As Object
    Set swMathUtil = swApp.GetMathUtility
    If swMathUtil Is Nothing Then Exit Function

    Dim oldX As Object
    Set oldX = swComp.Transform2
    If oldX Is Nothing Then Exit Function

    Dim vg As Variant
    Dim vo As Variant
    vg = globalXform.ArrayData
    vo = oldX.ArrayData
    If IsArray(vg) = False Or IsArray(vo) = False Then Exit Function
    If UBound(vg) < 12 Or UBound(vo) < 12 Then Exit Function

    Dim g(0 To 15) As Double
    Dim o(0 To 15) As Double
    Dim k As Long
    For k = 0 To 15
        If k <= UBound(vg) Then g(k) = CDbl(vg(k)) Else g(k) = 0#
        If k <= UBound(vo) Then o(k) = CDbl(vo(k)) Else o(k) = 0#
    Next k

    Dim n(0 To 15) As Double
    Dim col As Long
    Dim base As Long
    Dim i As Long
    Dim s As Double

    ' Rotation columns 0..2, then the translation at 9..11.
    For col = 0 To 3
        base = col * 3
        If col = 3 Then base = 9
        For k = 0 To 2
            s = 0#
            For i = 0 To 2
                s = s + g(k + 3 * i) * o(base + i)
            Next i
            If col = 3 Then s = s + g(9 + k)   ' translation picks up the offset
            n(base + k) = s
        Next k
    Next col

    n(12) = o(12)                              ' preserve scale
    If n(12) = 0# Then n(12) = 1#
    n(13) = 0#: n(14) = 0#: n(15) = 0#

    Dim newX As Object
    Set newX = swMathUtil.CreateTransform(n)
    If newX Is Nothing Then Exit Function

    swComp.Transform2 = newX

    ApplyRigidTransformToOneComponent = True
    Exit Function

ErrHandler:
    ApplyRigidTransformToOneComponent = False
End Function


' ------------------------------------------------------------
' Face/view correction AFTER BOM matching.
' The geometry was already physically rotated, so *Top is a true top and
' no snap-to-nearest-view guessing is needed. Running this after matching
' is what keeps ExportRows(i).CadPartIndex valid.
' ------------------------------------------------------------
Private Sub FinalizeFacesAfterPhysicalRotation(ByVal model As Object)
On Error GoTo ErrHandler

    If model Is Nothing Then Exit Sub
    If model.GetType <> swDocASSEMBLY Then Exit Sub

    LogLine "Finalizing faces after physical rotation."

    ' TCP side becomes the true SolidWorks *Top.
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 100

    If PersistCurrentViewAsStandardTop(model) Then
        LogLine "Final faces: TCP side persisted as SolidWorks *Top."
    Else
        LogLine "WARNING: Final faces could not persist SolidWorks *Top."
    End If

    ' Pot side becomes the true SolidWorks *Front.
    If AUTO_DEFINE_FRONT_FROM_HOLDER_POT_COM Then
        If DefineStandardFrontFromHolderAndPotCom(model) Then
            LogLine "Final faces: pots are closer than holders in SolidWorks *Front."
        Else
            LogLine "WARNING: Final faces could not define *Front from pot/holder geometry."
        End If
    End If

    ' Save CMS_TOP from the final corrected SolidWorks *Top.
    model.ShowNamedView2 "*Top", 5
    StabilizeActiveView model, 100

    On Error Resume Next
    model.DeleteNamedView CMS_TOP_VIEW_NAME
    Err.Clear
    model.NameView CMS_TOP_VIEW_NAME
    On Error GoTo ErrHandler

    model.ShowNamedView2 CMS_TOP_VIEW_NAME, -1
    StabilizeActiveView model, 100

    LogLine "Final faces complete: CMS_TOP saved from final corrected *Top."
    Exit Sub

ErrHandler:
    LogLine "FinalizeFacesAfterPhysicalRotation error: " & Err.Description
End Sub

' ============================================================
' END OF CMS ORIENTATION PATCH
' ============================================================
