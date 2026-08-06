<#
====================================================================
 Apply_CMS_Orientation_Patch.ps1
--------------------------------------------------------------------
 Applies the "rotate every leaf part component" + "faces after BOM
 matching" patch to the CMS XT export macro (gemini1.bas).

 What it changes:
   1. ProcessOneJob : resets gMoldGeometryIsSquareAndUpright = False
   2. StraightenAssemblyFromPlateTransform : per-component loop ->
      RotateEveryLeafPartComponentTogetherByMatrix
   3. ApplyMoldOrientationMatrix : same loop replacement
   4. ProcessOneJob : the *Top block ->
      FinalizeFacesAfterPhysicalRotation
   5. Appends all new helper functions before "END OF MODULE"

 Usage (PowerShell, from this folder):
   .\Apply_CMS_Orientation_Patch.ps1 -SourceBas "C:\path\to\gemini1.bas"

 Output:
   <source>_PATCHED.bas   (original file is never modified)
====================================================================
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$SourceBas,

    [string]$NewFunctionsBas = (Join-Path $PSScriptRoot 'CMS_Orientation_Patch_NewFunctions.bas'),

    [string]$OutFile
)

$ErrorActionPreference = 'Stop'

function Fail([string]$msg) {
    Write-Host "FAILED: $msg" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path -LiteralPath $SourceBas)) { Fail "Source file not found: $SourceBas" }
if (-not (Test-Path -LiteralPath $NewFunctionsBas)) { Fail "New-functions file not found: $NewFunctionsBas" }

if (-not $OutFile) {
    $dir  = Split-Path -Parent (Resolve-Path -LiteralPath $SourceBas)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($SourceBas)
    $OutFile = Join-Path $dir ($base + '_PATCHED.bas')
}

# ---- load, normalise line endings to LF for matching ----------------
$text = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $SourceBas))
$text = $text -replace "`r`n", "`n"
$text = $text -replace "`r", "`n"

$newFuncs = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $NewFunctionsBas))
$newFuncs = $newFuncs -replace "`r`n", "`n"
$newFuncs = $newFuncs -replace "`r", "`n"

# ---- idempotency guard ---------------------------------------------
if ($text -match 'RotateEveryLeafPartComponentTogetherByMatrix') {
    Fail "This file already contains RotateEveryLeafPartComponentTogetherByMatrix. Patch appears to be applied already."
}

function LF([string]$s) { ($s -replace "`r`n", "`n") -replace "`r", "`n" }

function Replace-Block {
    param(
        [string]$Label,
        [string]$Haystack,
        [string]$Old,
        [string]$New,
        [int]$ExpectedCount = 1
    )
    $o = LF $Old
    $n = LF $New

    $count = 0
    $idx = 0
    while (($idx = $Haystack.IndexOf($o, $idx)) -ge 0) {
        $count++
        $idx += $o.Length
    }

    if ($count -ne $ExpectedCount) {
        Fail "[$Label] expected $ExpectedCount match(es) but found $count. File differs from the version this patch was written against - nothing was written."
    }

    Write-Host ("  OK  [{0}] {1} replacement(s)" -f $Label, $count) -ForegroundColor Green
    return $Haystack.Replace($o, $n)
}

Write-Host "Patching: $SourceBas" -ForegroundColor Cyan

# ====================================================================
# 1. Reset the flag every job
# ====================================================================
$old1 = @'
    FinalStlCoordFrameReady = False

    Dim stlCoordI As Long
'@

$new1 = @'
    FinalStlCoordFrameReady = False
    gMoldGeometryIsSquareAndUpright = False

    Dim stlCoordI As Long
'@

$text = Replace-Block -Label 'flag reset' -Haystack $text -Old $old1 -New $new1 -ExpectedCount 1

# ====================================================================
# 2. StraightenAssemblyFromPlateTransform : per-component rotation loop
#    NOTE the loop text also appears in the never-called dead function
#    OrientMoldToTcpTopAndPotFront, so each replacement is anchored on
#    its own surrounding lines to guarantee it hits the right one.
# ====================================================================
$old2 = @'
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
'@

$new2 = @'
    Dim moved As Long, failed As Long
    moved = 0: failed = 0

    ' PATCH: rotate every leaf PART component as one rigid set instead of
    ' walking components one by one. Subassembly containers are skipped so
    ' nested parts are not double-rotated.
    If RotateEveryLeafPartComponentTogetherByMatrix(model, R) Then
        moved = 1
        failed = 0
    Else
        moved = 0
        failed = 1
    End If

    LogLine "STRAIGHTEN: rotated " & CStr(moved) & " component(s), " & _
'@

$text = Replace-Block -Label 'StraightenAssemblyFromPlateTransform loop' -Haystack $text -Old $old2 -New $new2 -ExpectedCount 1

# ====================================================================
# 3. ApplyMoldOrientationMatrix : per-component rotation loop
# ====================================================================
$old3 = @'
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
'@

$new3 = @'
    Dim vComps As Variant
    vComps = model.GetComponents(True)
    If IsEmpty(vComps) Then Exit Function

    Dim moved As Long, failed As Long
    Dim ci As Long
    moved = 0: failed = 0

    ' PATCH: rotate every leaf PART component as one rigid set instead of
    ' walking components one by one. Subassembly containers are skipped so
    ' nested parts are not double-rotated.
    If RotateEveryLeafPartComponentTogetherByMatrix(model, R) Then
        moved = 1
        failed = 0
    Else
        moved = 0
        failed = 1
    End If

    LogLine "ORIENT: rotated " & CStr(moved) & " top-level component(s), " & _
'@

$text = Replace-Block -Label 'ApplyMoldOrientationMatrix loop' -Haystack $text -Old $old3 -New $new3 -ExpectedCount 1

# ====================================================================
# 4. Face/view correction after BOM matching
# ====================================================================
$old4 = @'
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
'@

$new4 = @'
    ' PATCH: physical rotations already happened BEFORE BOM matching, so the
    ' CadPartIndex values the matcher stored are still valid. All that is left
    ' here is face/view correction:
    '     *Top   = TCP side
    '     *Front = pots closer than holders
    If gMoldGeometryIsSquareAndUpright Then
        FinalizeFacesAfterPhysicalRotation swModel
    Else
        EnsureCmsTopOrientationFromMatchedTcpBcp swModel, PERSIST_CMS_TOP_AS_STANDARD_VIEWS_BEFORE_BASE_SAVE
    End If
'@

$text = Replace-Block -Label 'faces after BOM match' -Haystack $text -Old $old4 -New $new4 -ExpectedCount 1

# ====================================================================
# 5. Append the new functions before END OF MODULE
# ====================================================================
$anchor = @'
' ============================================================
' END OF MODULE
' ============================================================
'@

$anchorLf = LF $anchor
$anchorCount = 0
$i = 0
while (($i = $text.IndexOf($anchorLf, $i)) -ge 0) { $anchorCount++; $i += $anchorLf.Length }

if ($anchorCount -eq 1) {
    $text = $text.Replace($anchorLf, ($newFuncs.TrimEnd("`n") + "`n`n" + $anchorLf))
    Write-Host "  OK  [new functions] inserted before END OF MODULE" -ForegroundColor Green
}
else {
    $text = $text.TrimEnd("`n") + "`n`n" + $newFuncs.TrimEnd("`n") + "`n"
    Write-Host "  OK  [new functions] appended at end of file (END OF MODULE anchor count = $anchorCount)" -ForegroundColor Yellow
}

# ---- sanity checks --------------------------------------------------
$required = @(
    'RotateEveryLeafPartComponentTogetherByMatrix',
    'SelectEveryLeafPartComponent',
    'FloatSelectedComponents',
    'BuildSolidWorksTransformAboutAssemblyCenter',
    'TryGetAssemblyCenterMeters',
    'ApplyTransformToSelectedComponents',
    'ApplyTransformToEveryLeafPartComponent',
    'ApplyRigidTransformToOneComponent',
    'FinalizeFacesAfterPhysicalRotation'
)
foreach ($r in $required) {
    if ($text -notmatch [regex]::Escape("Private Function $r") -and
        $text -notmatch [regex]::Escape("Private Sub $r")) {
        Fail "sanity check: $r is missing from the patched output."
    }
}

# ---- write out with CRLF (what the VBA editor expects) --------------
$out = $text -replace "`n", "`r`n"
[System.IO.File]::WriteAllText($OutFile, $out, [System.Text.Encoding]::Default)

Write-Host ""
Write-Host "Patched file written to:" -ForegroundColor Cyan
Write-Host "  $OutFile"
Write-Host ""
Write-Host "Next: in the SolidWorks VBA editor, remove the old module and" -ForegroundColor Cyan
Write-Host "File > Import File... the patched .bas, then Debug > Compile." -ForegroundColor Cyan
