Attribute VB_Name = "AI_Bridge_Import"
Option Explicit

' =============================================================================
' AI_Bridge_Import.bas
'
' PASTE-IN HELPER MODULE -- this is NOT a replacement for Module6121. It is a
' new, separate module meant to be dropped into the same workbook/project as
' Module6121 (or have its one public Sub called FROM Module6121).
'
' WHY THIS EXISTS
' ----------------
' The CMS AI Quoting web app (webapp/backend) classifies every job's CAD
' export and writes a flat, VBA-friendly "bridge" file every time you open a
' quote in the app:
'
'   <CMS_VBA_BRIDGE_DIR>\<JobID>_part_names.csv
'   <CMS_VBA_BRIDGE_DIR>\<JobID>_part_names.json
'
' CMS_VBA_BRIDGE_DIR defaults to webapp/backend/data/vba_bridge on whatever
' machine runs the backend. Point BRIDGE_FOLDER below at that folder (a local
' path if the backend runs on this PC, or a shared/mapped network path if it
' runs elsewhere).
'
' The CSV columns are:
'   Index, Component, Role, ResolvedName, Quote, Price, SecondaryPartingLine
'
'   Role                 machine role key (a_plate, b_plate, bottom_ejector_plate,
'                        latch_lock, leader_pin, ...) -- matches
'                        geometry_classifier/qwen_classify_xt_csv.py ROLES.
'   ResolvedName         human-readable CMS name the AI wants placed on the
'                        quote (e.g. "A Plate", "Bottom Ejector Plate",
'                        "Rail 1", "Leader Pin 3"). Already numbered for
'                        hardware that repeats; plates stay singular.
'   Quote                TRUE/FALSE -- whether the AI recommends this row be
'                        priced/shown as its own quote line.
'   Price                AI-computed price for this row (placeholder rates
'                        until you wire in real CMS pricing -- see
'                        webapp/backend/app/pricing.py).
'   SecondaryPartingLine TRUE for latch-lock/PLC/safety-strap hardware on a
'                        plate-sequenced base -- these mark a secondary
'                        opening/parting line and must NOT be used to flip
'                        the main A/B plate assignment.
'
' TODO FOR YOU (since I don't have your real Module6121.bas yet):
'   1. Set BRIDGE_FOLDER below.
'   2. Call CMS_AI_ImportPartNames from Module6121 at the point where it
'      currently looks up/types in plate and hardware names, passing the
'      active job's ID.
'   3. Replace the "TODO: wire into your real quote cell/named-range" line
'      inside WriteRowToQuote with your actual cell references or named
'      ranges (this is the one part I cannot do without seeing your macro --
'      paste Module6121.bas into the chat or add it to the repo and I will
'      wire this in directly, matching your exact Subs/variable names).
' =============================================================================

Private Const BRIDGE_FOLDER As String = "C:\CMS_AI\webapp_bridge\"  ' <-- set me
Private Const IMPORT_SHEET_NAME As String = "AI_Import"

' Entry point: call this with the active job's ID (e.g. "J8420", "T001015").
' It reads <BRIDGE_FOLDER><JobID>_part_names.csv and both (a) writes a clean
' AI_Import worksheet you can eyeball/copy from, and (b) calls WriteRowToQuote
' per row so you can wire it into your real quote cells.
Public Sub CMS_AI_ImportPartNames(ByVal JobID As String)
    Dim csvPath As String
    csvPath = BRIDGE_FOLDER & JobID & "_part_names.csv"

    If Dir(csvPath) = "" Then
        MsgBox "AI bridge file not found:" & vbCrLf & csvPath & vbCrLf & _
               "Open this job in the CMS AI Quoting app first (Quotes > " & JobID & " > Parts & Pricing) " & _
               "so it can export the bridge file, then try again.", vbExclamation, "CMS AI Bridge"
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = GetOrCreateImportSheet()

    Dim fnum As Integer
    fnum = FreeFile
    Open csvPath For Input As #fnum

    Dim lineText As String
    Dim rowIndex As Long
    rowIndex = 1

    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        Dim fields() As String
        fields = SplitCsvLine(lineText)

        If rowIndex = 1 Then
            ' Header row
            WriteHeaderRow ws, fields
        Else
            WriteImportRow ws, rowIndex, fields
            WriteRowToQuote JobID, fields
        End If

        rowIndex = rowIndex + 1
    Loop

    Close #fnum

    MsgBox "Imported " & (rowIndex - 2) & " AI-resolved part names for job " & JobID & "." & vbCrLf & _
           "See the '" & IMPORT_SHEET_NAME & "' sheet, and check WriteRowToQuote for wiring into your quote cells.", _
           vbInformation, "CMS AI Bridge"
End Sub

Private Function GetOrCreateImportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(IMPORT_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = IMPORT_SHEET_NAME
    Else
        ws.Cells.Clear
    End If

    Set GetOrCreateImportSheet = ws
End Function

Private Sub WriteHeaderRow(ByVal ws As Worksheet, ByVal fields() As String)
    Dim c As Long
    For c = LBound(fields) To UBound(fields)
        ws.Cells(1, c + 1).Value = fields(c)
        ws.Cells(1, c + 1).Font.Bold = True
    Next c
End Sub

Private Sub WriteImportRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal fields() As String)
    Dim c As Long
    For c = LBound(fields) To UBound(fields)
        ws.Cells(rowIndex, c + 1).Value = fields(c)
    Next c
End Sub

' Column order matches the bridge CSV header:
' 0=Index 1=Component 2=Role 3=ResolvedName 4=Quote 5=Price 6=SecondaryPartingLine
Private Sub WriteRowToQuote(ByVal JobID As String, ByVal fields() As String)
    If UBound(fields) < 6 Then Exit Sub

    Dim role As String, resolvedName As String, quoteFlag As String
    Dim price As Double, secondaryParting As Boolean

    role = fields(2)
    resolvedName = fields(3)
    quoteFlag = fields(4)
    price = Val(fields(5))
    secondaryParting = (UCase$(fields(6)) = "TRUE")

    ' Latch-lock / secondary-parting-line hardware should never be used to
    ' decide A/B plate identity or flip stack orientation -- it only marks a
    ' secondary opening line. Skip it here unless you specifically quote
    ' latch-lock assemblies as their own line item.
    If secondaryParting Then Exit Sub

    If UCase$(quoteFlag) <> "TRUE" Then Exit Sub

    ' ------------------------------------------------------------------
    ' TODO: wire into your real quote cell/named-range here. Example shape
    ' once you share Module6121's actual cell layout:
    '
    '   Select Case role
    '       Case "a_plate":               Range("QuoteSheet!B12").Value = resolvedName
    '       Case "b_plate":                Range("QuoteSheet!B13").Value = resolvedName
    '       Case "bottom_ejector_plate":   Range("QuoteSheet!B20").Value = resolvedName
    '       Case "ejector_plate":          Range("QuoteSheet!B19").Value = resolvedName
    '       Case "latch_lock":             ' see secondaryParting guard above
    '       Case Else                      ' fall through to generic hardware rows
    '   End Select
    '
    ' Until that's wired in, this Sub is a documented no-op placeholder so it
    ' compiles cleanly and is safe to call.
    ' ------------------------------------------------------------------
End Sub

' Minimal CSV line splitter (no embedded commas/quotes expected in this
' export -- component paths use "/" as a separator, not ","). If you later
' add commas to component names, swap this for a proper CSV parser.
Private Function SplitCsvLine(ByVal lineText As String) As String()
    SplitCsvLine = Split(lineText, ",")
End Function
