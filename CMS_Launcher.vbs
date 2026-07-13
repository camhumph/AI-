' ============================================================
' CMS_Launcher.vbs
' ----------------------------------------------------------
' Double-click this on the desktop to start a quote job.
' Does everything that does NOT need SolidWorks:
'   1. Asks for the C-number (job #)
'   2. Paste-box for the Gmail quote email
'   3. Parses customer job #, ship date, similar-to reference
'   4. Assigns the next quote number from the proposals folder
'   5. Fills and saves the customer proposal (Excel)
'   6. Writes a small handoff file so Module6121 knows the
'      assigned quote # and customer info without re-asking
'   7. Launches SolidWorks and runs Module6121 automatically
'
' SETTINGS  (edit these to match your machine):
' ============================================================

Const DOWNLOADS_FOLDER       = "C:\Users\lenovo\Downloads"
Const LOCAL_WORKSPACE_ROOT   = "C:\CMS_Local_Workspace"
Const QUOTE_PROPOSALS_FOLDER = "\\Mycloudex2ultra\mexico\Cameron's stuff\RON'S QUOTES\Quote-Proposals-2026"
Const JOB_ROOT_BASE          = "\\Mycloudex2ultra\mexico\Cameron's stuff\RON'S QUOTES"   ' Ron-only month folders live here
Const DEFAULT_CUSTOMER_PREFIX = ""                                     ' BMS only when the email/file actually says BMS
Const COMPANY_WIFI_SSID      = "NETGEAR"
Const PUBLIC_DATA_ROOT       = "\\Mycloudex2ultra\mexico\Cameron's stuff\Matching software"
Const PRIVATE_DATA_ROOT      = "C:\CMS_Local_Workspace\Matching"
Const SHOW_POPUPS            = False   ' unattended: log instead of stopping for OK boxes

' Proposal header defaults
Const CMS_CUSTOMER_NAME  = ""
Const CMS_ATTENTION      = "Todd Meng"
Const CMS_SALES_REP      = ""
Const CMS_PAYMENT_TERMS  = "Net 60"
Const PROPOSAL_TOTAL_CELL = "D38"
Const PROPOSAL_TOTAL_SOURCE_CELL = ""   ' leave blank until you tell us the grand-total cell

' SolidWorks paths  -- this machine has more than one version installed.
' ALWAYS use SolidWorks 2023: the "(3)" install + ProgID .31
' The plain "SOLIDWORKS\SLDWORKS.EXE" path opens 2025 — never use that.
Const SW_EXE    = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
Const SW_PROGID = "SldWorks.Application.31"   ' 31 = SolidWorks 2023 (32=2024, 33=2025)
Const SW_MACRO = "C:\CMS_Local_Workspace\Module6121.swp"   ' compiled macro — use .swp (RunMacro expects this)

' Handoff file written for Module6121 to read
Const HANDOFF_FILE = "C:\CMS_Local_Workspace\cms_handoff.txt"
Const TRAINING_TRIGGER = "C:\CMS_Local_Workspace\cms_training_xt.txt"
Const MACRO_STATUS_FILE = "C:\CMS_Local_Workspace\cms_macro_status.txt"
Const MACRO_STARTED_FILE = "C:\CMS_Local_Workspace\cms_macro_started.txt"
Const MACRO_DONE_FILE = "C:\CMS_Local_Workspace\cms_macro_done.txt"
Const MACRO_ERROR_FILE = "C:\CMS_Local_Workspace\cms_macro_error.txt"

' Gmail search (Python) settings
Const USE_GMAIL_SEARCH  = True
Const PYTHON_EXE        = "python"   ' or full path e.g. C:\Python312\python.exe
Const GMAIL_SCRIPT      = "C:\CMS_Local_Workspace\cms_gmail_search.py"
Const EMAIL_OUTPUT_FILE = "C:\CMS_Local_Workspace\cms_email.txt"

' ============================================================
' MAIN
' ============================================================
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim gAttachDir, gPreferredCNum, gCadPath, gCustomerPrefix, gCustomerName
gAttachDir = ""
gPreferredCNum = ""
gCadPath = ""
gCustomerPrefix = ""
gCustomerName = ""

' Make sure the local workspace exists (handoff + email files live here)
If Not fso.FolderExists(LOCAL_WORKSPACE_ROOT) Then fso.CreateFolder LOCAL_WORKSPACE_ROOT
If fso.FileExists(TRAINING_TRIGGER) Then
    fso.DeleteFile TRAINING_TRIGGER, True
    LogStep "cleared stale cms_training_xt.txt (live quote, not training)"
End If
' Clear stale macro launch acknowledgements from a previous cancelled/failed run.
DeleteIfExists MACRO_STATUS_FILE
DeleteIfExists MACRO_STARTED_FILE
DeleteIfExists MACRO_DONE_FILE
DeleteIfExists MACRO_ERROR_FILE
LogStep "===== launcher started ====="

' 1. Prefer an existing C-number (e.g. BMS-851100029-C18603 → C18603).
'    Only assign a new quote number when no C##### is present.
Dim cNum
cNum = ""

' /usemail : the email picker already wrote cms_email.txt for a chosen message,
'            so use that instead of searching Gmail for the newest one.
Dim gUseExistingEmail, ai
gUseExistingEmail = False
For ai = 0 To WScript.Arguments.Count - 1
    If LCase(WScript.Arguments(ai)) = "/usemail" Then gUseExistingEmail = True
Next

' Email quoting is done in the CMS AI Quoting webapp inbox (http://127.0.0.1:8000).
' Click the big blue Quote button there — this launcher only runs when the
' webapp (or an old /usemail handoff) starts it with /usemail.
If Not gUseExistingEmail Then
    Dim shellOpen
    Set shellOpen = CreateObject("WScript.Shell")
    LogStep "Opening CMS AI Quoting webapp inbox (quote emails there, not here)"
    shellOpen.Run "http://127.0.0.1:8000/email", 1, False
    WScript.Quit
End If

' 2. Get the email - search Gmail first, fall back to a paste box
Dim custJobNum, similarTo, shipDate, emailBody, gotEmail, customerPrefix, customerName
custJobNum = ""
similarTo  = ""
shipDate   = ""
gotEmail   = False
customerPrefix = ""
customerName = ""

If USE_GMAIL_SEARCH Or gUseExistingEmail Then
    gotEmail = RunGmailSearch(custJobNum, similarTo, shipDate)
End If

If Not gotEmail Then
    LogStep "no Gmail quote email data found - continuing without a paste prompt"
End If

' 3. Prefer C-number already on the job (folder / subject / attachments).
'    Example: BMS-851100029-C18603 → use C18603, do NOT assign C18635.
Dim quoteNum, quoteNoHyphen, foundC
foundC = ResolveExistingCNumber(custJobNum, gAttachDir)
If foundC <> "" Then
    cNum = foundC
    quoteNoHyphen = foundC
    If Left(UCase(foundC), 1) = "C" Then
        quoteNum = "C-" & Mid(foundC, 2)
    Else
        quoteNum = "C-" & ExtractDigits(foundC)
    End If
    LogStep "using existing C-number from job/email: " & cNum
Else
    quoteNum = GetNextQuoteNumber()
    If quoteNum = "" Then quoteNum = "C-00000"
    quoteNoHyphen = Replace(quoteNum, "-", "")
    cNum = quoteNoHyphen
    LogStep "no existing C-number found — assigned new quote: " & quoteNum
End If

' 5. Create the job folder in the current month folder and drop the
'    downloaded CAD/BOM files (from the email) into it.
Dim monthFolder, jobFolderName, jobFolderPath
monthFolder   = CurrentMonthFolder()
customerPrefix = CleanFolderToken(IIf(gCustomerPrefix <> "", gCustomerPrefix, DEFAULT_CUSTOMER_PREFIX))
customerName = IIf(gCustomerName <> "", gCustomerName, customerPrefix)
jobFolderName = BuildJobFolderName(customerPrefix, custJobNum, quoteNoHyphen)
jobFolderPath = CreateJobFolder(monthFolder, jobFolderName, gAttachDir)
If jobFolderPath = "" Then
    jobFolderPath = CreateJobFolder(LOCAL_WORKSPACE_ROOT, jobFolderName, gAttachDir)
    If jobFolderPath <> "" Then
        monthFolder = LOCAL_WORKSPACE_ROOT
        LogStep "network share unreachable — staged job files locally: " & jobFolderPath
    ElseIf gAttachDir <> "" And fso.FolderExists(gAttachDir) Then
        LogStep "job folder not created — macro will use AttachDir: " & gAttachDir
    End If
End If

' 6. Find the CAD file NOW so SolidWorks can open it BEFORE the macro runs.
gCadPath = FindBestCadInFolders(jobFolderPath, gAttachDir)
If gCadPath <> "" Then
    If IsGeneratedBaseCadPath(gCadPath) Then
        LogStep "WARNING: ignoring generated base CAD path: " & gCadPath
        gCadPath = ""
    End If
End If
If gCadPath <> "" Then
    LogStep "CAD to open first: " & gCadPath
Else
    LogStep "WARNING: no CAD file found yet in job/attach folders"
End If

' 7. Write the handoff file for Module6121 (includes CadPath so macro uses open model)
WriteHandoff cNum, quoteNum, custJobNum, similarTo, shipDate, monthFolder, jobFolderName, customerPrefix, customerName, gAttachDir, gCadPath

Dim proposalPath
proposalPath = ""
LogStep "Quote=" & quoteNum & "  Job=" & jobFolderName & "  CustJob=" & custJobNum & _
        "  Files=" & IIf(jobFolderPath <> "", "yes", "no") & "  CAD=" & IIf(gCadPath <> "", "yes", "no")

' 8. Open SolidWorks → open the part/assembly FIRST → then run Module6121.swp
Dim summary
summary = "Quote #: " & quoteNum & " | Job folder: " & jobFolderName & _
          " | Month folder: " & monthFolder & " | Customer Job#: " & IIf(custJobNum <> "", custJobNum, "(not found)") & _
          " | CAD: " & IIf(gCadPath <> "", gCadPath, "(none)") & _
          " | Job files: " & IIf(jobFolderPath <> "", jobFolderPath, "(share not reachable)")
LogStep summary

LogStep "launching SolidWorks, opening CAD, then Module6121.swp"
If LaunchSolidWorksOpenCadThenMacro() Then
    proposalPath = FillProposal(cNum, quoteNum, custJobNum)
    LogStep "Proposal filled after macro start: " & IIf(proposalPath <> "", proposalPath, "skipped")
Else
    LogStep "Proposal skipped because SolidWorks / CAD / macro did not start."
End If
LogStep "===== launcher done ====="

' ============================================================
' PROPOSAL FILL
' ============================================================
Function FillProposal(cNum, quoteNum, custJobNum)
    FillProposal = ""
    Dim templatePath
    templatePath = FindProposalTemplate()
    If templatePath = "" Then
        LogStep "Proposal template not found in Downloads; proposal skipped. Folder: " & DOWNLOADS_FOLDER
        Exit Function
    End If

    ' Copy template to a job-named file in Downloads (number once, not twice)
    Dim ext, destName, destPath
    ext      = fso.GetExtensionName(templatePath)
    destName = "Custom Quote #" & quoteNum & "." & ext
    destPath = LOCAL_WORKSPACE_ROOT & "\" & destName
    fso.CopyFile templatePath, destPath, True

    ' Fill header fields via Excel COM
    Dim xl, wb, ws
    On Error Resume Next
    Set xl = CreateObject("Excel.Application")
    If Err.Number <> 0 Then
        LogStep "Excel not found - proposal not filled."
        FillProposal = destPath
        Exit Function
    End If
    On Error GoTo 0
    xl.Visible       = False
    xl.DisplayAlerts = False
    xl.EnableEvents  = False

    Set wb = xl.Workbooks.Open(destPath)

    ' Calculation can only be set once a workbook is open
    On Error Resume Next
    xl.Calculation = -4135   ' xlCalculationManual
    On Error GoTo 0

    Set ws = Nothing
    On Error Resume Next
    Set ws = wb.Worksheets("Quote Request")
    On Error GoTo 0
    If IsNull(ws) Or IsEmpty(ws) Then Set ws = wb.Worksheets(1)

    ws.Range("C3").Value = IIf(gCustomerName <> "", gCustomerName, CMS_CUSTOMER_NAME)
    ws.Range("C4").Value = quoteNum
    If custJobNum <> "" Then ws.Range("C5").Value = custJobNum
    ws.Range("C6").Value = CMS_ATTENTION
    ws.Range("G7").Value = IIf(gCustomerPrefix <> "", gCustomerPrefix, CMS_SALES_REP)
    ws.Range("C9").Value = Now()
    ws.Range("E50").Value = "Payment Terms:  " & CMS_PAYMENT_TERMS

    On Error Resume Next
    wb.Worksheets(1).Calculate
    On Error GoTo 0
    wb.Save
    wb.Close True
    xl.Quit
    Set ws = Nothing: Set wb = Nothing: Set xl = Nothing

    ' Save numbered copy to the proposals folder on the network.
    EnsureFolderDeep QUOTE_PROPOSALS_FOLDER
    If fso.FolderExists(QUOTE_PROPOSALS_FOLDER) Then
        Dim netDest
        netDest = QUOTE_PROPOSALS_FOLDER & "\Custom Quote #" & quoteNum & "." & ext
        fso.CopyFile destPath, netDest, True
    End If

    FillProposal = destPath
End Function

' ============================================================
' QUOTE NUMBERING
' ============================================================
Function GetNextQuoteNumber()
    GetNextQuoteNumber = ""
    EnsureFolderDeep QUOTE_PROPOSALS_FOLDER

    Dim maxN, regFolder, regMax
    maxN = MaxQuoteNumberInFolder(QUOTE_PROPOSALS_FOLDER)

    ' Keep Ron's quotes in their own folder, but do not accidentally reuse a
    ' C-number that already exists in the regular proposal archive.
    regFolder = "\\Mycloudex2ultra\mexico\Downloads\Quote-Proposals-" & Year(Date)
    regMax = MaxQuoteNumberInFolder(regFolder)
    If regMax > maxN Then maxN = regMax

    If maxN > 0 Then GetNextQuoteNumber = "C-" & (maxN + 1)
End Function

Function MaxQuoteNumberInFolder(folderPath)
    MaxQuoteNumberInFolder = 0
    If Not fso.FolderExists(folderPath) Then Exit Function
    Dim folder, f, n
    Set folder = fso.GetFolder(folderPath)
    For Each f In folder.Files
        If Left(f.Name, 2) <> "~$" Then
            n = ExtractQuoteCNumber(f.Name)
            If n > MaxQuoteNumberInFolder Then MaxQuoteNumberInFolder = n
        End If
    Next
End Function

Function ExtractQuoteCNumber(nm)
    ExtractQuoteCNumber = 0
    Dim u, p, i, digits, ch
    u = UCase(nm)
    p = InStr(u, "C-")
    If p = 0 Then Exit Function
    i = p + 2: digits = ""
    Do While i <= Len(u)
        ch = Mid(u, i, 1)
        If ch >= "0" And ch <= "9" Then digits = digits & ch Else Exit Do
        i = i + 1
    Loop
    If digits <> "" Then ExtractQuoteCNumber = CLng(digits)
End Function

' ============================================================
' PROPOSAL TEMPLATE FINDER
' ============================================================
Function FindProposalTemplate()
    FindProposalTemplate = ""
    If Not fso.FolderExists(DOWNLOADS_FOLDER) Then Exit Function
    Dim folder, f, nm, ext
    Set folder = fso.GetFolder(DOWNLOADS_FOLDER)
    For Each f In folder.Files
        nm  = UCase(f.Name)
        ext = LCase(fso.GetExtensionName(f.Name))
        If (ext = "xls" Or ext = "xlsx" Or ext = "xlsm") And Left(f.Name, 2) <> "~$" Then
            If InStr(nm, "GRIND") = 0 Then
                If (InStr(nm, "CUSTOM") > 0 And InStr(nm, "QUOTE") > 0) Or InStr(nm, "PROPOSAL") > 0 Then
                    FindProposalTemplate = f.Path
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' ============================================================
' HANDOFF FILE  (Module6121 reads this at startup)
' ============================================================
Sub DeleteIfExists(ByVal p)
    On Error Resume Next
    If fso.FileExists(p) Then fso.DeleteFile p, True
    On Error GoTo 0
End Sub

Sub WriteHandoffAtomic(ByVal handoffPath, ByVal text)
    Dim tmpPath
    tmpPath = handoffPath & ".tmp"

    DeleteIfExists tmpPath

    Dim ts
    Set ts = fso.CreateTextFile(tmpPath, True)
    ts.Write text
    ts.Close

    DeleteIfExists handoffPath
    fso.MoveFile tmpPath, handoffPath
End Sub

Sub WriteHandoff(cNum, quoteNum, custJobNum, similarTo, shipDate, rootPath, jobFolder, customerPrefix, customerName, attachDir, cadPath)
    Dim body
    body = "CNum=" & cNum & vbCrLf & _
           "QuoteNum=" & quoteNum & vbCrLf & _
           "CustJob=" & custJobNum & vbCrLf & _
           "SimilarTo=" & similarTo & vbCrLf & _
           "ShipDate=" & shipDate & vbCrLf & _
           "RootPath=" & rootPath & vbCrLf & _
           "JobFolder=" & jobFolder & vbCrLf & _
           "CustomerPrefix=" & customerPrefix & vbCrLf & _
           "CustomerName=" & customerName & vbCrLf
    If attachDir <> "" Then body = body & "AttachDir=" & attachDir & vbCrLf
    If cadPath <> "" Then body = body & "CadPath=" & cadPath & vbCrLf
    WriteHandoffAtomic HANDOFF_FILE, body
End Sub

' Write BatchCount=N + Job1.* / Job2.* ... for sequential multi-quote runs.
Sub WriteBatchHandoff(ByVal jobs)
    ' jobs is a 1-based array of dictionaries OR a Collection of Scripting.Dictionary
    Dim i, n, body, d
    n = UBound(jobs)
    body = "BatchCount=" & n & vbCrLf & vbCrLf
    For i = 1 To n
        Set d = jobs(i)
        body = body & "Job" & i & ".CNum=" & d("CNum") & vbCrLf
        body = body & "Job" & i & ".QuoteNum=" & d("QuoteNum") & vbCrLf
        body = body & "Job" & i & ".CustJob=" & d("CustJob") & vbCrLf
        body = body & "Job" & i & ".SimilarTo=" & d("SimilarTo") & vbCrLf
        body = body & "Job" & i & ".ShipDate=" & d("ShipDate") & vbCrLf
        body = body & "Job" & i & ".RootPath=" & d("RootPath") & vbCrLf
        body = body & "Job" & i & ".JobFolder=" & d("JobFolder") & vbCrLf
        body = body & "Job" & i & ".CustomerPrefix=" & d("CustomerPrefix") & vbCrLf
        body = body & "Job" & i & ".CustomerName=" & d("CustomerName") & vbCrLf
        body = body & "Job" & i & ".AttachDir=" & d("AttachDir") & vbCrLf
        body = body & "Job" & i & ".CadPath=" & d("CadPath") & vbCrLf & vbCrLf
    Next
    WriteHandoffAtomic HANDOFF_FILE, body
End Sub

' Prefer C##### already present on the job (folder name, subject, attach path).
' BMS-851100029-C18603 → C18603. Returns "" if none found.
Function ResolveExistingCNumber(ByVal custJob, ByVal attachDir)
    ResolveExistingCNumber = ""
    Dim sources, i, hit
    sources = Array(gPreferredCNum, custJob, attachDir)
    ' Also scan cms_email.txt Subject / CustJob / AttachDir if present
    If fso.FileExists(EMAIL_OUTPUT_FILE) Then
        Dim ts, line, p, k, v
        Set ts = fso.OpenTextFile(EMAIL_OUTPUT_FILE, 1)
        Do Until ts.AtEndOfStream
            line = ts.ReadLine
            p = InStr(line, "=")
            If p > 0 Then
                k = UCase(Trim(Left(line, p - 1)))
                v = Trim(Mid(line, p + 1))
                If k = "SUBJECT" Or k = "CUSTJOB" Or k = "ATTACHDIR" Or k = "CNUM" Or k = "QUOTENUM" Then
                    hit = ExtractCNumberToken(v)
                    If hit <> "" Then
                        ts.Close
                        ResolveExistingCNumber = hit
                        Exit Function
                    End If
                End If
            End If
        Loop
        ts.Close
    End If
    For i = 0 To UBound(sources)
        hit = ExtractCNumberToken(CStr(sources(i)))
        If hit <> "" Then
            ResolveExistingCNumber = hit
            Exit Function
        End If
    Next
End Function

Function ExtractCNumberToken(s)
    ExtractCNumberToken = ""
    Dim u, p, i, ch, digits
    u = UCase(CStr(s))
    If u = "" Then Exit Function
    ' Prefer -C##### or _C##### or trailing C#####
    p = InStr(u, "-C")
    If p = 0 Then p = InStr(u, "_C")
    If p > 0 Then
        i = p + 2
        digits = ""
        Do While i <= Len(u)
            ch = Mid(u, i, 1)
            If ch >= "0" And ch <= "9" Then
                digits = digits & ch
            Else
                Exit Do
            End If
            i = i + 1
        Loop
        If Len(digits) >= 4 Then
            ExtractCNumberToken = "C" & digits
            Exit Function
        End If
    End If
    ' Standalone C##### token
    p = 1
    Do While p <= Len(u)
        If Mid(u, p, 1) = "C" And p < Len(u) Then
            If Mid(u, p + 1, 1) >= "0" And Mid(u, p + 1, 1) <= "9" Then
                If p = 1 Or Not ((Mid(u, p - 1, 1) >= "A" And Mid(u, p - 1, 1) <= "Z") Or (Mid(u, p - 1, 1) >= "0" And Mid(u, p - 1, 1) <= "9")) Then
                    i = p + 1
                    digits = ""
                    Do While i <= Len(u)
                        ch = Mid(u, i, 1)
                        If ch >= "0" And ch <= "9" Then
                            digits = digits & ch
                        Else
                            Exit Do
                        End If
                        i = i + 1
                    Loop
                    If Len(digits) >= 4 And Len(digits) <= 6 Then
                        ExtractCNumberToken = "C" & digits
                        Exit Function
                    End If
                End If
            End If
        End If
        p = p + 1
    Loop
End Function

Function IsGeneratedBaseCadPath(ByVal p)
    IsGeneratedBaseCadPath = False
    Dim u
    u = UCase(CStr(p))
    If u = "" Then Exit Function
    If InStr(u, "\BASE\") > 0 Then IsGeneratedBaseCadPath = True: Exit Function
    If InStr(u, "/BASE/") > 0 Then IsGeneratedBaseCadPath = True: Exit Function
End Function

' Rank CAD files: strongly prefer the assembly that matches this job's C-number.
' Example: 863700126-C18614.sldasm beats 863700102_RFQ_MB_ASM_....sldasm
' Never prefer previously exported \base\*.SLDASM outputs.
Function CadPriority(ext, fileName)
    Dim e, bonus, u
    e = LCase(ext)
    u = UCase(fileName)
    bonus = 0
    If cNum <> "" Then
        If InStr(u, UCase(cNum)) > 0 Then bonus = bonus + 500
    End If
    If quoteNoHyphen <> "" Then
        If InStr(u, UCase(quoteNoHyphen)) > 0 Then bonus = bonus + 400
    End If
    If jobFolderName <> "" Then
        If InStr(u, UCase(jobFolderName)) > 0 Then bonus = bonus + 200
    End If
    If InStr(u, "MOLDBASE") > 0 Or InStr(u, "MOLD_BASE") > 0 Then
        bonus = bonus + 30
    ElseIf InStr(u, "BASE") > 0 And InStr(u, "DATABASE") = 0 And InStr(u, "MOLDBASE") = 0 Then
        bonus = bonus + 10
    End If
    If InStr(u, "RFQ") > 0 And bonus < 400 Then bonus = bonus - 40
    Select Case e
        Case "sldasm": CadPriority = 100 + bonus
        Case "step", "stp": CadPriority = 80 + bonus
        Case "x_t", "x_b": CadPriority = 75 + bonus
        Case "igs", "iges": CadPriority = 70 + bonus
        Case "sldprt": CadPriority = 50 + bonus
        Case "prt": CadPriority = 45 + bonus
        Case Else: CadPriority = 0
    End Select
    If CadPriority < 0 Then CadPriority = 0
End Function

Function FindBestCadInFolder(folderPath)
    FindBestCadInFolder = ""
    If folderPath = "" Then Exit Function
    If Not fso.FolderExists(folderPath) Then Exit Function

    ' Never search generated CMS output folders.
    If UCase(fso.GetFileName(folderPath)) = "BASE" Then Exit Function
    If IsGeneratedBaseCadPath(folderPath) Then Exit Function

    Dim bestPath, bestScore, f, sub1, score, hit
    bestPath = "": bestScore = 0
    On Error Resume Next
    For Each f In fso.GetFolder(folderPath).Files
        If Not IsGeneratedBaseCadPath(f.Path) Then
            score = CadPriority(fso.GetExtensionName(f.Name), f.Name)
            If score > bestScore Then
                bestScore = score
                bestPath = f.Path
            End If
        End If
    Next
    For Each sub1 In fso.GetFolder(folderPath).SubFolders
        If UCase(Left(sub1.Name, 1)) <> "_" Then
            If UCase(sub1.Name) <> "BASE" Then
                hit = FindBestCadInFolder(sub1.Path)
                If hit <> "" Then
                    If Not IsGeneratedBaseCadPath(hit) Then
                        score = CadPriority(fso.GetExtensionName(hit), fso.GetFileName(hit))
                        If score > bestScore Then
                            bestScore = score
                            bestPath = hit
                        End If
                    End If
                End If
            End If
        End If
    Next
    On Error GoTo 0
    FindBestCadInFolder = bestPath
End Function

Function FindBestCadInFolders(jobFolder, attachDir)
    Dim a, b, sa, sb
    a = FindBestCadInFolder(jobFolder)
    b = FindBestCadInFolder(attachDir)
    If a <> "" And IsGeneratedBaseCadPath(a) Then a = ""
    If b <> "" And IsGeneratedBaseCadPath(b) Then b = ""
    If a = "" Then FindBestCadInFolders = b: Exit Function
    If b = "" Then FindBestCadInFolders = a: Exit Function
    sa = CadPriority(fso.GetExtensionName(a), fso.GetFileName(a))
    sb = CadPriority(fso.GetExtensionName(b), fso.GetFileName(b))
    If sb > sa Then FindBestCadInFolders = b Else FindBestCadInFolders = a
End Function

' Open SolidWorks 2023 → open the CAD part/assembly → THEN run Module6121.swp.
' This matches how you work manually and avoids the empty welcome-screen hang.
Function LaunchSolidWorksOpenCadThenMacro()
    Dim shell, sw, tries, macroPath, macroPaths(), mpIdx, pathCount
    Dim modNames, mi, okRun, ran, procNames, pi, macroErr
    Dim errs, warns, importErrors, docType, ext, opened
    Set shell = CreateObject("WScript.Shell")
    LaunchSolidWorksOpenCadThenMacro = False

    pathCount = 0
    If fso.FileExists(LOCAL_WORKSPACE_ROOT & "\Module6121.swp") Then
        ReDim macroPaths(0)
        macroPaths(0) = LOCAL_WORKSPACE_ROOT & "\Module6121.swp"
        pathCount = 1
        LogStep "using compiled macro: Module6121.swp"
    ElseIf fso.FileExists(SW_MACRO) Then
        ReDim macroPaths(0)
        macroPaths(0) = SW_MACRO
        pathCount = 1
    Else
        LogStep "ERROR: Module6121.swp not found at " & LOCAL_WORKSPACE_ROOT
        Exit Function
    End If

    LogStep "sw exe: " & SW_EXE & "  progid: " & SW_PROGID

    On Error Resume Next
    Set sw = GetObject(, SW_PROGID)
    If sw Is Nothing Then Set sw = CreateObject(SW_PROGID)
    On Error GoTo 0

    If sw Is Nothing Then
        If Not fso.FileExists(SW_EXE) Then
            LogStep "SolidWorks 2023 not found at: " & SW_EXE
            Exit Function
        End If
        LogStep "starting SolidWorks 2023..."
        shell.Run """" & SW_EXE & """", 1, False
        tries = 0
        Do
            WScript.Sleep 3000
            On Error Resume Next
            Set sw = GetObject(, SW_PROGID)
            On Error GoTo 0
            tries = tries + 1
        Loop Until (Not sw Is Nothing) Or tries > 20
        If sw Is Nothing Then
            LogStep "SolidWorks 2023 did not start in time"
            Exit Function
        End If
        LogStep "connected to SolidWorks 2023"
    Else
        LogStep "using existing SolidWorks 2023 session"
    End If

    On Error Resume Next
    sw.Visible = True
    On Error GoTo 0
    WScript.Sleep 2000

    tries = 0
    Do While tries < 45
        On Error Resume Next
        If Not sw.CommandInProgress Then Exit Do
        On Error GoTo 0
        WScript.Sleep 1000
        tries = tries + 1
    Loop
    On Error Resume Next
    sw.CommandInProgress = False
    On Error GoTo 0

    ' ---- OPEN THE CAD FIRST ----
    opened = False
    If gCadPath <> "" And IsGeneratedBaseCadPath(gCadPath) Then
        LogStep "WARNING: refusing to open generated \base\ assembly before macro: " & gCadPath
        gCadPath = ""
    End If
    If gCadPath <> "" And fso.FileExists(gCadPath) Then
        ext = LCase(fso.GetExtensionName(gCadPath))
        errs = 0: warns = 0: importErrors = 0
        LogStep "opening CAD before macro: " & gCadPath
        On Error Resume Next
        If ext = "sldasm" Then
            sw.OpenDoc6 gCadPath, 2, 1, "", errs, warns   ' swDocASSEMBLY=2, Silent=1
            If Err.Number = 0 Then opened = True
        ElseIf ext = "sldprt" Then
            sw.OpenDoc6 gCadPath, 1, 1, "", errs, warns   ' swDocPART=1
            If Err.Number = 0 Then opened = True
        Else
            ' STEP / X_T / IGES — LoadFile4
            sw.LoadFile4 gCadPath, "r", Nothing, importErrors
            If Err.Number = 0 Then opened = True
            If Not opened Then
                Err.Clear
                sw.LoadFile4 gCadPath, "", Nothing, importErrors
                If Err.Number = 0 Then opened = True
            End If
            If Not opened Then
                Err.Clear
                sw.OpenDoc6 gCadPath, 2, 1, "", errs, warns
                If Err.Number = 0 Then opened = True
            End If
        End If
        Err.Clear
        On Error GoTo 0
        If opened Then
            LogStep "CAD opened successfully — waiting for model to settle"
            WScript.Sleep 3000
            tries = 0
            Do While tries < 60
                On Error Resume Next
                If Not sw.CommandInProgress Then Exit Do
                On Error GoTo 0
                WScript.Sleep 1000
                tries = tries + 1
            Loop
            On Error Resume Next
            sw.CommandInProgress = False
            sw.Visible = True
            On Error GoTo 0
        Else
            LogStep "WARNING: OpenDoc/LoadFile failed for " & gCadPath & " — macro will try to open it"
        End If
    Else
        LogStep "WARNING: no CAD path to open first — macro will search job folder"
    End If

    ' ---- THEN RUN THE .SWP MACRO WITH STARTED ACKNOWLEDGEMENT ----
    ' Prefer main() — it routes to RunFromLauncher when cms_handoff.txt exists.
    ' Wait for cms_macro_started.txt so we do not assume a failed launch succeeded.
    If Not fso.FileExists(HANDOFF_FILE) Then
        LogStep "ERROR: Quote cancelled before launch: handoff file was not created: " & HANDOFF_FILE
        Exit Function
    End If

    WaitSeconds 3

    ran = False
    For mpIdx = 0 To pathCount - 1
        macroPath = macroPaths(mpIdx)
        LogStep "running macro with retry: " & macroPath
        If RunMacroWithRetry(sw, macroPath, "Module6121", "main", 90) Then
            ran = True
            Exit For
        End If
        ' Fallback: call RunFromLauncher directly if main routing fails.
        If RunMacroWithRetry(sw, macroPath, "Module6121", "RunFromLauncher", 60) Then
            ran = True
            Exit For
        End If
    Next

    If Not ran Then
        LogStep "ERROR: SolidWorks opened, but Module6121 did not acknowledge launch (no cms_macro_started.txt)."
    End If
    LaunchSolidWorksOpenCadThenMacro = ran
End Function

Function RunMacroWithRetry(ByVal swApp, ByVal macroPath, ByVal moduleName, ByVal procName, ByVal timeoutSeconds)
    RunMacroWithRetry = False

    Dim startTime, attempt, runOk, runErr, waitStart
    startTime = Timer
    attempt = 0

    Do
        attempt = attempt + 1
        DeleteIfExists MACRO_STARTED_FILE
        DeleteIfExists MACRO_ERROR_FILE

        runOk = False
        runErr = 0

        On Error Resume Next
        Err.Clear
        swApp.CommandInProgress = True
        runOk = swApp.RunMacro2(macroPath, moduleName, procName, 0, runErr)
        If Err.Number <> 0 Or runOk = False Then
            Err.Clear
            runOk = swApp.RunMacro(macroPath, moduleName, procName)
        End If
        swApp.CommandInProgress = False
        On Error GoTo 0

        LogStep "RunMacro attempt " & attempt & " module=" & moduleName & " proc=" & procName & " ok=" & CStr(runOk) & " macroErr=" & runErr

        waitStart = Timer
        Do
            If fso.FileExists(MACRO_STARTED_FILE) Then
                LogStep "macro acknowledged STARTED via " & MACRO_STARTED_FILE
                RunMacroWithRetry = True
                Exit Function
            End If
            If fso.FileExists(MACRO_ERROR_FILE) Then
                LogStep "macro wrote ERROR file quickly: " & MACRO_ERROR_FILE
                RunMacroWithRetry = True
                Exit Function
            End If
            WaitSeconds 1
            If Timer < waitStart Then Exit Do
            If Timer - waitStart >= 10 Then Exit Do
        Loop

        WaitSeconds 2

        If Timer < startTime Then Exit Do
        If Timer - startTime >= timeoutSeconds Then Exit Do
    Loop
End Function

Sub WaitSeconds(ByVal sec)
    Dim t
    t = Timer
    Do
        WScript.Sleep 250
        If Timer < t Then Exit Do
        If Timer - t >= sec Then Exit Do
    Loop
End Sub

' Legacy name kept for any external callers — routes to open-CAD-first path.
Function LaunchSolidWorksAndMacro()
    LaunchSolidWorksAndMacro = LaunchSolidWorksOpenCadThenMacro()
End Function
' Show the Python job picker: a select-all list of unopened quote emails.
Sub LaunchJobPicker()
    Dim scriptPath, shell, cmd
    scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\cms_gmail_search.py"
    If Not fso.FileExists(scriptPath) Then scriptPath = GMAIL_SCRIPT
    If Not fso.FileExists(scriptPath) Then
        LogStep "cms_gmail_search.py was not found next to this launcher."
        Exit Sub
    End If
    Set shell = CreateObject("WScript.Shell")
    cmd = """" & PYTHON_EXE & """ """ & scriptPath & """ --pick"
    shell.Run cmd, 1, False   ' 1 = show the picker window; don't wait
End Sub

' ============================================================
' GMAIL SEARCH  (runs the Python script, reads its output file)
' ============================================================
' Refresh purchased-component prices (DME lookup) into the price-list CSV before
' the macro reads it. Non-fatal: if Python/Playwright isn't set up, just skip.
Sub RunPriceLookup()
    On Error Resume Next
    Dim scriptPath, shell, cmd
    scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\cms_price_lookup.py"
    If Not fso.FileExists(scriptPath) Then
        LogStep "price lookup script not found next to launcher - skipping"
        Exit Sub
    End If
    Set shell = CreateObject("WScript.Shell")
    cmd = """" & PYTHON_EXE & """ """ & scriptPath & """ --all"
    shell.Run cmd, 0, True   ' 0 = hidden, True = wait so prices are ready for the macro
    LogStep "price lookup finished"
    On Error GoTo 0
End Sub

Function RunGmailSearch(ByRef custJob, ByRef similar, ByRef ship)
    RunGmailSearch = False
    On Error Resume Next

    If Not gUseExistingEmail Then
        ' Find the Python script next to this .vbs (falls back to the configured path)
        Dim scriptPath
        scriptPath = fso.GetParentFolderName(WScript.ScriptFullName) & "\cms_gmail_search.py"
        If Not fso.FileExists(scriptPath) Then scriptPath = GMAIL_SCRIPT
        If Not fso.FileExists(scriptPath) Then
            LogStep "gmail search script not found - continuing without email metadata"
            Exit Function
        End If

        ' Clear any stale result, then search Gmail for the newest matching email
        If fso.FileExists(EMAIL_OUTPUT_FILE) Then fso.DeleteFile EMAIL_OUTPUT_FILE, True
        Dim shell, cmd
        Set shell = CreateObject("WScript.Shell")
        cmd = """" & PYTHON_EXE & """ """ & scriptPath & """ --from-launcher"
        shell.Run cmd, 0, True   ' 0 = hidden, True = wait for it to finish
    End If
    On Error GoTo 0

    If Not fso.FileExists(EMAIL_OUTPUT_FILE) Then Exit Function

    Dim ts, line, p, k, v, found, errMsg
    found = False : errMsg = ""
    Set ts = fso.OpenTextFile(EMAIL_OUTPUT_FILE, 1)
    Do Until ts.AtEndOfStream
        line = ts.ReadLine
        p = InStr(line, "=")
        If p > 0 Then
            k = Trim(Left(line, p - 1))
            v = Trim(Mid(line, p + 1))
            Select Case UCase(k)
                Case "FOUND":     If v = "1" Then found = True
                Case "CUSTJOB":   custJob = v
                Case "CNUM":      gPreferredCNum = v
                Case "QUOTENUM":  If gPreferredCNum = "" Then gPreferredCNum = Replace(v, "-", "")
                Case "SIMILARTO": similar = v
                Case "SHIPDATE":  ship = v
                Case "ATTACHDIR": gAttachDir = v
                Case "CUSTOMERPREFIX": gCustomerPrefix = v
                Case "CUSTOMERNAME": gCustomerName = v
                Case "ERROR":     errMsg = v
            End Select
        End If
    Loop
    ts.Close

    If errMsg <> "" Then
        LogStep "Gmail search reported an error: " & errMsg & ". Continuing without paste prompt."
        Exit Function
    End If

    RunGmailSearch = found
End Function


Function ExtractNumberAfter(body, token)
    ExtractNumberAfter = ""
    Dim u, p, i, ch, started, res
    u = UCase(body): res = "": started = False
    p = InStr(u, UCase(token))
    If p = 0 Then Exit Function
    i = p + Len(token)
    Do While i <= Len(u)
        ch = Mid(u, i, 1)
        If ch >= "0" And ch <= "9" Then
            res = res & ch: started = True
        ElseIf started Then
            Exit Do
        End If
        i = i + 1
    Loop
    ExtractNumberAfter = res
End Function

Function ExtractTextAfter(body, token)
    ExtractTextAfter = ""
    Dim u, p, e, s
    u = UCase(body)
    p = InStr(u, UCase(token))
    If p = 0 Then Exit Function
    p = p + Len(token)
    e = InStr(p, body, vbLf)
    If e = 0 Then e = Len(body) + 1
    s = Mid(body, p, e - p)
    s = Replace(s, vbCr, "")
    s = Trim(s)
    Do While Left(s, 1) = ":" Or Left(s, 1) = "-"
        s = Trim(Mid(s, 2))
    Loop
    ExtractTextAfter = s
End Function

' ============================================================
' STRING HELPERS
' ============================================================
Function NormalizeJobNum(s)
    s = UCase(Trim(s))
    If s = "" Then NormalizeJobNum = "" : Exit Function
    If Left(s, 1) = "C" Then s = Mid(s, 2)
    s = Replace(s, "-", "")
    If s = "" Then NormalizeJobNum = "" Else NormalizeJobNum = "C" & s
End Function

Function ExtractDigits(s)
    Dim i, ch, res
    res = ""
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If ch >= "0" And ch <= "9" Then res = res & ch
    Next
    ExtractDigits = res
End Function

Function IIf(cond, a, b)
    If cond Then IIf = a Else IIf = b
End Function

' Append a line to the Downloads log so the whole run is recorded.
Sub LogStep(msg)
    On Error Resume Next
    Dim f
    Set f = fso.OpenTextFile(DOWNLOADS_FOLDER & "\CMS_Quote_Log.txt", 8, True)  ' 8 = append, create
    f.WriteLine "[" & Now & "] launcher: " & msg
    f.Close
End Sub

' ============================================================
' CURRENT MONTH FOLDER + JOB ZIP
' Month folder pattern:  \\...\RON'S QUOTES\000000006.June 2026
' ============================================================
Sub EnsureFolderDeep(folderPath)
    On Error Resume Next
    If folderPath = "" Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub
    Dim parent
    parent = fso.GetParentFolderName(folderPath)
    If parent <> "" Then
        If Not fso.FolderExists(parent) Then EnsureFolderDeep parent
    End If
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
End Sub
Function CurrentMonthFolder()
    CurrentMonthFolder = JOB_ROOT_BASE & "\" & MonthFolderName(Now)
End Function

Function MonthFolderName(d)
    Dim m, y
    m = Month(d)
    y = Year(d)
    MonthFolderName = Right("00000000" & m, 9) & "." & MonthName(m) & " " & y
End Function

' Create  <monthFolder>\<jobName>\  and copy the downloaded email files into it.
' If no files were downloaded, drop an empty placeholder zip so the folder isn't bare.
' Returns the job folder path (or "" if the share is not reachable).
Function CleanFolderToken(s)
    Dim i, ch, out, lastDash
    s = Trim(CStr(s))
    out = "" : lastDash = False
    For i = 1 To Len(s)
        ch = Mid(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
            lastDash = False
        ElseIf ch = " " Or ch = "-" Or ch = "_" Or ch = "." Then
            If Not lastDash And out <> "" Then out = out & "-"
            lastDash = True
        End If
    Next
    Do While Right(out, 1) = "-"
        out = Left(out, Len(out) - 1)
    Loop
    CleanFolderToken = out
End Function

Function BuildJobFolderName(prefix, custJob, quoteNoHyphen)
    Dim p, j, core
    p = CleanFolderToken(prefix)
    j = CleanFolderToken(custJob)
    If j <> "" Then
        If p <> "" And UCase(Left(j, Len(p) + 1)) <> UCase(p & "-") And UCase(j) <> UCase(p) Then
            core = p & "-" & j
        Else
            core = j
        End If
    ElseIf p <> "" Then
        core = p
    Else
        core = "Quote"
    End If
    BuildJobFolderName = core & "-" & quoteNoHyphen
End Function
Function CreateJobFolder(monthFolder, jobName, attachDir)
    CreateJobFolder = ""
    On Error Resume Next
    EnsureFolderDeep monthFolder
    If Not fso.FolderExists(monthFolder) Then Exit Function   ' share not reachable
    Dim jp
    jp = monthFolder & "\" & jobName
    If Not fso.FolderExists(jp) Then fso.CreateFolder jp
    If Not fso.FolderExists(jp) Then Exit Function
    Dim copied
    copied = 0
    If attachDir <> "" Then
        If fso.FolderExists(attachDir) Then copied = CopyDirContents(attachDir, jp)
    End If
    LogStep "job folder " & jp & " - copied " & copied & " file(s) from " & IIf(attachDir <> "", attachDir, "(none)")
    CreateJobFolder = jp
End Function

' Recursively copy the contents of src into dst. Returns the number of files copied.
Function CopyDirContents(src, dst)
    Dim n, f, sub1, d2
    n = 0
    On Error Resume Next
    For Each f In fso.GetFolder(src).Files
        fso.CopyFile f.Path, dst & "\" & f.Name, True
        If Err.Number = 0 Then n = n + 1
        Err.Clear
    Next
    For Each sub1 In fso.GetFolder(src).SubFolders
        d2 = dst & "\" & sub1.Name
        If Not fso.FolderExists(d2) Then fso.CreateFolder d2
        n = n + CopyDirContents(sub1.Path, d2)
    Next
    CopyDirContents = n
End Function

' Write a valid (empty) zip = the 22-byte End-Of-Central-Directory record.
Sub WriteEmptyZip(path)
    On Error Resume Next
    Dim s
    Set s = CreateObject("ADODB.Stream")
    s.Type = 2                       ' text mode so we can write exact bytes
    s.Charset = "Windows-1252"       ' single-byte, no BOM
    s.Open
    s.WriteText Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, Chr(0))   ' PK\05\06 + 18 nulls
    s.SaveToFile path, 2             ' 2 = overwrite
    s.Close
End Sub

Function Q(ByVal s)
    Q = Chr(34) & s & Chr(34)
End Function
