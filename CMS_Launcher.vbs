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
' SolidWorks 2023 is the "(3)" install; the plain path opens 2025.
Const SW_EXE    = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
Const SW_PROGID = "SldWorks.Application.31"   ' 31 = SolidWorks 2023 (32=2024, 33=2025)
Const SW_MACRO = "C:\CMS_Local_Workspace\Module6121.swb"   ' source macro first; .swp fallback below

' Handoff file written for Module6121 to read
Const HANDOFF_FILE = "C:\CMS_Local_Workspace\cms_handoff.txt"

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
Dim gAttachDir
gAttachDir = ""

' Make sure the local workspace exists (handoff + email files live here)
If Not fso.FolderExists(LOCAL_WORKSPACE_ROOT) Then fso.CreateFolder LOCAL_WORKSPACE_ROOT
LogStep "===== launcher started ====="

' 1. No typing needed - the job number comes from the email and the
'    quote number is assigned automatically from the proposals folder.
Dim cNum
cNum = ""

' /usemail : the email picker already wrote cms_email.txt for a chosen message,
'            so use that instead of searching Gmail for the newest one.
Dim gUseExistingEmail, ai
gUseExistingEmail = False
For ai = 0 To WScript.Arguments.Count - 1
    If LCase(WScript.Arguments(ai)) = "/usemail" Then gUseExistingEmail = True
Next

' First screen: show the job picker (select-all list of unopened emails) and quit.
' The picker then runs this launcher again with /usemail for each chosen job.
If Not gUseExistingEmail Then
    LaunchJobPicker
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

' 3. Assign the next quote number from the proposals folder
Dim quoteNum
quoteNum = GetNextQuoteNumber()
If quoteNum = "" Then
    If cNum <> "" Then quoteNum = "C-" & ExtractDigits(cNum) Else quoteNum = "C-00000"
End If

' 4. If no C-number was typed, the job IS the assigned quote number
Dim quoteNoHyphen
quoteNoHyphen = Replace(quoteNum, "-", "")
If cNum = "" Then cNum = quoteNoHyphen

' 5. Create the job folder in the current month folder and drop the
'    downloaded CAD/BOM files (from the email) into it.
Dim monthFolder, jobFolderName, jobFolderPath
monthFolder   = CurrentMonthFolder()
customerPrefix = CleanFolderToken(IIf(gCustomerPrefix <> "", gCustomerPrefix, DEFAULT_CUSTOMER_PREFIX))
customerName = IIf(gCustomerName <> "", gCustomerName, customerPrefix)
jobFolderName = BuildJobFolderName(customerPrefix, custJobNum, quoteNoHyphen)
jobFolderPath = CreateJobFolder(monthFolder, jobFolderName, gAttachDir)

' 6. Write the handoff file for Module6121 (includes the month folder + job folder)
WriteHandoff cNum, quoteNum, custJobNum, similarTo, shipDate, monthFolder, jobFolderName, customerPrefix, customerName

Dim proposalPath
proposalPath = ""
LogStep "Quote=" & quoteNum & "  Job=" & jobFolderName & "  CustJob=" & custJobNum & _
        "  Files=" & IIf(jobFolderPath <> "", "yes", "no") & "  Proposal=pending macro start"

' 7. Summary, then launch SolidWorks + macro - all in one unattended run
Dim summary
summary = "Quote #: " & quoteNum & " | Job folder: " & jobFolderName & _
          " | Month folder: " & monthFolder & " | Customer Job#: " & IIf(custJobNum <> "", custJobNum, "(not found)") & _
          " | Similar to: " & IIf(similarTo  <> "", similarTo,  "(not found)") & _
          " | Ship date: " & IIf(shipDate   <> "", shipDate,   "(not found)") & _
          " | Job files: " & IIf(jobFolderPath <> "", jobFolderPath, "(share not reachable)") & _
          " | Proposal: pending macro start"
LogStep summary

LogStep "updating purchased-component prices (DME lookup)"
RunPriceLookup

LogStep "launching SolidWorks + Module6121"
If LaunchSolidWorksAndMacro() Then
    proposalPath = FillProposal(cNum, quoteNum, custJobNum)
    LogStep "Proposal filled after macro start: " & IIf(proposalPath <> "", proposalPath, "skipped")
Else
    LogStep "Proposal skipped because SolidWorks macro did not start."
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
Sub WriteHandoff(cNum, quoteNum, custJobNum, similarTo, shipDate, rootPath, jobFolder, customerPrefix, customerName)
    Dim f
    Set f = fso.CreateTextFile(HANDOFF_FILE, True)
    f.WriteLine "CNum="      & cNum
    f.WriteLine "QuoteNum="  & quoteNum
    f.WriteLine "CustJob="   & custJobNum
    f.WriteLine "SimilarTo=" & similarTo
    f.WriteLine "ShipDate="  & shipDate
    f.WriteLine "RootPath="  & rootPath
    f.WriteLine "JobFolder=" & jobFolder
    f.WriteLine "CustomerPrefix=" & customerPrefix
    f.WriteLine "CustomerName=" & customerName
    f.Close
End Sub

' ============================================================
' LAUNCH SOLIDWORKS + MACRO
' ============================================================
Function LaunchSolidWorksAndMacro()
    Dim shell, sw, tries, macroPath, macroFolder
    Set shell = CreateObject("WScript.Shell")

    macroFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    macroPath = SW_MACRO
    If Not fso.FileExists(macroPath) Then macroPath = macroFolder & "\Module6121.swb"
    If Not fso.FileExists(macroPath) Then macroPath = "C:\CMS_Local_Workspace\Module6121.swp"
    If Not fso.FileExists(macroPath) Then macroPath = macroFolder & "\Module6121.swp"

    On Error Resume Next
    Set sw = GetObject(, SW_PROGID)
    If sw Is Nothing Then Set sw = CreateObject(SW_PROGID)
    On Error GoTo 0

    If IsNull(sw) Or IsEmpty(sw) Or (sw Is Nothing) Then
        If Not fso.FileExists(SW_EXE) Then
            LogStep "SolidWorks 2023 not found at: " & SW_EXE & ". Update SW_EXE in CMS_Launcher.vbs."
            LaunchSolidWorksAndMacro = False
            Exit Function
        End If
        shell.Run """" & SW_EXE & """", 1, False
        tries = 0
        Do
            WScript.Sleep 3000
            On Error Resume Next
            Set sw = GetObject(, SW_PROGID)
            If sw Is Nothing Then Set sw = GetObject(, "SldWorks.Application")
            On Error GoTo 0
            tries = tries + 1
        Loop Until (Not (sw Is Nothing)) Or tries > 20
        If sw Is Nothing Then
            LogStep "SolidWorks 2023 did not start in time; macro not run."
            LaunchSolidWorksAndMacro = False
            Exit Function
        End If
    End If

    On Error Resume Next
    sw.Visible = True
    On Error GoTo 0
    WScript.Sleep 1500

    If Not fso.FileExists(macroPath) Then
        LogStep "Module6121.swb/.swp not found next to launcher or at: " & SW_MACRO & ". Macro not run."
        LaunchSolidWorksAndMacro = False
        Exit Function
    End If

    Dim modNames, mi, okRun, ran
    modNames = Array("Module6121", "Module61211", "Module612111", "Module1", "main", "Module2", "Module3")
    ran = False
    For mi = 0 To UBound(modNames)
        On Error Resume Next
        sw.CommandInProgress = True
        okRun = sw.RunMacro(macroPath, modNames(mi), "RunFromLauncher")
        sw.CommandInProgress = False
        If Err.Number = 0 And okRun = True Then
            On Error GoTo 0
            ran = True
            LogStep "macro started via module '" & modNames(mi) & "' from " & macroPath
            Exit For
        End If
        LogStep "RunMacro failed module '" & modNames(mi) & "' err=" & Err.Number & " " & Err.Description
        Err.Clear
        On Error GoTo 0
    Next

    If Not ran Then
        LogStep "RunMacro could not start the macro for any known module name from: " & macroPath
        On Error Resume Next
        sw.Visible = True
        On Error GoTo 0
    End If
    LaunchSolidWorksAndMacro = ran
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
