' ============================================================
' RunModule6121.vbs
' Thin SolidWorks macro launcher (VBScript COM — more reliable than PS1 ByRef).
' Usage:
'   wscript RunModule6121.vbs
'   wscript RunModule6121.vbs "C:\CMS_Local_Workspace\Module6121.swp"
' ============================================================
Option Explicit

Const LOCAL_WORKSPACE = "C:\CMS_Local_Workspace"
Const SW_EXE = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
Const SW_PROGID = "SldWorks.Application.31"
Const LOG_FILE = "C:\CMS_Local_Workspace\CMS_Quote_Log.txt"
Const STATUS_FILE = "C:\CMS_Local_Workspace\cms_launcher_status.txt"
Const STARTED_FILE = "C:\CMS_Local_Workspace\cms_macro_started.txt"
Const ERROR_FILE = "C:\CMS_Local_Workspace\cms_macro_error.txt"
Const DONE_FILE = "C:\CMS_Local_Workspace\cms_macro_done.txt"
Const RUNNER_TAG = "macro-runner-v3"

Dim fso, shell, macroPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If WScript.Arguments.Count >= 1 Then
    macroPath = WScript.Arguments(0)
Else
    macroPath = LOCAL_WORKSPACE & "\Module6121.swp"
End If

LogLine RUNNER_TAG & ": starting; macro=" & macroPath

If Not fso.FileExists(macroPath) Then
    LogLine RUNNER_TAG & ": ERROR macro not found: " & macroPath
    WScript.Quit 2
End If

Dim fi
Set fi = fso.GetFile(macroPath)
LogLine RUNNER_TAG & ": macro size=" & fi.Size & " modified=" & fi.DateLastModified
If fi.Size < 1000 Then
    LogLine RUNNER_TAG & ": WARNING .swp looks too small — recompile Module6121.bas -> .swp"
End If

DeleteIfExists STARTED_FILE
DeleteIfExists ERROR_FILE
DeleteIfExists DONE_FILE

' Always start from a clean SolidWorks process.
LogLine RUNNER_TAG & ": force-closing any running SolidWorks..."
On Error Resume Next
shell.Run "taskkill /F /IM SLDWORKS.exe /T", 0, True
shell.Run "taskkill /F /IM sldworks.exe /T", 0, True
shell.Run "taskkill /F /IM SLDWORKS_FCE.exe /T", 0, True
On Error GoTo 0
WScript.Sleep 5000
LogLine RUNNER_TAG & ": SolidWorks force-close done — starting fresh session"

Dim sw
Set sw = Nothing
On Error Resume Next
Set sw = CreateObject(SW_PROGID)
On Error GoTo 0

If sw Is Nothing Then
    If Not fso.FileExists(SW_EXE) Then
        LogLine RUNNER_TAG & ": ERROR SolidWorks exe missing: " & SW_EXE
        WScript.Quit 3
    End If
    LogLine RUNNER_TAG & ": starting SolidWorks..."
    shell.Run """" & SW_EXE & """", 1, False
    Dim tries
    tries = 0
    Do
        WScript.Sleep 2000
        On Error Resume Next
        Set sw = GetObject(, SW_PROGID)
        On Error GoTo 0
        tries = tries + 1
    Loop Until (Not sw Is Nothing) Or tries > 30
End If

If sw Is Nothing Then
    LogLine RUNNER_TAG & ": ERROR could not connect to SolidWorks"
    WScript.Quit 3
End If

On Error Resume Next
sw.Visible = True
sw.UserControl = True
sw.CommandInProgress = False
On Error GoTo 0
WScript.Sleep 1500

Dim pairs, methods, mi, entry, parts, pairList
pairList = ""

On Error Resume Next
methods = sw.GetMacroMethods(macroPath, 1)
If IsEmpty(methods) Or IsNull(methods) Then methods = sw.GetMacroMethods(macroPath, 0)
On Error GoTo 0

If IsArray(methods) Then
    For mi = 0 To UBound(methods)
        entry = CStr(methods(mi))
        LogLine RUNNER_TAG & ": GetMacroMethods entry: " & entry
        parts = Split(entry, ".")
        If UBound(parts) >= 1 Then
            If pairList <> "" Then pairList = pairList & "|"
            pairList = pairList & parts(0) & Chr(1) & parts(1)
        End If
    Next
Else
    LogLine RUNNER_TAG & ": GetMacroMethods returned no entry points — recompile .swp or enable macros"
End If

If pairList <> "" Then pairList = pairList & "|"
pairList = pairList & "Module61211" & Chr(1) & "main" & "|" & _
           "Module61211" & Chr(1) & "RunFromLauncher" & "|" & _
           "Module6121" & Chr(1) & "main" & "|" & _
           "Module6121" & Chr(1) & "RunFromLauncher" & "|" & _
           "Module1" & Chr(1) & "main"

Dim arr, pi, moduleName, procName, runOk, runErr, vbaErr, waited
arr = Split(pairList, "|")

For pi = 0 To UBound(arr)
    parts = Split(arr(pi), Chr(1))
    If UBound(parts) >= 1 Then
    moduleName = parts(0)
    procName = parts(1)

    DeleteIfExists STARTED_FILE
    DeleteIfExists ERROR_FILE

    On Error Resume Next
    sw.CommandInProgress = False
    Err.Clear
    runOk = False
    runErr = CLng(0)
    ' Prefer RunMacro (no ByRef)
    runOk = sw.RunMacro(macroPath, moduleName, procName)
    vbaErr = Err.Number
    If (runOk = False) Or (vbaErr <> 0) Then
        Err.Clear
        runErr = CLng(0)
        runOk = sw.RunMacro2(macroPath, moduleName, procName, 0, runErr)
        vbaErr = Err.Number
    End If
    If (runOk = False) Or (vbaErr <> 0) Then
        Err.Clear
        runErr = CLng(0)
        runOk = sw.RunMacro2(macroPath, moduleName, procName, 1, runErr)
        vbaErr = Err.Number
    End If
    On Error GoTo 0

    LogLine RUNNER_TAG & ": attempt module='" & moduleName & "' proc='" & procName & _
            "' ok=" & CStr(runOk) & " err=" & runErr & " vbaErr=" & vbaErr

    If WaitAck(12) Then
        LogLine RUNNER_TAG & ": SUCCESS via COM module=" & moduleName & " proc=" & procName
        WScript.Quit 0
    End If
    End If
Next

' Fallback: command line /m
LogLine RUNNER_TAG & ": COM RunMacro failed — trying SLDWORKS.EXE /m"
DeleteIfExists STARTED_FILE
DeleteIfExists ERROR_FILE
On Error Resume Next
shell.Run """" & SW_EXE & """ /m """ & macroPath & """", 1, False
On Error GoTo 0
If WaitAck(50) Then
    LogLine RUNNER_TAG & ": SUCCESS via /m fallback"
    WScript.Quit 0
End If

LogLine RUNNER_TAG & ": FAILED — recompile Module6121.bas to Module6121.swp (see webapp\COMPILE_MODULE6121.bat)"
LogLine RUNNER_TAG & ": Also check Tools > Options > System Options > Macro (enable / trusted folder)"
WScript.Quit 4

Function WaitAck(ByVal seconds)
    Dim t, i
    WaitAck = False
    t = Timer
    Do
        If fso.FileExists(STARTED_FILE) Or fso.FileExists(ERROR_FILE) Then
            WaitAck = True
            Exit Function
        End If
        WScript.Sleep 1000
        If Timer < t Then Exit Do
        If Timer - t >= seconds Then Exit Do
    Loop
End Function

Sub DeleteIfExists(ByVal p)
    On Error Resume Next
    If fso.FileExists(p) Then fso.DeleteFile p, True
End Sub

Sub LogLine(ByVal msg)
    On Error Resume Next
    Dim line, f
    line = "[" & Now & "] " & msg
    Set f = fso.OpenTextFile(LOG_FILE, 8, True)
    f.WriteLine line
    f.Close
    Set f = fso.OpenTextFile(STATUS_FILE, 2, True)
    f.WriteLine line
    f.Close
End Sub
