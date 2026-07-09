' ============================================================
' RunTrainingXtLauncher.vbs
' Same SolidWorks 2023 launch path as CMS_Launcher.vbs, but runs
' Module6121 RunTrainingXtExport (training XT dimension export).
'
' Python writes C:\CMS_Local_Workspace\cms_training_xt.txt first, then:
'   wscript RunTrainingXtLauncher.vbs
' ============================================================

Const LOCAL_WORKSPACE_ROOT = "C:\CMS_Local_Workspace"
Const TRAINING_HANDOFF     = "C:\CMS_Local_Workspace\cms_training_xt.txt"
Const SW_EXE               = "C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS (3)\SLDWORKS.EXE"
Const SW_PROGID            = "SldWorks.Application.31"
Const SW_MACRO             = "C:\CMS_Local_Workspace\Module6121.swb"
Const LOG_FILE             = "C:\CMS_Local_Workspace\CMS_Training_XT_Launcher_Log.txt"

Dim fso, shell
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

If Not fso.FolderExists(LOCAL_WORKSPACE_ROOT) Then fso.CreateFolder LOCAL_WORKSPACE_ROOT

LogStep "===== training XT launcher started ====="

If Not fso.FileExists(TRAINING_HANDOFF) Then
    LogStep "ERROR: handoff file missing: " & TRAINING_HANDOFF
    WriteDoneFromHandoff "ERROR", "", "Handoff file missing"
    WScript.Quit 1
End If

Dim doneFile
doneFile = ReadHandoffValue("DoneFile")
If doneFile = "" Then doneFile = LOCAL_WORKSPACE_ROOT & "\cms_training_xt_done.txt"

If Not LaunchSolidWorksTrainingMacro() Then
    LogStep "ERROR: SolidWorks 2023 macro did not start"
    WriteDoneFile doneFile, "ERROR", "", "SolidWorks 2023 macro failed to start — check " & LOG_FILE
    WScript.Quit 2
End If

LogStep "===== training XT macro started (SolidWorks 2023) ====="
WScript.Quit 0

' ------------------------------------------------------------
Sub LogStep(msg)
    On Error Resume Next
    Dim f
    Set f = fso.OpenTextFile(LOG_FILE, 8, True)
    f.WriteLine "[" & Now & "] " & msg
    f.Close
End Sub

Function ReadHandoffValue(key)
    ReadHandoffValue = ""
    On Error Resume Next
    Dim f, line, p, k, v
    Set f = fso.OpenTextFile(TRAINING_HANDOFF, 1)
    Do While Not f.AtEndOfStream
        line = Trim(f.ReadLine)
        p = InStr(line, "=")
        If p > 0 Then
            k = Trim(Left(line, p - 1))
            v = Trim(Mid(line, p + 1))
            If UCase(k) = UCase(key) Then ReadHandoffValue = v
        End If
    Loop
    f.Close
End Function

Sub WriteDoneFromHandoff(status, xtCsv, message)
    Dim df
    df = ReadHandoffValue("DoneFile")
    If df = "" Then df = LOCAL_WORKSPACE_ROOT & "\cms_training_xt_done.txt"
    WriteDoneFile df, status, xtCsv, message
End Sub

Sub WriteDoneFile(path, status, xtCsv, message)
    On Error Resume Next
    Dim parent
    parent = fso.GetParentFolderName(path)
    If parent <> "" And Not fso.FolderExists(parent) Then fso.CreateFolder parent
    Dim f
    Set f = fso.CreateTextFile(path, True)
    f.WriteLine "Status=" & status
    f.WriteLine "XtCsv=" & xtCsv
    f.WriteLine "PartCount=0"
    f.WriteLine "Message=" & message
    f.Close
End Sub

Function LaunchSolidWorksTrainingMacro()
    LaunchSolidWorksTrainingMacro = False
    Dim sw, tries, macroPath, macroFolder
    macroFolder = fso.GetParentFolderName(WScript.ScriptFullName)
    macroPath = SW_MACRO
    If Not fso.FileExists(macroPath) Then macroPath = macroFolder & "\Module6121.swb"
    If Not fso.FileExists(macroPath) Then macroPath = LOCAL_WORKSPACE_ROOT & "\Module6121.swp"
    If Not fso.FileExists(macroPath) Then macroPath = macroFolder & "\Module6121.swp"

    LogStep "macro path: " & macroPath
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
    WScript.Sleep 1500

    If Not fso.FileExists(macroPath) Then
        LogStep "Module6121 macro not found"
        Exit Function
    End If

    Dim modNames, mi, okRun, ran, procNames, pi
    modNames = Array("Module6121", "Module61211", "Module612111", "Module1", "main", "Module2", "Module3")
    procNames = Array("RunTrainingXtExport", "main")
    ran = False

    For pi = 0 To UBound(procNames)
        For mi = 0 To UBound(modNames)
            On Error Resume Next
            sw.CommandInProgress = True
            okRun = sw.RunMacro(macroPath, modNames(mi), procNames(pi))
            sw.CommandInProgress = False
            If Err.Number = 0 And okRun = True Then
                On Error GoTo 0
                ran = True
                LogStep "macro started: module=" & modNames(mi) & " proc=" & procNames(pi)
                Exit For
            End If
            LogStep "RunMacro failed module=" & modNames(mi) & " proc=" & procNames(pi) & " err=" & Err.Number
            Err.Clear
            On Error GoTo 0
        Next
        If ran Then Exit For
    Next

    LaunchSolidWorksTrainingMacro = ran
End Function
