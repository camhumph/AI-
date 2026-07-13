param(
    [Parameter(Mandatory=$true)][string]$MacroPath,
    [Parameter(Mandatory=$true)][string]$SwExe,
    [Parameter(Mandatory=$true)][string]$ProgId,
    [string]$LogFile = "C:\Users\lenovo\Downloads\CMS_Quote_Log.txt",
    [string]$Procedure = "main",
    [int]$TimeoutSeconds = 90
)

$ErrorActionPreference = "Continue"

$MacroStatusFile = "C:\CMS_Local_Workspace\cms_macro_status.txt"
$MacroStartedFile = "C:\CMS_Local_Workspace\cms_macro_started.txt"
$MacroDoneFile = "C:\CMS_Local_Workspace\cms_macro_done.txt"
$MacroErrorFile = "C:\CMS_Local_Workspace\cms_macro_error.txt"
$TrainingTrigger = "C:\CMS_Local_Workspace\cms_training_xt.txt"
$HandoffFile = "C:\CMS_Local_Workspace\cms_handoff.txt"

function Write-LauncherLog {
    param([string]$Message)
    $line = ("[{0}] macro-runner: {1}" -f (Get-Date), $Message)
    foreach ($target in @(
        $LogFile,
        "C:\CMS_Local_Workspace\CMS_Quote_Log.txt"
    )) {
        try {
            $folder = Split-Path -Parent $target
            if ($folder -and -not (Test-Path -LiteralPath $folder)) {
                New-Item -ItemType Directory -Force -Path $folder | Out-Null
            }
            Add-Content -LiteralPath $target -Value $line
        } catch {
        }
    }
    try {
        Set-Content -LiteralPath "C:\CMS_Local_Workspace\cms_launcher_status.txt" -Value $line -Encoding UTF8
    } catch {
    }
}

function Remove-IfExists {
    param([string]$Path)
    try {
        if ($Path -and (Test-Path -LiteralPath $Path)) {
            Remove-Item -LiteralPath $Path -Force -ErrorAction SilentlyContinue
        }
    } catch {
    }
}

Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

[ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
interface IOleMessageFilter {
    [PreserveSig]
    int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
    [PreserveSig]
    int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType);
    [PreserveSig]
    int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType);
}

public class OleMessageFilter : IOleMessageFilter {
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter newFilter, out IOleMessageFilter oldFilter);

    public static void Register() {
        IOleMessageFilter oldFilter;
        CoRegisterMessageFilter(new OleMessageFilter(), out oldFilter);
    }

    public static void Revoke() {
        IOleMessageFilter oldFilter;
        CoRegisterMessageFilter(null, out oldFilter);
    }

    public int HandleInComingCall(int dwCallType, IntPtr htaskCaller, int dwTickCount, IntPtr lpInterfaceInfo) {
        return 0;
    }

    public int RetryRejectedCall(IntPtr htaskCallee, int dwTickCount, int dwRejectType) {
        if (dwRejectType == 2) return 250;
        return -1;
    }

    public int MessagePending(IntPtr htaskCallee, int dwTickCount, int dwPendingType) {
        return 2;
    }
}
"@

try {
    [OleMessageFilter]::Register()
    Write-LauncherLog "starting; macro=$MacroPath procedure=$Procedure"

    if (-not (Test-Path -LiteralPath $MacroPath)) {
        Write-LauncherLog "macro file not found: $MacroPath"
        exit 2
    }

    # Live quote handoff must win; clear stale launch status files.
    if (Test-Path -LiteralPath $HandoffFile) {
        Remove-IfExists $TrainingTrigger
    }
    Remove-IfExists $MacroStatusFile
    Remove-IfExists $MacroStartedFile
    Remove-IfExists $MacroDoneFile
    Remove-IfExists $MacroErrorFile

    $sw = $null
    try {
        $sw = [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
    } catch {
        $sw = $null
    }

    if ($null -eq $sw) {
        try {
            $sw = New-Object -ComObject $ProgId
        } catch {
            $sw = $null
        }
    }

    if ($null -eq $sw -and (Test-Path -LiteralPath $SwExe)) {
        Start-Process -FilePath $SwExe | Out-Null
        for ($i = 0; $i -lt 30 -and $null -eq $sw; $i++) {
            Start-Sleep -Seconds 2
            try {
                $sw = [Runtime.InteropServices.Marshal]::GetActiveObject($ProgId)
            } catch {
                $sw = $null
            }
        }
    }

    if ($null -eq $sw) {
        Write-LauncherLog "could not connect to SolidWorks"
        exit 3
    }

    try { $sw.Visible = $true } catch {}
    try { $sw.UserControl = $true } catch {}
    Start-Sleep -Seconds 3

    $moduleNames = @("Module6121", "Module61211", "Module612111", "Module1", "main", "Module2", "Module3")
    # Prefer main() so handoff routing (quote vs training) stays in VBA.
    $procedureNames = @($Procedure, "main", "RunFromLauncher") | Select-Object -Unique

    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    $ran = $false
    $attempt = 0

    while ((Get-Date) -lt $deadline -and -not $ran) {
        $attempt++
        Remove-IfExists $MacroStartedFile
        Remove-IfExists $MacroErrorFile

        foreach ($procName in $procedureNames) {
            foreach ($moduleName in $moduleNames) {
                try {
                    try { $sw.CommandInProgress = $false } catch {}
                    try { $sw.UserControl = $true } catch {}
                    $macroErr = 0
                    $ok2 = $false
                    $vbaErr = 0
                    # Never set CommandInProgress=$true before RunMacro — SW 2023 returns False/err=0.
                    try {
                        $ok2 = $sw.RunMacro2($MacroPath, $moduleName, $procName, 0, [ref]$macroErr)
                    } catch {
                        $ok2 = $false
                        $vbaErr = 1
                    }
                    if ($ok2 -ne $true) {
                        try {
                            $macroErr = 0
                            $ok2 = $sw.RunMacro2($MacroPath, $moduleName, $procName, 1, [ref]$macroErr)
                        } catch {
                            $ok2 = $false
                        }
                    }
                    if ($ok2 -ne $true) {
                        try {
                            $ok2 = $sw.RunMacro($MacroPath, $moduleName, $procName)
                        } catch {
                            $ok2 = $false
                        }
                    }
                    Write-LauncherLog "attempt=$attempt RunMacro module='$moduleName' proc='$procName' ok=$ok2 err=$macroErr"
                } catch {
                    Write-LauncherLog ("RunMacro failed module='$moduleName' proc='$procName': " + $_.Exception.Message)
                }

                $waitUntil = (Get-Date).AddSeconds(10)
                while ((Get-Date) -lt $waitUntil) {
                    if (Test-Path -LiteralPath $MacroStartedFile) {
                        Write-LauncherLog "macro acknowledged STARTED via $MacroStartedFile"
                        $ran = $true
                        break
                    }
                    if (Test-Path -LiteralPath $MacroErrorFile) {
                        Write-LauncherLog "macro wrote ERROR file: $MacroErrorFile"
                        $ran = $true
                        break
                    }
                    Start-Sleep -Seconds 1
                }
                if ($ran) { break }
            }
            if ($ran) { break }
        }

        if (-not $ran) {
            Start-Sleep -Seconds 2
        }
    }

    if (-not $ran) {
        Write-LauncherLog "SolidWorks opened, but Module6121 did not acknowledge launch within ${TimeoutSeconds}s"
        exit 4
    }
} finally {
    try { [OleMessageFilter]::Revoke() } catch {}
}
