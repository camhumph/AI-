param(
    [Parameter(Mandatory=$true)][string]$MacroPath,
    [Parameter(Mandatory=$true)][string]$SwExe,
    [Parameter(Mandatory=$true)][string]$ProgId,
    [string]$LogFile = "C:\Users\lenovo\Downloads\CMS_Quote_Log.txt",
    [string]$Procedure = "RunFromLauncher"
)

$ErrorActionPreference = "Continue"

function Write-LauncherLog {
    param([string]$Message)
    try {
        $folder = Split-Path -Parent $LogFile
        if ($folder -and -not (Test-Path -LiteralPath $folder)) {
            New-Item -ItemType Directory -Force -Path $folder | Out-Null
        }
        Add-Content -LiteralPath $LogFile -Value ("[{0}] macro-runner: {1}" -f (Get-Date), $Message)
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
    Write-LauncherLog "starting; macro=$MacroPath"

    if (-not (Test-Path -LiteralPath $MacroPath)) {
        Write-LauncherLog "macro file not found: $MacroPath"
        exit 2
    }

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
    Start-Sleep -Milliseconds 750

    $moduleNames = @("Module6121", "Module61211", "Module612111", "Module1", "main", "Module2", "Module3")
    $ext = [System.IO.Path]::GetExtension($MacroPath).ToLowerInvariant()
    if ($ext -eq ".swp") {
        $procedureNames = @($Procedure)
    } else {
        $procedureNames = @($Procedure, "RunFromLauncher", "main")
    }

    $ran = $false
    foreach ($procName in $procedureNames) {
        foreach ($moduleName in $moduleNames) {
            try {
                try { $sw.CommandInProgress = $true } catch {}
                $ok = $sw.RunMacro($MacroPath, $moduleName, $procName)
                try { $sw.CommandInProgress = $false } catch {}
                if ($ok -eq $true) {
                    Write-LauncherLog "macro started via RunMacro module '$moduleName' procedure '$procName'"
                    $ran = $true
                    break
                }
            } catch {
                try { $sw.CommandInProgress = $false } catch {}
                Write-LauncherLog ("RunMacro module '$moduleName' procedure '$procName' failed: " + $_.Exception.Message)
            }

            if (-not $ran) {
                try {
                    try { $sw.CommandInProgress = $true } catch {}
                    $macroErr = 0
                    $ok2 = $sw.RunMacro2($MacroPath, $moduleName, $procName, 1, [ref]$macroErr)
                    try { $sw.CommandInProgress = $false } catch {}
                    if ($ok2 -eq $true) {
                        Write-LauncherLog "macro started via RunMacro2 module '$moduleName' procedure '$procName' err=$macroErr"
                        $ran = $true
                        break
                    } else {
                        Write-LauncherLog "RunMacro2 returned false for module '$moduleName' procedure '$procName' err=$macroErr"
                    }
                } catch {
                    try { $sw.CommandInProgress = $false } catch {}
                    Write-LauncherLog ("RunMacro2 module '$moduleName' procedure '$procName' failed: " + $_.Exception.Message)
                }
            }
            Start-Sleep -Milliseconds 300
        }
        if ($ran) { break }
    }

    if (-not $ran) {
        Write-LauncherLog "RunMacro/RunMacro2 could not start allowed procedure(s) for any known module name"
        exit 4
    }
} finally {
    try { [OleMessageFilter]::Revoke() } catch {}
}
