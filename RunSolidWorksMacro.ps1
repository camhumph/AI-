param(
    [Parameter(Mandatory=$true)][string]$MacroPath,
    [Parameter(Mandatory=$true)][string]$SwExe,
    [Parameter(Mandatory=$true)][string]$ProgId,
    [string]$LogFile = "C:\Users\lenovo\Downloads\CMS_Quote_Log.txt",
    [string]$Procedure = "main",
    [int]$TimeoutSeconds = 120
)

$ErrorActionPreference = "Continue"

$LocalWorkspace = "C:\CMS_Local_Workspace"
$MacroStatusFile = "$LocalWorkspace\cms_macro_status.txt"
$MacroStartedFile = "$LocalWorkspace\cms_macro_started.txt"
$MacroDoneFile = "$LocalWorkspace\cms_macro_done.txt"
$MacroErrorFile = "$LocalWorkspace\cms_macro_error.txt"
$TrainingTrigger = "$LocalWorkspace\cms_training_xt.txt"
$HandoffFile = "$LocalWorkspace\cms_handoff.txt"

function Write-LauncherLog {
    param([string]$Message)
    $line = ("[{0}] macro-runner: {1}" -f (Get-Date), $Message)
    foreach ($target in @(
        $LogFile,
        "$LocalWorkspace\CMS_Quote_Log.txt"
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
        Set-Content -LiteralPath "$LocalWorkspace\cms_launcher_status.txt" -Value $line -Encoding UTF8
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

function Wait-ForMacroAck {
    param([int]$Seconds = 12)
    $until = (Get-Date).AddSeconds($Seconds)
    while ((Get-Date) -lt $until) {
        if (Test-Path -LiteralPath $MacroStartedFile) {
            Write-LauncherLog "macro acknowledged STARTED via $MacroStartedFile"
            return $true
        }
        if (Test-Path -LiteralPath $MacroErrorFile) {
            Write-LauncherLog "macro wrote ERROR file: $MacroErrorFile"
            return $true
        }
        Start-Sleep -Seconds 1
    }
    return $false
}

function Get-MacroEntryPoints {
    param($Sw, [string]$Path)
    $entries = New-Object System.Collections.Generic.List[object]
    try {
        # swMethodsWithoutArguments = 1 (SW 2020+); try 0 and 1
        foreach ($opt in @(1, 0, 2)) {
            try {
                $methods = $Sw.GetMacroMethods($Path, $opt)
                if ($null -eq $methods) { continue }
                foreach ($m in @($methods)) {
                    if (-not $m) { continue }
                    $parts = [string]$m -split "\.", 2
                    if ($parts.Count -ge 2) {
                        $entries.Add([pscustomobject]@{ Module = $parts[0]; Proc = $parts[1]; Raw = [string]$m }) | Out-Null
                    }
                }
                if ($entries.Count -gt 0) { break }
            } catch {
            }
        }
    } catch {
        Write-LauncherLog ("GetMacroMethods failed: " + $_.Exception.Message)
    }
    return $entries
}

function Invoke-RunMacroCom {
    param($Sw, [string]$Path, [string]$Module, [string]$Proc)
    $ok = $false
    $errCode = [int]0

    # 1) RunMacro (no ByRef) — most reliable from PowerShell
    try {
        $ok = [bool]$Sw.RunMacro($Path, $Module, $Proc)
    } catch {
        $ok = $false
        Write-LauncherLog ("RunMacro exception module='$Module' proc='$Proc': " + $_.Exception.Message)
    }
    if ($ok) { return @{ Ok = $true; Err = 0; Via = "RunMacro" } }

    # 2) RunMacro2 with explicit Int32 ByRef (PowerShell often breaks Long ByRef)
    foreach ($opt in @([int]0, [int]1)) {
        try {
            $errCode = [int]0
            $ok = [bool]$Sw.RunMacro2($Path, $Module, $Proc, $opt, [ref]$errCode)
            if ($ok) { return @{ Ok = $true; Err = $errCode; Via = "RunMacro2 opt=$opt" } }
        } catch {
            Write-LauncherLog ("RunMacro2 exception module='$Module' proc='$Proc' opt=$opt: " + $_.Exception.Message)
        }
    }
    return @{ Ok = $false; Err = $errCode; Via = "none" }
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
    Write-LauncherLog "macro-runner-v3: starting; macro=$MacroPath procedure=$Procedure"

    if (-not (Test-Path -LiteralPath $MacroPath)) {
        Write-LauncherLog "macro file not found: $MacroPath"
        exit 2
    }

    try {
        $fi = Get-Item -LiteralPath $MacroPath
        Write-LauncherLog ("macro file size={0} bytes modified={1}" -f $fi.Length, $fi.LastWriteTime)
        if ($fi.Length -lt 1000) {
            Write-LauncherLog "WARNING: Module6121.swp looks too small — recompile Module6121.bas to .swp in SolidWorks VBA editor"
        }
        # Detect accidental text/.swb renamed to .swp
        $head = Get-Content -LiteralPath $MacroPath -TotalCount 1 -ErrorAction SilentlyContinue
        if ($head -match "Attribute VB_Name|VERSION 5\.00|Begin\s+\{") {
            Write-LauncherLog "ERROR: $MacroPath looks like text/.swb source, not a compiled .swp. Recompile in SolidWorks (File > Save as .swp)."
        }
    } catch {
    }

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
    try { $sw.CommandInProgress = $false } catch {}
    Start-Sleep -Seconds 2

    # Discover real module/proc names from the .swp (beats guessing Module61211 etc.)
    $discovered = @(Get-MacroEntryPoints -Sw $sw -Path $MacroPath)
    if ($discovered.Count -gt 0) {
        foreach ($e in $discovered) {
            Write-LauncherLog ("GetMacroMethods entry: {0}.{1}" -f $e.Module, $e.Proc)
        }
    } else {
        Write-LauncherLog "GetMacroMethods returned no entry points — .swp may be corrupt/stale or macros disabled in SolidWorks options"
    }

    $pairs = New-Object System.Collections.Generic.List[object]
    foreach ($e in $discovered) {
        $pairs.Add([pscustomobject]@{ Module = $e.Module; Proc = $e.Proc }) | Out-Null
    }
    # Prefer main / RunFromLauncher from discovered list first, then guesses.
    foreach ($procName in @($Procedure, "main", "RunFromLauncher") | Select-Object -Unique) {
        foreach ($moduleName in @("Module61211", "Module6121", "Module1")) {
            $pairs.Add([pscustomobject]@{ Module = $moduleName; Proc = $procName }) | Out-Null
        }
    }

    $deadline = (Get-Date).AddSeconds([Math]::Max(30, $TimeoutSeconds - 25))
    $ran = $false
    $attempt = 0
    $seen = @{}

    while ((Get-Date) -lt $deadline -and -not $ran) {
        $attempt++
        Remove-IfExists $MacroStartedFile
        Remove-IfExists $MacroErrorFile
        try { $sw.CommandInProgress = $false } catch {}

        foreach ($pair in $pairs) {
            $key = "$($pair.Module)|$($pair.Proc)"
            if ($attempt -eq 1 -and $seen.ContainsKey($key)) { continue }
            $seen[$key] = $true

            $result = Invoke-RunMacroCom -Sw $sw -Path $MacroPath -Module $pair.Module -Proc $pair.Proc
            Write-LauncherLog ("attempt={0} via={1} module='{2}' proc='{3}' ok={4} err={5}" -f `
                $attempt, $result.Via, $pair.Module, $pair.Proc, $result.Ok, $result.Err)

            if (Wait-ForMacroAck -Seconds 8) {
                $ran = $true
                break
            }
        }

        if (-not $ran) {
            Start-Sleep -Seconds 2
        }
    }

    # Fallback: command-line /m (new or same SW) — works when COM RunMacro is blocked.
    if (-not $ran -and (Test-Path -LiteralPath $SwExe)) {
        Write-LauncherLog "COM RunMacro failed — falling back to SLDWORKS.EXE /m"
        Remove-IfExists $MacroStartedFile
        Remove-IfExists $MacroErrorFile
        try {
            Start-Process -FilePath $SwExe -ArgumentList @("/m", $MacroPath) | Out-Null
            if (Wait-ForMacroAck -Seconds 45) {
                $ran = $true
            } else {
                Write-LauncherLog "SLDWORKS.EXE /m did not produce cms_macro_started.txt within 45s"
            }
        } catch {
            Write-LauncherLog ("SLDWORKS.EXE /m failed: " + $_.Exception.Message)
        }
    }

    if (-not $ran) {
        Write-LauncherLog "SolidWorks opened, but Module6121 did not acknowledge launch within ${TimeoutSeconds}s"
        Write-LauncherLog "FIX: In SolidWorks VBA editor, import Module6121.bas, Debug>Compile, File>Save As Module6121.swp into C:\CMS_Local_Workspace\"
        Write-LauncherLog "FIX: Tools>Options>System Options>Macro — allow macros / trusted locations"
        exit 4
    }
} finally {
    try { [OleMessageFilter]::Revoke() } catch {}
}
