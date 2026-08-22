<#
.SYNOPSIS
    Converts legacy .doc files to .docx format using Word COM Automation.

.DESCRIPTION
    Batch-converts .doc (Word 97-2003) files to .docx (Open XML) format.
    Supports single file, folder, and recursive folder processing.
    Produces structured output suitable for pipeline consumption and logging.

.PARAMETER Path
    Path to a .doc file or a directory containing .doc files.

.PARAMETER Recurse
    When Path is a directory, include all subdirectories.

.PARAMETER Overwrite
    Overwrite existing .docx files. Without this switch, existing targets are skipped.

.PARAMETER LogPath
    Optional. Path to a log file. If specified, all activity is appended to this file.

.EXAMPLE
    .\Convert-DocToDocx.ps1 -Path "C:\Reports" -Recurse -Overwrite

.EXAMPLE
    .\Convert-DocToDocx.ps1 -Path "C:\Reports\Q1.doc" -LogPath "C:\Logs\convert.log"
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to a .doc file or directory.")]
    [ValidateScript({
        if (-not (Test-Path -LiteralPath $_)) {
            throw "Path does not exist: $_"
        }
        return $true
    })]
    [string]$Path,

    [switch]$Recurse,
    [switch]$Overwrite,

    [Parameter(Mandatory = $false)]
    [string]$LogPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

#region Logging

function Write-Log {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry     = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "ERROR" { Write-Host $entry -ForegroundColor Red }
        "WARN"  { Write-Host $entry -ForegroundColor Yellow }
        default { Write-Host $entry }
    }

    if ($LogPath) {
        $entry | Out-File -FilePath $LogPath -Append -Encoding UTF8
    }
}

#endregion

#region File Discovery

function Get-DocFiles {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,

        [switch]$IncludeChildren
    )

    $resolved = Resolve-Path -LiteralPath $InputPath
    $item     = Get-Item -LiteralPath $resolved

    if ($item.PSIsContainer) {
        $params = @{
            LiteralPath = $item.FullName
            File        = $true
            Filter      = "*.doc"           # Filter at provider level; faster than Where-Object
            Recurse     = $IncludeChildren.IsPresent
        }

        # Exclude .docx, .docm etc. that also match "*.doc*" at provider level
        $files = Get-ChildItem @params |
            Where-Object { $_.Extension -ieq ".doc" } |
            Sort-Object FullName

        return @($files)   # Force array; fixes .Count bug on single-item return
    }

    if ($item.Extension -ine ".doc") {
        throw "Input file must have a .doc extension: $($item.FullName)"
    }

    return @(, $item)      # Force single-element array
}

#endregion

#region COM Helpers

function Close-WordDocument {
    param([System.__ComObject]$Document)

    if ($null -eq $Document) { return }

    try {
        $wdDoNotSaveChanges = 0
        $Document.Close([ref]$wdDoNotSaveChanges)
    }
    catch {
        Write-Log "Warning: Could not close document gracefully - $($_.Exception.Message)" -Level WARN
    }
    finally {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($Document)
        }
        catch { }
    }
}

function Stop-WordProcess {
    <#
    .SYNOPSIS
        Last-resort cleanup; kills any orphaned WINWORD process started by this session.
    #>
    param([int]$ProcessId)

    if ($ProcessId -le 0) { return }

    try {
        $proc = Get-Process -Id $ProcessId -ErrorAction SilentlyContinue
        if ($null -ne $proc -and -not $proc.HasExited) {
            Write-Log "Force-killing orphaned Word process (PID $ProcessId)." -Level WARN
            Stop-Process -Id $ProcessId -Force -ErrorAction SilentlyContinue
        }
    }
    catch { }
}

function Get-ComObjectProcessId {
    <#
    .SYNOPSIS
        Retrieves the PID of a COM-spawned Word instance via Win32 API.
    #>
    param([System.__ComObject]$WordApp)

    try {
        $hwnd = $WordApp.Application.Hwnd
        $pid  = [uint32]0
        $null = [Win32Api.User32]::GetWindowThreadProcessId([IntPtr]$hwnd, [ref]$pid)
        return [int]$pid
    }
    catch {
        return 0
    }
}

# P/Invoke for GetWindowThreadProcessId
if (-not ([System.Management.Automation.PSTypeName]"Win32Api.User32").Type) {
    Add-Type -Namespace Win32Api -Name User32 -MemberDefinition @"
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
"@
}

#endregion

#region Main Execution

Write-Log "Script started. Path=$Path | Recurse=$Recurse | Overwrite=$Overwrite"

# 1. Collect target files
$docFiles = @(Get-DocFiles -InputPath $Path -IncludeChildren:$Recurse)

if ($docFiles.Count -eq 0) {
    Write-Log "No .doc files found under '$Path'." -Level WARN
    exit 0
}

Write-Log "Found $($docFiles.Count) .doc file(s) to process."

# 2. Initialize result containers
$converted = [System.Collections.Generic.List[pscustomobject]]::new()
$skipped   = [System.Collections.Generic.List[pscustomobject]]::new()
$failed    = [System.Collections.Generic.List[pscustomobject]]::new()

# 3. Word COM Automation
$word      = $null
$wordPid   = 0
$wdFormatXMLDocument = 12   # .docx (Open XML); wdFormatXMLDocument enum value

try {
    $word              = New-Object -ComObject Word.Application
    $word.Visible      = $false
    $word.DisplayAlerts = 0   # wdAlertsNone
    $word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
    $word.Options.UpdateLinksAtOpen = $false

    $wordPid = Get-ComObjectProcessId -WordApp $word
    Write-Log "Word COM instance created (PID $wordPid)."

    $totalFiles  = $docFiles.Count
    $currentFile = 0

    foreach ($file in $docFiles) {
        $currentFile++
        $sourcePath = $file.FullName
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $pctComplete = [math]::Round(($currentFile / $totalFiles) * 100)

        Write-Progress -Activity "Converting .doc to .docx" `
                       -Status "$currentFile / $totalFiles - $($file.Name)" `
                       -PercentComplete $pctComplete

        # Skip logic
        if ((Test-Path -LiteralPath $targetPath) -and -not $Overwrite) {
            $skipped.Add([pscustomobject]@{
                Source = $sourcePath
                Target = $targetPath
                Reason = "Target already exists"
            })
            Write-Log "Skipped (exists): $sourcePath" -Level WARN
            continue
        }

        if ((Test-Path -LiteralPath $targetPath) -and $Overwrite) {
            Remove-Item -LiteralPath $targetPath -Force
            Write-Log "Removed existing target: $targetPath"
        }

        # Convert
        $document = $null

        try {
            # Open: ConfirmConversions=$false, ReadOnly=$true, AddToRecentFiles=$false
            $document = $word.Documents.Open(
                $sourcePath,        # FileName
                $false,             # ConfirmConversions
                $true,              # ReadOnly
                $false,             # AddToRecentFiles
                [Type]::Missing,    # PasswordDocument
                [Type]::Missing,    # PasswordTemplate
                $true               # Revert (revert if already open)
            )

            $document.SaveAs2([ref]$targetPath, [ref]$wdFormatXMLDocument)

            $converted.Add([pscustomobject]@{
                Source = $sourcePath
                Target = $targetPath
            })
            Write-Log "Converted: $sourcePath -> $targetPath"
        }
        catch {
            $failed.Add([pscustomobject]@{
                Source = $sourcePath
                Error  = $_.Exception.Message
            })
            Write-Log "FAILED: $sourcePath - $($_.Exception.Message)" -Level ERROR
        }
        finally {
            Close-WordDocument -Document $document
        }
    }

    Write-Progress -Activity "Converting .doc to .docx" -Completed
}
finally {
    # Graceful shutdown
    if ($null -ne $word) {
        try {
            $word.Quit([ref]0)
            Write-Log "Word COM instance closed gracefully."
        }
        catch {
            Write-Log "Word.Quit() failed - $($_.Exception.Message)" -Level WARN
        }
        finally {
            try {
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
            }
            catch { }
        }
    }

    # Fallback: force-kill if still running
    Start-Sleep -Milliseconds 500
    Stop-WordProcess -ProcessId $wordPid

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

#endregion

#region Summary Report

Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "  CONVERSION SUMMARY" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Converted : $($converted.Count)" -ForegroundColor Green
Write-Host "  Skipped   : $($skipped.Count)"   -ForegroundColor Yellow
Write-Host "  Failed    : $($failed.Count)"     -ForegroundColor Red
Write-Host ""

if ($converted.Count -gt 0) {
    Write-Host "-- Converted --" -ForegroundColor Green
    $converted | Format-Table Source, Target -AutoSize
}

if ($skipped.Count -gt 0) {
    Write-Host "-- Skipped --" -ForegroundColor Yellow
    $skipped | Format-Table Source, Target, Reason -AutoSize
}

if ($failed.Count -gt 0) {
    Write-Host "-- Failed --" -ForegroundColor Red
    $failed | Format-Table Source, Error -AutoSize
}

#endregion

#region Pipeline Output & Exit

# Emit structured output for pipeline consumers
[pscustomobject]@{
    TotalFiles    = $docFiles.Count
    Converted     = $converted.Count
    Skipped       = $skipped.Count
    Failed        = $failed.Count
    ConvertedList = $converted
    SkippedList   = $skipped
    FailedList    = $failed
}

if ($failed.Count -gt 0) {
    exit 1
}

exit 0

#endregion
