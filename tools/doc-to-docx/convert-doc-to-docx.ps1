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

.PARAMETER OutputFormat
    Output records as the default human-readable table or as one compressed JSON array.

.PARAMETER TimeoutSeconds
    Maximum time for one document's Word conversion. Defaults to 300 seconds.

.PARAMETER Password
    Alias for DocumentKey. Accepts a String or SecureString. When omitted,
    HWPX_WIZ_DOC_PASSWORD is used if present.

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
    [string]$LogPath,

    [ValidateSet("Table", "Json")]
    [string]$OutputFormat = "Table",

    [ValidateRange(1, 86400)]
    [int]$TimeoutSeconds = 300,

    [Parameter(Mandatory = $false)]
    [Alias("Password")]
    [ValidateScript({
        if ($_ -is [string] -or $_ -is [System.Security.SecureString]) {
            return $true
        }
        throw "DocumentKey must be a String or SecureString."
    })]
    [object]$DocumentKey
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$isDotSourced = $MyInvocation.InvocationName -eq "."
$script:ConversionLogPath = $null

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

    if ($OutputFormat -eq "Table") {
        switch ($Level) {
            "ERROR" { Write-Information $entry -InformationAction Continue }
            "WARN"  { Write-Information $entry -InformationAction Continue }
            default { Write-Information $entry -InformationAction Continue }
        }
    }

    if ($script:ConversionLogPath) {
        Add-Content -LiteralPath $script:ConversionLogPath -Value $entry -Encoding UTF8
    }
}

function Resolve-DocumentKey {
    param(
        [Parameter()]
        [object]$Value
    )

    if ($null -eq $Value) {
        $Value = $env:HWPX_WIZ_DOC_PASSWORD
    }
    if ($null -eq $Value -or ($Value -is [string] -and [string]::IsNullOrEmpty($Value))) {
        return $null
    }
    if ($Value -is [string]) {
        return $Value
    }
    if ($Value -isnot [System.Security.SecureString]) {
        throw "DocumentKey must be a String or SecureString."
    }

    $passwordPointer = [IntPtr]::Zero
    try {
        $passwordPointer = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Value)
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($passwordPointer)
    }
    finally {
        if ($passwordPointer -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($passwordPointer)
        }
    }
}

#endregion

#region File Discovery

function Get-DocFile {
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
        catch {
            Write-Log "Could not release Word document COM object - $($_.Exception.Message)" -Level WARN
        }
    }
}

if (-not ([System.Management.Automation.PSTypeName]"Win32Api.User32").Type) {
    Add-Type -Namespace Win32Api -Name User32 -MemberDefinition @"
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
"@
}

if (-not ([System.Management.Automation.PSTypeName]"HwpxWiz.ProcessTimeoutGuard").Type) {
    Add-Type -TypeDefinition @"
using System;
using System.Diagnostics;
using System.Threading;

namespace HwpxWiz
{
    public sealed class ProcessTimeoutGuard : IDisposable
    {
        private Timer timer;
        private Process process;
        private int lifecycle;
        private int disposeStarted;
        private int timedOut;

        public bool TimedOut
        {
            get { return Volatile.Read(ref timedOut) == 1; }
        }

        public void Arm(Process ownedProcess, int timeoutMilliseconds)
        {
            process = ownedProcess;
            timer = new Timer(OnTimeout, null, timeoutMilliseconds, Timeout.Infinite);
        }

        private void OnTimeout(object state)
        {
            if (Interlocked.CompareExchange(ref lifecycle, 2, 0) != 0) { return; }
            Interlocked.Exchange(ref timedOut, 1);

            try
            {
                if (process != null && !process.HasExited)
                {
                    process.Kill();
                }
            }
            catch
            {
                // The main PowerShell runspace reports any resulting COM failure.
            }
        }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref disposeStarted, 1) == 1) { return; }
            Interlocked.CompareExchange(ref lifecycle, 1, 0);
            if (timer == null) { return; }

            using (var completed = new ManualResetEvent(false))
            {
                if (timer.Dispose(completed))
                {
                    completed.WaitOne();
                }
            }
        }
    }
}
"@
}

function Get-WordProcess {
    param([System.__ComObject]$WordApp)

    $probeDocument = $null
    try {
        $probeDocument = $WordApp.Documents.Add()
        $windowHandle = $probeDocument.ActiveWindow.Hwnd
        if ($null -eq $windowHandle -or [int]$windowHandle -eq 0) {
            return $null
        }

        [uint32]$ownedProcessId = 0
        $null = [Win32Api.User32]::GetWindowThreadProcessId(
            [IntPtr]$windowHandle,
            [ref]$ownedProcessId
        )
        if ($ownedProcessId -le 0) {
            return $null
        }

        return Get-Process -Id $ownedProcessId -ErrorAction Stop
    }
    catch {
        Write-Log "Could not resolve Word process - $($_.Exception.Message)" -Level WARN
        return $null
    }
    finally {
        Close-WordDocument -Document $probeDocument
    }
}

function Stop-OwnedWordProcess {
    [CmdletBinding(SupportsShouldProcess)]
    param([System.Diagnostics.Process]$Process)

    if ($null -eq $Process) { return }

    try {
        if ($Process.HasExited) { return }
        if ($Process.WaitForExit(5000)) { return }

        if ($PSCmdlet.ShouldProcess("WINWORD process $($Process.Id)", "Force stop")) {
            Write-Log "Force-killing orphaned Word process (PID $($Process.Id))." -Level WARN
            $Process.Kill()
        }
    }
    catch {
        Write-Log "Could not inspect or stop owned Word process - $($_.Exception.Message)" -Level WARN
    }
}

#endregion

#region Main Execution

# 1. Collect target files
$docFiles = @(Get-DocFile -InputPath $Path -IncludeChildren:$Recurse)

if ($LogPath) {
    if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($LogPath)) {
        $record = [pscustomobject]@{
            Status = "Failed"
            Source = $Path
            Target = $null
            Reason = "LogPath must be literal"
            Error = "LogPath must not contain wildcard characters: $LogPath"
        }
        if ($isDotSourced) {
            return @($record)
        }
        if ($OutputFormat -eq "Json") {
            ConvertTo-Json -InputObject @($record) -Compress
        }
        else {
            $record
        }
        exit 1
    }

    $logParent = Split-Path -Parent ([System.IO.Path]::GetFullPath($LogPath))
    if ([string]::IsNullOrWhiteSpace($logParent)) {
        $logParent = (Get-Location).Path
    }
    if (-not (Test-Path -LiteralPath $logParent -PathType Container)) {
        $record = [pscustomobject]@{
            Status = "Failed"
            Source = $Path
            Target = $null
            Reason = "LogPath parent does not exist"
            Error = "LogPath parent does not exist: $logParent"
        }
        if ($OutputFormat -eq "Table") {
            Write-Information $record.Error -InformationAction Continue
        }
        if ($isDotSourced) {
            return @($record)
        }
        if ($OutputFormat -eq "Json") {
            ConvertTo-Json -InputObject @($record) -Compress
        }
        else {
            $record
        }
        exit 1
    }

    $fullLogPath = [System.IO.Path]::GetFullPath($LogPath)
    $logCandidates = @(
        foreach ($docFile in $docFiles) {
            $docFile.FullName
            [System.IO.Path]::ChangeExtension($docFile.FullName, ".docx")
        }
    )
    $logConflicts = @(
        $logCandidates |
            Where-Object {
                [System.StringComparer]::OrdinalIgnoreCase.Equals($_, $fullLogPath)
            }
    )
    if ($logConflicts.Count -gt 0) {
        $record = [pscustomobject]@{
            Status = "Failed"
            Source = $Path
            Target = $fullLogPath
            Reason = "LogPath conflicts with a source or output"
            Error = "LogPath must not overwrite a source or output: $fullLogPath"
        }
        if ($OutputFormat -eq "Table") {
            Write-Information $record.Error -InformationAction Continue
        }
        if ($isDotSourced) {
            return @($record)
        }
        if ($OutputFormat -eq "Json") {
            ConvertTo-Json -InputObject @($record) -Compress
        }
        else {
            $record
        }
        exit 1
    }

    $script:ConversionLogPath = $fullLogPath
}

Write-Log "Script started. Path=$Path | Recurse=$Recurse | Overwrite=$Overwrite"

if ($docFiles.Count -eq 0) {
    Write-Log "No .doc files found under '$Path'." -Level WARN
    if ($isDotSourced) {
        return @()
    }
    if ($OutputFormat -eq "Json") {
        ConvertTo-Json -InputObject @() -Compress
    }
    exit 0
}

Write-Log "Found $($docFiles.Count) .doc file(s) to process."

# 2. Initialize result containers
$converted = [System.Collections.Generic.List[pscustomobject]]::new()
$skipped   = [System.Collections.Generic.List[pscustomobject]]::new()
$failed    = [System.Collections.Generic.List[pscustomobject]]::new()

# 3. Word COM Automation
$word        = $null
$wordProcess = $null
$wordPid     = 0
$wordProvenance = "Unknown"
$wordUserControl = $null
$originalDisplayAlerts = $null
$originalAutomationSecurity = $null
$originalUpdateLinks = $null
$settingsCaptured = $false
$resolvedDocumentKey = Resolve-DocumentKey -Value $DocumentKey
$wdFormatXMLDocument = 12   # .docx (Open XML); wdFormatXMLDocument enum value

try {
    $word              = New-Object -ComObject Word.Application
    $originalDisplayAlerts = $word.DisplayAlerts
    $originalAutomationSecurity = $word.AutomationSecurity
    $originalUpdateLinks = $word.Options.UpdateLinksAtOpen
    $settingsCaptured = $true
    $word.Visible      = $false
    $word.DisplayAlerts = 0   # wdAlertsNone
    $word.AutomationSecurity = 3   # msoAutomationSecurityForceDisable
    $word.Options.UpdateLinksAtOpen = $false

    try {
        $wordUserControl = [bool]$word.UserControl
    }
    catch {
        Write-Log "Could not resolve Word UserControl state - $($_.Exception.Message)" -Level WARN
    }

    $wordProcess = Get-WordProcess -WordApp $word
    if ($null -ne $wordProcess) {
        $wordPid = $wordProcess.Id
    }
    if ($null -ne $wordUserControl) {
        $wordProvenance = if ($wordUserControl) { "Borrowed" } else { "Owned" }
    }
    Write-Log "Word COM instance created (PID $wordPid, provenance $wordProvenance)."

    $totalFiles  = $docFiles.Count
    $currentFile = 0
    $stopAfterTimeout = $false

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
                Status = "Skipped"
                Source = $sourcePath
                Target = $targetPath
                Reason = "Target already exists"
                Error = $null
            })
            Write-Log "Skipped (exists): $sourcePath" -Level WARN
            continue
        }

        # Convert
        $document = $null
        $backupPath = $null
        $timeoutGuard = $null
        $timeoutGuardDisposed = $false
        $timeoutUnavailable = $false
        [string]$stagedPath = Join-Path $file.DirectoryName (
            ".{0}.{1}.docx" -f $file.BaseName, [Guid]::NewGuid().ToString("N")
        )

        try {
            if ($wordProvenance -eq "Owned" -and $null -eq $wordProcess) {
                $timeoutUnavailable = $true
                throw "Owned Word process could not be resolved for timeout enforcement."
            }
            if ($wordProvenance -eq "Owned" -and $null -ne $wordProcess) {
                $timeoutGuard = [HwpxWiz.ProcessTimeoutGuard]::new()
                $timeoutGuard.Arm($wordProcess, $TimeoutSeconds * 1000)
            }

            $passwordDocument = if ($resolvedDocumentKey) {
                $resolvedDocumentKey
            }
            else {
                [Type]::Missing
            }
            # Open: ConfirmConversions=$false, ReadOnly=$true, AddToRecentFiles=$false
            $document = $word.Documents.Open(
                $sourcePath,        # FileName
                $false,             # ConfirmConversions
                $true,              # ReadOnly
                $false,             # AddToRecentFiles
                $passwordDocument,  # PasswordDocument
                [Type]::Missing,    # PasswordTemplate
                $false              # Revert (preserve changes in an already-open document)
            )

            $document.SaveAs2([ref]$stagedPath, [ref]$wdFormatXMLDocument)
            Close-WordDocument -Document $document
            $document = $null

            if ($null -ne $timeoutGuard) {
                $timeoutGuard.Dispose()
                $timeoutGuardDisposed = $true
                if ($timeoutGuard.TimedOut) {
                    throw "Document conversion exceeded TimeoutSeconds=$TimeoutSeconds."
                }
            }

            if (-not (Test-Path -LiteralPath $stagedPath)) {
                throw "Word did not create staged DOCX output: $stagedPath"
            }

            if (Test-Path -LiteralPath $targetPath) {
                if (-not $Overwrite) {
                    throw "Target appeared during conversion: $targetPath"
                }
                [string]$backupPath = Join-Path $file.DirectoryName (
                    ".{0}.{1}.backup.docx" -f $file.BaseName, [Guid]::NewGuid().ToString("N")
                )
                [System.IO.File]::Replace($stagedPath, $targetPath, $backupPath)
            }
            else {
                [System.IO.File]::Move($stagedPath, $targetPath)
            }
            $stagedPath = $null
            if ($backupPath -and (Test-Path -LiteralPath $backupPath)) {
                Remove-Item -LiteralPath $backupPath -Force
                $backupPath = $null
            }

            $converted.Add([pscustomobject]@{
                Status = "Converted"
                Source = $sourcePath
                Target = $targetPath
                Reason = $null
                Error = $null
            })
            Write-Log "Converted: $sourcePath -> $targetPath"
        }
        catch {
            if ($null -ne $timeoutGuard -and -not $timeoutGuardDisposed) {
                $timeoutGuard.Dispose()
                $timeoutGuardDisposed = $true
            }
            $timedOut = $null -ne $timeoutGuard -and $timeoutGuard.TimedOut
            $failed.Add([pscustomobject]@{
                Status = "Failed"
                Source = $sourcePath
                Target = $targetPath
                Reason = if ($timedOut) {
                    "Timeout"
                }
                elseif ($timeoutUnavailable) {
                    "TimeoutUnavailable"
                }
                else {
                    $null
                }
                Error  = $_.Exception.Message
            })
            Write-Log "FAILED: $sourcePath - $($_.Exception.Message)" -Level ERROR
            if ($timedOut) {
                $stopAfterTimeout = $true
                Write-Log "Aborting remaining files after document timeout." -Level WARN
            }
        }
        finally {
            if ($null -ne $timeoutGuard -and -not $timeoutGuardDisposed) {
                $timeoutGuard.Dispose()
                $timeoutGuardDisposed = $true
            }
            Close-WordDocument -Document $document
            if ($stagedPath -and (Test-Path -LiteralPath $stagedPath)) {
                try {
                    Remove-Item -LiteralPath $stagedPath -Force
                }
                catch {
                    Write-Log "Could not remove staged DOCX '$stagedPath' - $($_.Exception.Message)" -Level WARN
                }
            }
            if ($backupPath -and (Test-Path -LiteralPath $backupPath)) {
                try {
                    Remove-Item -LiteralPath $backupPath -Force
                }
                catch {
                    Write-Log "Could not remove DOCX backup '$backupPath' - $($_.Exception.Message)" -Level WARN
                }
            }
        }

        if ($stopAfterTimeout) {
            if ($currentFile -lt $totalFiles) {
                for ($remainingIndex = $currentFile; $remainingIndex -lt $totalFiles; $remainingIndex++) {
                    $remainingFile = $docFiles[$remainingIndex]
                    $failed.Add([pscustomobject]@{
                        Status = "Failed"
                        Source = $remainingFile.FullName
                        Target = [System.IO.Path]::ChangeExtension($remainingFile.FullName, ".docx")
                        Reason = "AbortedAfterTimeout"
                        Error = "Conversion was not attempted after an earlier document timeout."
                    })
                }
            }
            break
        }
    }

    Write-Progress -Activity "Converting .doc to .docx" -Completed
}
finally {
    # Graceful shutdown
    if ($null -ne $word -and $settingsCaptured) {
        try {
            $word.DisplayAlerts = $originalDisplayAlerts
            $word.AutomationSecurity = $originalAutomationSecurity
            $word.Options.UpdateLinksAtOpen = $originalUpdateLinks
        }
        catch {
            Write-Log "Could not restore Word application settings - $($_.Exception.Message)" -Level WARN
        }
    }

    if ($null -ne $word -and $wordProvenance -eq "Owned") {
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
            catch {
                Write-Log "Could not release Word application COM object - $($_.Exception.Message)" -Level WARN
            }
        }
    }
    elseif ($null -ne $word) {
        Write-Log "Word application provenance is $wordProvenance; leaving application running."
    }

    if ($null -ne $word -and $wordProvenance -ne "Owned") {
        try {
            [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($word)
        }
        catch {
            Write-Log "Could not release Word application COM object - $($_.Exception.Message)" -Level WARN
        }
    }

    if ($null -eq $wordProcess -or $wordProvenance -ne "Owned") {
        Write-Log "Word process is not owned; skipping force-stop." -Level WARN
    }
    else {
        Stop-OwnedWordProcess -Process $wordProcess -Confirm:$false
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    $resolvedDocumentKey = $null
}

#endregion

#region Summary Report

if ($OutputFormat -eq "Table") {
    Write-Information "" -InformationAction Continue
    Write-Information "===========================================" -InformationAction Continue
    Write-Information "  CONVERSION SUMMARY" -InformationAction Continue
    Write-Information "===========================================" -InformationAction Continue
    Write-Information "" -InformationAction Continue
    Write-Information "  Converted : $($converted.Count)" -InformationAction Continue
    Write-Information "  Skipped   : $($skipped.Count)" -InformationAction Continue
    Write-Information "  Failed    : $($failed.Count)" -InformationAction Continue
    Write-Information "" -InformationAction Continue

    if ($converted.Count -gt 0) {
        Write-Information "-- Converted --" -InformationAction Continue
        $converted | Format-Table Source, Target -AutoSize | Out-Host
    }

    if ($skipped.Count -gt 0) {
        Write-Information "-- Skipped --" -InformationAction Continue
        $skipped | Format-Table Source, Target, Reason -AutoSize | Out-Host
    }

    if ($failed.Count -gt 0) {
        Write-Information "-- Failed --" -InformationAction Continue
        $failed | Format-Table Source, Error -AutoSize | Out-Host
    }
}

#endregion

#region Pipeline Output & Exit

$records = @(
    foreach ($record in $converted) { $record }
    foreach ($record in $skipped) { $record }
    foreach ($record in $failed) { $record }
)

if ($isDotSourced) {
    return $records
}

if ($OutputFormat -eq "Json") {
    ConvertTo-Json -InputObject $records -Compress
}
else {
    $records
}
if ($failed.Count -gt 0) {
    exit 1
}

exit 0

#endregion
