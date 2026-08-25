<#
.SYNOPSIS
    Converts legacy .doc files to adjacent .docx and Markdown files.

.DESCRIPTION
    Uses Microsoft Word for .doc to .docx conversion, then official Kordoc for
    .docx to Markdown conversion. Outputs stay beside each source document.
    Existing outputs are preserved unless -Overwrite is supplied.

.PARAMETER Path
    Path to a .doc file or a directory containing .doc files.

.PARAMETER Recurse
    When Path is a directory, include all subdirectories.

.PARAMETER Overwrite
    Replace existing DOCX and Markdown outputs only after staged conversion succeeds.

.PARAMETER OutputFormat
    Output records as the default human-readable table or as one compressed JSON array.

.PARAMETER TimeoutSeconds
    Maximum time for one document's Word conversion. Defaults to 300 seconds.

.PARAMETER LogPath
    Optional. Path to a log file shared by the DOCX and Markdown conversion stages.

.PARAMETER DocumentKey
    Optional String or SecureString document password. When omitted,
    HWPX_WIZ_DOC_PASSWORD is inherited by the DOC conversion stage.

.PARAMETER KordocCliPath
    Optional. Path to an alternate Kordoc cli.js entrypoint.

.EXAMPLE
    .\convert-doc-to-md.ps1 -Path "C:\Reports\Q1.doc" -OutputFormat Json
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [ValidateScript({
        if (-not (Test-Path -LiteralPath $_)) {
            throw "Path does not exist: $_"
        }
        return $true
    })]
    [string]$Path,

    [switch]$Recurse,
    [switch]$Overwrite,

    [ValidateSet("Table", "Json")]
    [string]$OutputFormat = "Table",

    [ValidateRange(1, 86400)]
    [int]$TimeoutSeconds = 300,

    [Parameter()]
    [string]$LogPath,

    [Parameter()]
    [Alias("Password")]
    [ValidateScript({
        if ($_ -is [string] -or $_ -is [System.Security.SecureString]) {
            return $true
        }
        throw "DocumentKey must be a String or SecureString."
    })]
    [object]$DocumentKey,

    [Parameter()]
    [string]$KordocCliPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$isDotSourced = $MyInvocation.InvocationName -eq "."
$script:ConversionLogPath = $null

function Get-LegacyDocFile {
    param(
        [Parameter(Mandatory)]
        [string]$InputPath,

        [switch]$IncludeChildren
    )

    $item = Get-Item -LiteralPath (Resolve-Path -LiteralPath $InputPath)

    if (-not $item.PSIsContainer) {
        if ($item.Extension -ine ".doc") {
            throw "Input file must have a .doc extension: $($item.FullName)"
        }
        return @(, $item)
    }

    $params = @{
        LiteralPath = $item.FullName
        File        = $true
        Filter      = "*.doc"
        Recurse     = $IncludeChildren.IsPresent
    }

    return @(
        Get-ChildItem @params |
            Where-Object { $_.Extension -ieq ".doc" } |
            Sort-Object FullName
    )
}

function Write-Status {
    param(
        [Parameter(Mandatory)]
        [string]$Message,

        [ValidateSet("INFO", "WARN", "ERROR")]
        [string]$Level = "INFO"
    )

    $entry = "[{0}] [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
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

$docConverter = Join-Path $PSScriptRoot "convert-doc-to-docx.ps1"
if (-not (Test-Path -LiteralPath $docConverter)) {
    throw "DOC converter not found: $docConverter"
}

$docFiles = @(Get-LegacyDocFile -InputPath $Path -IncludeChildren:$Recurse)

if ($LogPath) {
    if ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($LogPath)) {
        $record = [pscustomobject]@{
            Status = "Failed"
            Source = $Path
            Target = $null
            Reason = "LogPath must be literal"
            Error = "LogPath must not contain wildcard characters: $LogPath"
            Docx = $null
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
            Docx = $null
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
            [System.IO.Path]::ChangeExtension($docFile.FullName, ".md")
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
            Docx = $null
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

if ($docFiles.Count -eq 0) {
    Write-Status "No .doc files found under '$Path'." -Level WARN
    if ($isDotSourced) {
        return @()
    }
    if ($OutputFormat -eq "Json") {
        ConvertTo-Json -InputObject @() -Compress
    }
    exit 0
}

function Resolve-KordocCli {
    param(
        [Parameter()]
        [string]$OverridePath
    )

    if ($OverridePath) {
        $resolvedOverride = Resolve-Path -LiteralPath $OverridePath -ErrorAction Stop
        return $resolvedOverride.Path
    }

    $localCli = Join-Path (Split-Path -Parent (Split-Path -Parent $PSScriptRoot)) `
        "tools/kordoc/node_modules/kordoc/dist/cli.js"
    if (-not (Test-Path -LiteralPath $localCli -PathType Leaf)) {
        throw "Local Kordoc CLI is missing: $localCli. Run 'npm ci --prefix tools/kordoc'."
    }

    return $localCli
}

$node = (Get-Command node.exe -ErrorAction Stop).Source
$kordocCli = Resolve-KordocCli -OverridePath $KordocCliPath
$converted = [System.Collections.Generic.List[pscustomobject]]::new()
$skipped = [System.Collections.Generic.List[pscustomobject]]::new()
$failed = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($file in $docFiles) {
    $sourcePath = $file.FullName
    $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
    $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
    $stagedMarkdownPath = $null
    $backupMarkdownPath = $null
    $docxRollbackPath = $null
    $createdDocxThisRun = $false
    $fileCompleted = $false
    $failureReason = $null

    try {
        $hadExistingDocx = Test-Path -LiteralPath $docxPath
        $needDocxConversion = -not $hadExistingDocx

        if ((Test-Path -LiteralPath $docxPath) -and -not $Overwrite) {
            $docxItem = Get-Item -LiteralPath $docxPath
            if ($docxItem.LastWriteTimeUtc -lt $file.LastWriteTimeUtc) {
                throw "Existing DOCX is older than the source DOC. Re-run with -Overwrite: $docxPath"
            }
        }

        if ($needDocxConversion -or $Overwrite) {
            if ($hadExistingDocx) {
                [string]$docxRollbackPath = Join-Path $file.DirectoryName (
                    ".{0}.{1}.rollback.docx" -f $file.BaseName, [Guid]::NewGuid().ToString("N")
                )
                Copy-Item -LiteralPath $docxPath -Destination $docxRollbackPath -Force
            }

            $converterArgs = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", $docConverter,
                "-Path", $sourcePath
            )
            if ($Overwrite) {
                $converterArgs += "-Overwrite"
            }
            if ($LogPath) {
                $converterArgs += @("-LogPath", $LogPath)
            }
            $converterArgs += @("-TimeoutSeconds", $TimeoutSeconds)
            $converterArgs += @("-OutputFormat", "Json")

            $documentKeyWasBound = $PSBoundParameters.ContainsKey("DocumentKey")
            $hadPreviousEnvironmentPassword = Test-Path Env:HWPX_WIZ_DOC_PASSWORD
            $previousEnvironmentPassword = $env:HWPX_WIZ_DOC_PASSWORD
            try {
                if ($documentKeyWasBound) {
                    $resolvedChildDocumentKey = Resolve-DocumentKey -Value $DocumentKey
                    if ($resolvedChildDocumentKey) {
                        $env:HWPX_WIZ_DOC_PASSWORD = $resolvedChildDocumentKey
                    }
                    else {
                        Remove-Item Env:HWPX_WIZ_DOC_PASSWORD -ErrorAction SilentlyContinue
                    }
                }
                $docConverterOutput = @(& powershell.exe @converterArgs)
                $docConverterExitCode = $LASTEXITCODE
            }
            finally {
                if ($hadPreviousEnvironmentPassword) {
                    $env:HWPX_WIZ_DOC_PASSWORD = $previousEnvironmentPassword
                }
                else {
                    Remove-Item Env:HWPX_WIZ_DOC_PASSWORD -ErrorAction SilentlyContinue
                }
                $resolvedChildDocumentKey = $null
            }
            $docConverterRecords = @()
            if ($docConverterOutput.Count -gt 0) {
                $docConverterJson = $docConverterOutput -join [Environment]::NewLine
                $parsedDocConverterRecords = $docConverterJson | ConvertFrom-Json
                $docConverterRecords = @(
                    foreach ($parsedRecord in $parsedDocConverterRecords) {
                        $parsedRecord
                    }
                )
            }
            if ($docConverterExitCode -ne 0) {
                $timedOutRecords = @(
                    $docConverterRecords |
                        Where-Object {
                            $null -ne $_ -and
                            $null -ne $_.PSObject.Properties["Reason"] -and
                            $_.Reason -eq "Timeout"
                        }
                )
                if ($timedOutRecords.Count -gt 0) {
                    $failureReason = "Timeout"
                }
                throw (
                    "DOC to DOCX converter exited with code ${docConverterExitCode}: " +
                    ($docConverterOutput -join [Environment]::NewLine)
                )
            }
            $createdDocxThisRun = -not $hadExistingDocx
        } else {
            Write-Status "Using existing DOCX: $docxPath"
        }

        if (-not (Test-Path -LiteralPath $docxPath)) {
            throw "DOCX was not created: $docxPath"
        }

        $docxItem = Get-Item -LiteralPath $docxPath
        if ((Test-Path -LiteralPath $markdownPath) -and -not $Overwrite) {
            $markdownItem = Get-Item -LiteralPath $markdownPath
            if (
                $markdownItem.LastWriteTimeUtc -lt $file.LastWriteTimeUtc -or
                $markdownItem.LastWriteTimeUtc -lt $docxItem.LastWriteTimeUtc
            ) {
                throw "Existing Markdown is older than its DOC or DOCX input. Re-run with -Overwrite: $markdownPath"
            }

            $skipped.Add([pscustomobject]@{
                Status = "Skipped"
                Source = $sourcePath
                Target = $markdownPath
                Reason = "Markdown is up to date"
                Error = $null
                Docx = $docxPath
            })
            Write-Status "Skipped; Markdown is up to date: $markdownPath" -Level WARN
            $fileCompleted = $true
            continue
        }

        [string]$stagedMarkdownPath = Join-Path $file.DirectoryName (
            ".{0}.{1}.md" -f $file.BaseName, [Guid]::NewGuid().ToString("N")
        )
        $hadKordocEnvironmentPassword = Test-Path Env:HWPX_WIZ_DOC_PASSWORD
        $kordocEnvironmentPassword = $env:HWPX_WIZ_DOC_PASSWORD
        try {
            Remove-Item Env:HWPX_WIZ_DOC_PASSWORD -ErrorAction SilentlyContinue
            & $node $kordocCli --silent -o $stagedMarkdownPath $docxPath
            $kordocExitCode = $LASTEXITCODE
        }
        finally {
            if ($hadKordocEnvironmentPassword) {
                $env:HWPX_WIZ_DOC_PASSWORD = $kordocEnvironmentPassword
            }
            else {
                Remove-Item Env:HWPX_WIZ_DOC_PASSWORD -ErrorAction SilentlyContinue
            }
        }
        if ($kordocExitCode -ne 0) {
            throw "Kordoc exited with code $kordocExitCode"
        }
        if (-not (Test-Path -LiteralPath $stagedMarkdownPath)) {
            throw "Kordoc did not create staged Markdown output: $stagedMarkdownPath"
        }

        if (Test-Path -LiteralPath $markdownPath) {
            if (-not $Overwrite) {
                throw "Markdown target appeared during conversion: $markdownPath"
            }
            [string]$backupMarkdownPath = Join-Path $file.DirectoryName (
                ".{0}.{1}.backup.md" -f $file.BaseName, [Guid]::NewGuid().ToString("N")
            )
            [System.IO.File]::Replace($stagedMarkdownPath, $markdownPath, $backupMarkdownPath)
        }
        else {
            [System.IO.File]::Move($stagedMarkdownPath, $markdownPath)
        }
        $stagedMarkdownPath = $null
        $fileCompleted = $true
        if ($backupMarkdownPath -and (Test-Path -LiteralPath $backupMarkdownPath)) {
            Remove-Item -LiteralPath $backupMarkdownPath -Force
            $backupMarkdownPath = $null
        }

        $converted.Add([pscustomobject]@{
            Status = "Converted"
            Source = $sourcePath
            Target = $markdownPath
            Reason = $null
            Error = $null
            Docx = $docxPath
        })
        Write-Status "Converted: $sourcePath -> $markdownPath"
    }
    catch {
        $failed.Add([pscustomobject]@{
            Status = "Failed"
            Source = $sourcePath
            Target = $markdownPath
            Reason = $failureReason
            Error = $_.Exception.Message
            Docx = $docxPath
        })
        Write-Status "FAILED: $sourcePath - $($_.Exception.Message)" -Level ERROR
    }
    finally {
        if ($fileCompleted) {
            if ($docxRollbackPath -and (Test-Path -LiteralPath $docxRollbackPath)) {
                try {
                    Remove-Item -LiteralPath $docxRollbackPath -Force
                    $docxRollbackPath = $null
                }
                catch {
                    Write-Status "Could not remove DOCX rollback '$docxRollbackPath' - $($_.Exception.Message)" -Level WARN
                }
            }
        }
        else {
            if ($docxRollbackPath -and (Test-Path -LiteralPath $docxRollbackPath)) {
                try {
                    [string]$restoreScratchPath = Join-Path $file.DirectoryName (
                        ".{0}.{1}.restore.docx" -f $file.BaseName, [Guid]::NewGuid().ToString("N")
                    )
                    [System.IO.File]::Replace($docxRollbackPath, $docxPath, $restoreScratchPath)
                    $docxRollbackPath = $null
                    Remove-Item -LiteralPath $restoreScratchPath -Force
                }
                catch {
                    Write-Status "Could not restore previous DOCX from '$docxRollbackPath' - $($_.Exception.Message)" -Level ERROR
                }
            }
            elseif ($createdDocxThisRun -and (Test-Path -LiteralPath $docxPath)) {
                try {
                    Remove-Item -LiteralPath $docxPath -Force
                }
                catch {
                    Write-Status "Could not remove newly created DOCX '$docxPath' - $($_.Exception.Message)" -Level WARN
                }
            }
        }

        if ($stagedMarkdownPath -and (Test-Path -LiteralPath $stagedMarkdownPath)) {
            try {
                Remove-Item -LiteralPath $stagedMarkdownPath -Force
            }
            catch {
                Write-Status "Could not remove staged Markdown '$stagedMarkdownPath' - $($_.Exception.Message)" -Level WARN
            }
        }
        if ($backupMarkdownPath -and (Test-Path -LiteralPath $backupMarkdownPath)) {
            try {
                Remove-Item -LiteralPath $backupMarkdownPath -Force
            }
            catch {
                Write-Status "Could not remove Markdown backup '$backupMarkdownPath' - $($_.Exception.Message)" -Level WARN
            }
        }
    }
}

if ($OutputFormat -eq "Table") {
    Write-Information "" -InformationAction Continue
    Write-Information "===========================================" -InformationAction Continue
    Write-Information "  DOC TO MARKDOWN SUMMARY" -InformationAction Continue
    Write-Information "===========================================" -InformationAction Continue
    Write-Information "  Converted : $($converted.Count)" -InformationAction Continue
    Write-Information "  Skipped   : $($skipped.Count)" -InformationAction Continue
    Write-Information "  Failed    : $($failed.Count)" -InformationAction Continue
    Write-Information "" -InformationAction Continue

    if ($converted.Count -gt 0) {
        $converted | Format-Table Source, Docx, Target -AutoSize | Out-Host
    }
    if ($skipped.Count -gt 0) {
        $skipped | Format-Table Source, Target, Reason -AutoSize | Out-Host
    }
    if ($failed.Count -gt 0) {
        $failed | Format-Table Source, Error -AutoSize | Out-Host
    }
}

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
