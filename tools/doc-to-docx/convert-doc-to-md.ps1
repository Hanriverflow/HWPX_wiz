<#
.SYNOPSIS
    Converts legacy .doc files to adjacent .docx and Markdown files.

.DESCRIPTION
    Uses Microsoft Word for .doc to .docx conversion, then official Kordoc for
    .docx to Markdown conversion. Outputs stay beside each source document.
    Existing outputs are preserved unless -Overwrite is supplied.
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
    [switch]$Overwrite
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

function Get-LegacyDocFiles {
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
    switch ($Level) {
        "ERROR" { Write-Host $entry -ForegroundColor Red }
        "WARN"  { Write-Host $entry -ForegroundColor Yellow }
        default { Write-Host $entry }
    }
}

$docConverter = Join-Path $PSScriptRoot "convert-doc-to-docx.ps1"
if (-not (Test-Path -LiteralPath $docConverter)) {
    throw "DOC converter not found: $docConverter"
}

$npx = (Get-Command npx.cmd -ErrorAction Stop).Source
$docFiles = @(Get-LegacyDocFiles -InputPath $Path -IncludeChildren:$Recurse)

if ($docFiles.Count -eq 0) {
    Write-Status "No .doc files found under '$Path'." -Level WARN
    exit 0
}

$converted = [System.Collections.Generic.List[pscustomobject]]::new()
$skipped = [System.Collections.Generic.List[pscustomobject]]::new()
$failed = [System.Collections.Generic.List[pscustomobject]]::new()

foreach ($file in $docFiles) {
    $sourcePath = $file.FullName
    $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
    $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")

    if ((Test-Path -LiteralPath $markdownPath) -and -not $Overwrite) {
        $skipped.Add([pscustomobject]@{
            Source = $sourcePath
            Reason = "Markdown already exists"
            Target = $markdownPath
        })
        Write-Status "Skipped; Markdown already exists: $markdownPath" -Level WARN
        continue
    }

    try {
        $needDocxConversion = -not (Test-Path -LiteralPath $docxPath)

        if ((Test-Path -LiteralPath $docxPath) -and -not $Overwrite) {
            $docxItem = Get-Item -LiteralPath $docxPath
            if ($docxItem.LastWriteTimeUtc -lt $file.LastWriteTimeUtc) {
                throw "Existing DOCX is older than the source DOC. Re-run with -Overwrite: $docxPath"
            }
        }

        if ($needDocxConversion -or $Overwrite) {
            $converterArgs = @(
                "-NoProfile",
                "-ExecutionPolicy", "Bypass",
                "-File", $docConverter,
                "-Path", $sourcePath
            )
            if ($Overwrite) {
                $converterArgs += "-Overwrite"
            }

            & powershell.exe @converterArgs
            if ($LASTEXITCODE -ne 0) {
                throw "DOC to DOCX converter exited with code $LASTEXITCODE"
            }
        } else {
            Write-Status "Using existing DOCX: $docxPath"
        }

        if (-not (Test-Path -LiteralPath $docxPath)) {
            throw "DOCX was not created: $docxPath"
        }

        & $npx -y kordoc@4 --silent -o $markdownPath $docxPath
        if ($LASTEXITCODE -ne 0) {
            throw "Kordoc exited with code $LASTEXITCODE"
        }
        if (-not (Test-Path -LiteralPath $markdownPath)) {
            throw "Markdown was not created: $markdownPath"
        }

        $converted.Add([pscustomobject]@{
            Source   = $sourcePath
            Docx     = $docxPath
            Markdown = $markdownPath
        })
        Write-Status "Converted: $sourcePath -> $markdownPath"
    }
    catch {
        $failed.Add([pscustomobject]@{
            Source = $sourcePath
            Error  = $_.Exception.Message
        })
        Write-Status "FAILED: $sourcePath - $($_.Exception.Message)" -Level ERROR
    }
}

Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "  DOC TO MARKDOWN SUMMARY" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "  Converted : $($converted.Count)" -ForegroundColor Green
Write-Host "  Skipped   : $($skipped.Count)" -ForegroundColor Yellow
Write-Host "  Failed    : $($failed.Count)" -ForegroundColor Red
Write-Host ""

if ($converted.Count -gt 0) {
    $converted | Format-Table Source, Docx, Markdown -AutoSize
}
if ($skipped.Count -gt 0) {
    $skipped | Format-Table Source, Target, Reason -AutoSize
}
if ($failed.Count -gt 0) {
    $failed | Format-Table Source, Error -AutoSize
    exit 1
}

exit 0
