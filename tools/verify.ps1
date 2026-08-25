<#
.SYNOPSIS
    Runs the repository's prerequisite, test, analyzer, and lock checks.

.DESCRIPTION
    This is the canonical local verification entry point. It never installs
    missing tools. Use -PrerequisiteCheckOnly to validate the machine without
    recursively invoking the Pester suite.
#>

[CmdletBinding()]
param(
    [switch]$PrerequisiteCheckOnly,

    [ValidateSet("Full", "Static")]
    [string]$Tier = "Full",

    [switch]$IncludeHwpx,

    [string[]]$SimulateMissing = @()
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:VerifierRoot = Split-Path -Parent $PSScriptRoot
$script:MinimumPester = [version]"6.1.0"
$script:MinimumAnalyzer = [version]"1.25.0"

function Get-InstalledModuleVersion {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    $module = Get-Module -ListAvailable -Name $Name |
        Sort-Object Version -Descending |
        Select-Object -First 1
    if ($null -eq $module) {
        return $null
    }

    return [version]$module.Version
}

function Get-RepositoryPrerequisiteFailure {
    param(
        [Parameter()]
        [string[]]$SimulatedMissing = @(),

        [ValidateSet("Full", "Static")]
        [string]$VerificationTier = "Full"
    )

    $missing = @(
        $SimulatedMissing |
            ForEach-Object { $_ -split "," } |
            ForEach-Object { $_.Trim() }
    )
    $failures = [System.Collections.Generic.List[string]]::new()

    $pesterVersion = Get-InstalledModuleVersion -Name "Pester"
    if ($missing -contains "Pester" -or
        $null -eq $pesterVersion -or $pesterVersion -lt $script:MinimumPester) {
        $actual = if ($null -eq $pesterVersion) { "not installed" } else { $pesterVersion }
        $failures.Add(
            "Pester >= $($script:MinimumPester) is required (found $actual). " +
            "Install-Module Pester -MinimumVersion $($script:MinimumPester) -Scope CurrentUser"
        )
    }

    $analyzerVersion = Get-InstalledModuleVersion -Name "PSScriptAnalyzer"
    if ($missing -contains "PSScriptAnalyzer" -or
        $null -eq $analyzerVersion -or $analyzerVersion -lt $script:MinimumAnalyzer) {
        $actual = if ($null -eq $analyzerVersion) { "not installed" } else { $analyzerVersion }
        $failures.Add(
            "PSScriptAnalyzer >= $($script:MinimumAnalyzer) is required (found $actual). " +
            "Install-Module PSScriptAnalyzer -MinimumVersion $($script:MinimumAnalyzer) -Scope CurrentUser"
        )
    }

    $uv = Get-Command uv -ErrorAction SilentlyContinue
    if ($missing -contains "uv" -or $null -eq $uv) {
        $failures.Add("uv is required on PATH; install uv and reopen PowerShell so the uv command is available.")
    }

    $node = Get-Command node.exe -ErrorAction SilentlyContinue
    if ($missing -contains "Node" -or $null -eq $node) {
        $failures.Add("Node.js is required on PATH for the local Kordoc CLI.")
    }

    $npm = Get-Command npm.cmd -ErrorAction SilentlyContinue
    if ($null -eq $npm) {
        $npm = Get-Command npm -ErrorAction SilentlyContinue
    }
    if ($missing -contains "npm" -or $null -eq $npm) {
        $failures.Add("npm is required on PATH to validate the locked local Kordoc installation.")
    }

    if ($VerificationTier -eq "Full") {
        $wordComType = [type]::GetTypeFromProgID("Word.Application")
        if ($missing -contains "Office" -or $null -eq $wordComType) {
            $failures.Add(
                "Office integration tier requires registered Word.Application COM automation; " +
                "run on a supported machine with Microsoft Word installed."
            )
        }
    }

    $kordocRoot = Join-Path $script:VerifierRoot "tools/kordoc"
    $kordocPackagePath = Join-Path $kordocRoot "package.json"
    $kordocLockPath = Join-Path $kordocRoot "package-lock.json"
    $kordocCliPath = Join-Path $kordocRoot "node_modules/kordoc/dist/cli.js"
    foreach ($requiredKordocPath in @($kordocPackagePath, $kordocLockPath, $kordocCliPath)) {
        if ($missing -contains "Kordoc" -or
            -not (Test-Path -LiteralPath $requiredKordocPath -PathType Leaf)) {
            $failures.Add(
                "Local Kordoc installation is incomplete: missing $requiredKordocPath. " +
                "Run 'npm ci --prefix .\tools\kordoc'."
            )
        }
    }

    $lockPath = Join-Path $script:VerifierRoot "uv.lock"
    $projectPath = Join-Path $script:VerifierRoot "pyproject.toml"
    $pythonPath = Join-Path $script:VerifierRoot ".venv\Scripts\python.exe"
    foreach ($requiredPath in @($lockPath, $projectPath, $pythonPath)) {
        if ($missing -contains "Python" -or
            -not (Test-Path -LiteralPath $requiredPath -PathType Leaf)) {
            $failures.Add("Locked Python environment is incomplete: missing $requiredPath. Run 'uv sync --locked'.")
        }
    }

    if ($missing -contains "HwpxPython") {
        $failures.Add("HWPX skill Python imports are unavailable: hwpx, lxml, win32com.")
    }
    elseif (
        $null -ne $uv -and
        (Test-Path -LiteralPath $pythonPath -PathType Leaf)
    ) {
        $importExitCode = Invoke-VerifyExternalCommand `
            -FilePath $uv.Source `
            -Arguments @(
                "run", "--project", $script:VerifierRoot,
                "python", "-c", "import hwpx, lxml, win32com"
            )
        if ($importExitCode -ne 0) {
            $failures.Add(
                "HWPX skill Python imports failed: hwpx, lxml, win32com."
            )
        }
    }

    return @($failures)
}

function Invoke-VerifyExternalCommand {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string[]]$Arguments = @()
    )

    $previousErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $childOutput = & $FilePath @Arguments 2>&1
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }

    foreach ($line in $childOutput) {
        Write-Information $line -InformationAction Continue
    }
    return $exitCode
}

function Invoke-HwpxValidationSmoke {
    param(
        [Parameter(Mandatory)]
        [string]$SkillRoot,

        [Parameter(Mandatory)]
        [string]$UvPath
    )

    $validatorPath = Join-Path $SkillRoot "scripts\validate.py"
    $templatePath = Join-Path $SkillRoot "assets\report-template.hwpx"
    if (
        -not (Test-Path -LiteralPath $validatorPath -PathType Leaf) -or
        -not (Test-Path -LiteralPath $templatePath -PathType Leaf)
    ) {
        Write-Information "Optional HWPX validation skipped; hwpx skill assets were not found." `
            -InformationAction Continue
        return 0
    }

    $baseArguments = @(
        "run", "--project", $script:VerifierRoot,
        "python", $validatorPath, $templatePath
    )
    $layoutExitCode = Invoke-VerifyExternalCommand `
        -FilePath $UvPath `
        -Arguments ($baseArguments + "--layout")
    if ($layoutExitCode -ne 0) {
        Write-Warning "Optional HWPX layout validation failed with exit code $layoutExitCode."
        return 0
    }

    $hancomType = [type]::GetTypeFromProgID("HWPFrame.HwpObject")
    if ($null -eq $hancomType) {
        Write-Information "Optional HWPX Hancom validation skipped; COM is not registered." `
            -InformationAction Continue
        return 0
    }

    $hancomExitCode = Invoke-VerifyExternalCommand `
        -FilePath $UvPath `
        -Arguments ($baseArguments + "--hancom")
    if ($hancomExitCode -ne 0) {
        Write-Warning "Optional HWPX Hancom validation failed with exit code $hancomExitCode."
        return 0
    }

    Write-Information "Optional HWPX validation smoke passed." -InformationAction Continue
    return 0
}

function Get-LockedKordocVersion {
    param(
        [Parameter(Mandatory)]
        [string]$LockPath,

        [Parameter(Mandatory)]
        [string]$NodePath
    )

    $nodeScript = @'
const fs = require("fs");
const lock = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
process.stdout.write(lock.packages[""].dependencies.kordoc);
'@
    $versionOutput = $nodeScript | & $NodePath - $LockPath 2>&1
    $exitCode = $LASTEXITCODE
    if ($exitCode -ne 0) {
        throw "Could not read locked Kordoc version from '$LockPath'."
    }

    return [version](($versionOutput | Select-Object -First 1).ToString().Trim())
}

function Get-KordocMcpConfigurationWarning {
    param(
        [Parameter(Mandatory)]
        [string]$ConfigPath,

        [Parameter(Mandatory)]
        [string]$ExpectedCliPath
    )

    if (-not (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
        return "Codex config was not found at '$ConfigPath'; register the local Kordoc CLI."
    }

    $config = [System.IO.File]::ReadAllText($ConfigPath)
    $match = [regex]::Match(
        $config,
        "(?ms)^\[mcp_servers\.kordoc\]\s*(?<body>.*?)(?=^\[|\z)"
    )
    if (-not $match.Success) {
        return "Codex Kordoc MCP is not registered; register the local Kordoc CLI."
    }

    $block = $match.Groups["body"].Value.Replace("\", "/").ToLowerInvariant()
    $expected = $ExpectedCliPath.Replace("\", "/").ToLowerInvariant()
    if (
        $block -notmatch 'command\s*=\s*"node\.exe"' -or
        -not $block.Contains($expected) -or
        $block -notmatch 'args\s*=.*"mcp"'
    ) {
        return "Codex Kordoc MCP does not use the repository-local Kordoc CLI: $ExpectedCliPath"
    }

    return $null
}

function Invoke-RepositoryVerification {
    param(
        [switch]$PrerequisiteOnly,

        [ValidateSet("Full", "Static")]
        [string]$VerificationTier = "Full",

        [switch]$IncludeHwpxSmoke,

        [Parameter()]
        [string[]]$SimulatedMissing = @()
    )

    $originalLocation = (Get-Location).Path
    try {
        Set-Location -LiteralPath $script:VerifierRoot
        $failures = @(
            Get-RepositoryPrerequisiteFailure `
                -SimulatedMissing $SimulatedMissing `
                -VerificationTier $VerificationTier
        )
        if ($failures.Count -gt 0) {
            foreach ($failure in $failures) {
                Write-Information "ERROR: $failure" -InformationAction Continue
            }
            return 1
        }

        if ($PrerequisiteOnly) {
            Write-Information "Prerequisite check passed." -InformationAction Continue
            return 0
        }

        $npm = (Get-Command npm.cmd -ErrorAction SilentlyContinue)
        if ($null -eq $npm) {
            $npm = Get-Command npm -ErrorAction Stop
        }
        $kordocRoot = Join-Path $script:VerifierRoot "tools/kordoc"
        $npmExitCode = Invoke-VerifyExternalCommand `
            -FilePath $npm.Source `
            -Arguments @("ci", "--dry-run", "--ignore-scripts", "--prefix", $kordocRoot)
        if ($npmExitCode -ne 0) {
            Write-Error "Locked Kordoc npm dependency gate failed with exit code $npmExitCode."
            return 1
        }

        $packagePath = Join-Path $kordocRoot "package.json"
        $lockPath = Join-Path $kordocRoot "package-lock.json"
        $package = Get-Content -LiteralPath $packagePath -Raw | ConvertFrom-Json
        $node = (Get-Command node.exe -ErrorAction Stop).Source
        $declaredKordoc = [version]$package.dependencies.kordoc
        $lockedKordoc = Get-LockedKordocVersion -LockPath $lockPath -NodePath $node
        if ($lockedKordoc -ne $declaredKordoc) {
            Write-Error (
                "Kordoc version gate failed: package.json=$declaredKordoc " +
                "package-lock.json=$lockedKordoc."
            )
            return 1
        }

        $kordocCliPath = Join-Path $kordocRoot "node_modules/kordoc/dist/cli.js"
        $kordocVersionOutput = & $node $kordocCliPath --version 2>&1
        $kordocVersionExitCode = $LASTEXITCODE
        if ($kordocVersionExitCode -ne 0) {
            Write-Error "Local Kordoc CLI version check failed with exit code $kordocVersionExitCode."
            return 1
        }
        $kordocVersion = [version](($kordocVersionOutput | Select-Object -First 1).ToString().Trim())
        if ($kordocVersion -ne $declaredKordoc) {
            Write-Error (
                "Local Kordoc CLI version gate failed: found $kordocVersion " +
                "package.json declares $declaredKordoc."
            )
            return 1
        }
        Write-Information (
            "Kordoc dependency gate passed: locked $lockedKordoc, CLI $kordocVersion."
        ) -InformationAction Continue
        $codexConfigPath = Join-Path $HOME ".codex\config.toml"
        $mcpWarning = Get-KordocMcpConfigurationWarning `
            -ConfigPath $codexConfigPath `
            -ExpectedCliPath $kordocCliPath
        if ($mcpWarning) {
            Write-Warning $mcpWarning
        }
        if ($IncludeHwpxSmoke) {
            $uvCommand = (Get-Command uv -ErrorAction Stop).Source
            $hwpxSkillRoot = Join-Path $HOME ".agents\skills\hwpx"
            $null = Invoke-HwpxValidationSmoke `
                -SkillRoot $hwpxSkillRoot `
                -UvPath $uvCommand
        }
        if ($VerificationTier -eq "Full") {
            Write-Information (
                "Office integration tier: required and available (Word.Application COM registered)."
            ) -InformationAction Continue
        }
        else {
            Write-Information "Static verification tier: Office integration skipped." `
                -InformationAction Continue
        }

        Import-Module Pester -MinimumVersion $script:MinimumPester -Force
        Import-Module PSScriptAnalyzer -MinimumVersion $script:MinimumAnalyzer -Force

        Write-Information "Running Pester tests..." -InformationAction Continue
        $previousErrorAction = $ErrorActionPreference
        try {
            $ErrorActionPreference = "Continue"
            $pesterParameters = @{
                Path = Join-Path $script:VerifierRoot "tests"
                Output = "Detailed"
                PassThru = $true
            }
            if ($VerificationTier -eq "Static") {
                $pesterParameters.TagFilter = @("Static")
            }
            $pesterResult = Invoke-Pester @pesterParameters
        }
        finally {
            $ErrorActionPreference = $previousErrorAction
        }
        if ($null -eq $pesterResult -or $pesterResult.FailedCount -gt 0) {
            Write-Error "Pester reported failed tests."
            return 1
        }

        Write-Information "Running PSScriptAnalyzer on tools..." -InformationAction Continue
        $toolFindings = @(Invoke-ScriptAnalyzer -Path (Join-Path $script:VerifierRoot "tools") -Recurse)
        if ($toolFindings.Count -gt 0) {
            $toolFindings | Write-Error
            return 1
        }

        Write-Information "Running PSScriptAnalyzer on tests..." -InformationAction Continue
        $testFindings = @(Invoke-ScriptAnalyzer -Path (Join-Path $script:VerifierRoot "tests") -Recurse)
        if ($testFindings.Count -gt 0) {
            $testFindings | Write-Error
            return 1
        }

        Write-Information "Running uv lock --check..." -InformationAction Continue
        $uv = (Get-Command uv -ErrorAction Stop).Source
        $uvExitCode = Invoke-VerifyExternalCommand -FilePath $uv -Arguments @("lock", "--check")
        if ($uvExitCode -ne 0) {
            Write-Error "uv lock --check failed with exit code $uvExitCode."
            return 1
        }

        Write-Information "Repository verification passed." -InformationAction Continue
        return 0
    }
    catch {
        Write-Error $_
        return 1
    }
    finally {
        Set-Location -LiteralPath $originalLocation
    }
}

$isDotSourced = $MyInvocation.InvocationName -eq "."
$exitCode = Invoke-RepositoryVerification `
    -PrerequisiteOnly:$PrerequisiteCheckOnly `
    -VerificationTier $Tier `
    -IncludeHwpxSmoke:$IncludeHwpx `
    -SimulatedMissing $SimulateMissing
if ($isDotSourced) {
    return $exitCode
}

exit $exitCode
