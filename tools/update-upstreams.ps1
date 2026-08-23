<#
.SYNOPSIS
    Checks pinned upstream versions and optionally applies approved updates.

.DESCRIPTION
    The default mode is read-only. Use -Apply to request an explicit
    confirmation before changing the local Kordoc installation or pulling the
    separate hwpx-skill clone. Use -Yes only for an already reviewed,
    non-interactive update.
#>

[CmdletBinding(SupportsShouldProcess)]
param(
    [ValidateSet("All", "Kordoc", "HwpxSkill")]
    [string]$Component = "All",

    [switch]$Apply,

    [switch]$Yes,

    [switch]$FailOnUpdate,

    [ValidatePattern("^\d+\.\d+\.\d+$")]
    [string]$KordocVersion,

    [string]$HwpxSkillPath = (Join-Path $HOME ".agents\skills\hwpx")
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$script:RepoRoot = Split-Path -Parent $PSScriptRoot
$script:KordocRoot = Join-Path $script:RepoRoot "tools\kordoc"
$script:KordocPackagePath = Join-Path $script:KordocRoot "package.json"
$script:VerifierPath = Join-Path $script:RepoRoot "tools\verify.ps1"
$script:VerifierTestPath = Join-Path $script:RepoRoot "tests\Verify.Tests.ps1"
$script:RequestedKordocVersion = $KordocVersion
$script:SelectedHwpxSkillPath = $HwpxSkillPath

function Resolve-CommandPath {
    param(
        [Parameter(Mandatory)]
        [string]$PreferredName,

        [Parameter(Mandatory)]
        [string]$FallbackName
    )

    $command = Get-Command $PreferredName -ErrorAction SilentlyContinue
    if ($null -eq $command) {
        $command = Get-Command $FallbackName -ErrorAction Stop
    }

    return $command.Source
}

function Invoke-CapturedCommand {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string[]]$Arguments = @()
    )

    $previousErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = @(& $FilePath @Arguments 2>&1)
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }

    if ($exitCode -ne 0) {
        $text = $output -join [Environment]::NewLine
        throw "Command failed ($exitCode): $FilePath $($Arguments -join ' ')`n$text"
    }

    return $output
}

function Get-VersionFromNpm {
    param(
        [Parameter(Mandatory)]
        [string]$NpmPath,

        [Parameter(Mandatory)]
        [string]$PackageSpec
    )

    $output = Invoke-CapturedCommand `
        -FilePath $NpmPath `
        -Arguments @("view", $PackageSpec, "version", "--json")
    $text = ($output | ForEach-Object { $_.ToString().Trim() } |
        Where-Object { $_ }) -join ""

    try {
        $parsed = $text | ConvertFrom-Json
        if ($parsed -is [Array]) {
            $text = [string]$parsed[-1]
        }
        else {
            $text = [string]$parsed
        }
    }
    catch {
        $text = $text.Trim('"')
    }

    return [version]$text.Trim()
}

function Get-KordocPlan {
    param(
        [Parameter(Mandatory)]
        [string]$NpmPath
    )

    $package = Get-Content -LiteralPath $script:KordocPackagePath -Raw |
        ConvertFrom-Json
    $current = [version]$package.dependencies.kordoc
    $latestSameMajor = Get-VersionFromNpm `
        -NpmPath $NpmPath `
        -PackageSpec "kordoc@$($current.Major)"
    $latestOverall = Get-VersionFromNpm `
        -NpmPath $NpmPath `
        -PackageSpec "kordoc"
    $target = if ($script:RequestedKordocVersion) {
        [version]$script:RequestedKordocVersion
    }
    else {
        $latestSameMajor
    }

    return [pscustomobject]@{
        Component = "Kordoc"
        Current = $current
        LatestSameMajor = $latestSameMajor
        LatestOverall = $latestOverall
        Target = $target
        HasUpdate = $target -gt $current
    }
}

function Get-OptionalCommandOutput {
    param(
        [Parameter(Mandatory)]
        [string]$FilePath,

        [Parameter()]
        [string[]]$Arguments = @()
    )

    $previousErrorAction = $ErrorActionPreference
    try {
        $ErrorActionPreference = "Continue"
        $output = @(& $FilePath @Arguments 2>$null)
        $exitCode = $LASTEXITCODE
    }
    finally {
        $ErrorActionPreference = $previousErrorAction
    }

    if ($exitCode -ne 0) {
        return $null
    }

    return $output
}

function Get-LatestHwpxTag {
    param(
        [Parameter(Mandatory)]
        [string[]]$RemoteLines
    )

    $tags = foreach ($line in $RemoteLines) {
        $parts = $line.ToString().Trim() -split "\s+"
        if ($parts.Count -eq 2 -and
            $parts[1] -match "refs/tags/v?(\d+\.\d+\.\d+)$") {
            [pscustomobject]@{
                Name = $parts[1].Substring($parts[1].LastIndexOf("/") + 1)
                Version = [version]$Matches[1]
                Commit = $parts[0]
            }
        }
    }

    return $tags | Sort-Object Version | Select-Object -Last 1
}

function Get-HwpxSkillPlan {
    if (-not (Test-Path -LiteralPath $script:SelectedHwpxSkillPath -PathType Container)) {
        return [pscustomobject]@{
            Component = "HwpxSkill"
            Installed = $false
            CurrentCommit = $null
            CurrentTag = $null
            Branch = $null
            Dirty = $false
            LatestTag = $null
            RemoteMainCommit = $null
            HasUpdate = $false
        }
    }

    $gitPath = Resolve-CommandPath -PreferredName "git.exe" -FallbackName "git"
    $remote = Get-OptionalCommandOutput `
        -FilePath $gitPath `
        -Arguments @("-C", $script:SelectedHwpxSkillPath, "remote", "get-url", "origin")
    if ($null -eq $remote -or @($remote).Count -eq 0) {
        throw "hwpx-skill clone has no readable origin remote: $script:SelectedHwpxSkillPath"
    }

    $currentCommit = (Invoke-CapturedCommand -FilePath $gitPath -Arguments @(
        "-C", $script:SelectedHwpxSkillPath, "rev-parse", "HEAD"
    ) | Select-Object -First 1).ToString().Trim()
    $branch = (Get-OptionalCommandOutput -FilePath $gitPath -Arguments @(
        "-C", $script:SelectedHwpxSkillPath, "symbolic-ref", "--short", "HEAD"
    ) | Select-Object -First 1)
    $currentTag = Get-OptionalCommandOutput -FilePath $gitPath -Arguments @(
        "-C", $script:SelectedHwpxSkillPath, "describe", "--tags", "--exact-match", "HEAD"
    )
    $dirty = @(
        Invoke-CapturedCommand -FilePath $gitPath -Arguments @(
            "-C", $script:SelectedHwpxSkillPath, "status", "--porcelain"
        )
    ).Count -gt 0
    $remoteUrl = ($remote | Select-Object -First 1).ToString().Trim()
    $remoteMain = Get-OptionalCommandOutput -FilePath $gitPath -Arguments @(
        "ls-remote", $remoteUrl, "refs/heads/main"
    )
    $remoteTags = Invoke-CapturedCommand -FilePath $gitPath -Arguments @(
        "ls-remote", "--tags", "--refs", $remoteUrl
    )
    $latestTag = Get-LatestHwpxTag -RemoteLines $remoteTags
    $remoteMainCommit = if ($remoteMain) {
        ($remoteMain | Select-Object -First 1).ToString().Split("`t")[0]
    }
    else {
        $null
    }
    $currentTagValue = if ($currentTag) {
        ($currentTag | Select-Object -First 1).ToString().Trim()
    }
    else {
        $null
    }

    return [pscustomobject]@{
        Component = "HwpxSkill"
        Installed = $true
        CurrentCommit = $currentCommit
        CurrentTag = $currentTagValue
        Branch = if ($branch) { $branch.ToString().Trim() } else { $null }
        Dirty = $dirty
        LatestTag = $latestTag
        RemoteMainCommit = $remoteMainCommit
        HasUpdate = $null -ne $remoteMainCommit -and $remoteMainCommit -ne $currentCommit
    }
}

function Write-KordocPlan {
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    Write-Output "Kordoc"
    Write-Output "  Current local : $($Plan.Current)"
    Write-Output "  Latest $($Plan.Current.Major).x : $($Plan.LatestSameMajor)"
    Write-Output "  Latest overall: $($Plan.LatestOverall)"
    Write-Output "  Proposed      : $($Plan.Target)"
    if ($Plan.LatestOverall.Major -gt $Plan.Current.Major) {
        Write-Warning "A newer Kordoc major exists. It is not selected automatically."
    }
}

function Write-HwpxPlan {
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    Write-Output "hwpx-skill"
    if (-not $Plan.Installed) {
        Write-Output "  Status        : not installed at $script:SelectedHwpxSkillPath"
        return
    }

    Write-Output "  Current tag   : $($Plan.CurrentTag)"
    Write-Output "  Current commit: $($Plan.CurrentCommit)"
    Write-Output "  Branch        : $($Plan.Branch)"
    Write-Output "  Working tree  : $(if ($Plan.Dirty) { 'DIRTY' } else { 'clean' })"
    if ($Plan.LatestTag) {
        Write-Output "  Latest tag    : $($Plan.LatestTag.Name) ($($Plan.LatestTag.Commit))"
    }
    Write-Output "  Remote main   : $($Plan.RemoteMainCommit)"
}

function Confirm-Apply {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    if ($Yes) {
        return $true
    }

    $answer = Read-Host "$Prompt [y/N]"
    return $answer -match "^(?i:y|yes)$"
}

function Set-ExactText {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [string]$Path,

        [Parameter(Mandatory)]
        [string]$OldText,

        [Parameter(Mandatory)]
        [string]$NewText
    )

    $text = [System.IO.File]::ReadAllText($Path)
    $occurrences = ([regex]::Matches(
        $text,
        [regex]::Escape($OldText)
    )).Count
    if ($occurrences -ne 1) {
        throw "Expected one pinned version in '$Path', found $occurrences."
    }

    if ($PSCmdlet.ShouldProcess($Path, "Update pinned version")) {
        $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::WriteAllText(
            $Path,
            $text.Replace($OldText, $NewText),
            $utf8NoBom
        )
    }
}

function Update-KordocBaseline {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [Parameter(Mandatory)]
        [version]$OldVersion,

        [Parameter(Mandatory)]
        [version]$NewVersion
    )

    if ($PSCmdlet.ShouldProcess(
            "tools/verify.ps1 and tests/Verify.Tests.ps1",
            "Update Kordoc version baseline"
        )) {
        Set-ExactText `
            -Path $script:VerifierPath `
            -OldText ('$script:RequiredKordocVersion = [version]"{0}"' -f $OldVersion) `
            -NewText ('$script:RequiredKordocVersion = [version]"{0}"' -f $NewVersion)
        Set-ExactText `
            -Path $script:VerifierTestPath `
            -OldText ('Should -Be ([version]"{0}")' -f $OldVersion) `
            -NewText ('Should -Be ([version]"{0}")' -f $NewVersion)
    }

    Write-Warning (
        "Review and update the pinned version in README.md, docs/UPSTREAMS.md, " +
        "docs/USAGE_GUIDE.md, and the Codex Kordoc MCP registration."
    )
}

function Invoke-KordocUpdate {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    if (-not (Confirm-Apply -Prompt "Apply Kordoc $($Plan.Current) -> $($Plan.Target)?")) {
        Write-Output "Kordoc update declined."
        return
    }

    $npmPath = Resolve-CommandPath -PreferredName "npm.cmd" -FallbackName "npm"
    if ($PSCmdlet.ShouldProcess(
            "tools/kordoc",
            "Install kordoc@$($Plan.Target) with an exact npm lock"
        )) {
        $null = Invoke-CapturedCommand -FilePath $npmPath -Arguments @(
            "install", "--prefix", $script:KordocRoot,
            "--save-exact", "kordoc@$($Plan.Target)"
        )
        $null = Invoke-CapturedCommand -FilePath $npmPath -Arguments @(
            "ci", "--prefix", $script:KordocRoot, "--ignore-scripts"
        )
        Update-KordocBaseline -OldVersion $Plan.Current -NewVersion $Plan.Target
        Write-Output "Kordoc update applied locally. Run tools/verify.ps1 before committing."
    }
}

function Invoke-HwpxSkillUpdate {
    [CmdletBinding(SupportsShouldProcess)]
    param([Parameter(Mandatory)][pscustomobject]$Plan)

    if (-not $Plan.Installed) {
        Write-Output "hwpx-skill update skipped; clone is not installed."
        return
    }
    if ($Plan.Dirty) {
        throw "hwpx-skill working tree is dirty; commit or stash it before updating."
    }
    if ($Plan.Branch -ne "main") {
        throw "hwpx-skill is on '$($Plan.Branch)'; switch to main before pull --ff-only."
    }
    if (-not (Confirm-Apply -Prompt "Pull hwpx-skill main into $($script:SelectedHwpxSkillPath)?")) {
        Write-Output "hwpx-skill update declined."
        return
    }

    $gitPath = Resolve-CommandPath -PreferredName "git.exe" -FallbackName "git"
    if ($PSCmdlet.ShouldProcess($script:SelectedHwpxSkillPath, "git pull --ff-only")) {
        $output = Invoke-CapturedCommand -FilePath $gitPath -Arguments @(
            "-C", $script:SelectedHwpxSkillPath, "pull", "--ff-only"
        )
        $output | ForEach-Object { Write-Output $_ }
        Write-Warning (
            "Record the new hwpx-skill tag/commit in docs/UPSTREAMS.md, then " +
            "run the HWPX compatibility checks."
        )
    }
}

if ($Yes -and -not $Apply) {
    throw "-Yes requires -Apply."
}

$npmPath = $null
$kordocPlan = $null
$hwpxPlan = $null
if ($Component -in @("All", "Kordoc")) {
    $npmPath = Resolve-CommandPath -PreferredName "npm.cmd" -FallbackName "npm"
    $kordocPlan = Get-KordocPlan -NpmPath $npmPath
    Write-KordocPlan -Plan $kordocPlan
}
if ($Component -in @("All", "HwpxSkill")) {
    $hwpxPlan = Get-HwpxSkillPlan
    Write-HwpxPlan -Plan $hwpxPlan
}

$updates = @(
    if ($kordocPlan -and $kordocPlan.HasUpdate) { $kordocPlan }
    if ($hwpxPlan -and $hwpxPlan.HasUpdate) { $hwpxPlan }
)

if ($updates.Count -eq 0) {
    Write-Output "No upstream updates detected for the selected components."
    exit 0
}

if (-not $Apply) {
    Write-Output ""
    Write-Output "Updates are available, but no files or external clones were changed."
    Write-Output "Review the report, then re-run with -Apply to request confirmation."
    if ($FailOnUpdate) {
        exit 10
    }
    exit 0
}

foreach ($update in $updates) {
    if ($update.Component -eq "Kordoc") {
        Invoke-KordocUpdate -Plan $update
    }
    elseif ($update.Component -eq "HwpxSkill") {
        Invoke-HwpxSkillUpdate -Plan $update
    }
}
