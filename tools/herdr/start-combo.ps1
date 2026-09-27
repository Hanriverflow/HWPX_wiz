<#
.SYNOPSIS
    Creates or reuses the HWPX_wiz PLAN, BUILD, and VERIFY workspace in Herdr.

.DESCRIPTION
    Creates a dedicated Herdr tab and stages one OMO launch command in each
    pane without pressing Enter. This prevents autonomous startup turns and
    lets the user start PLAN, BUILD, and VERIFY deliberately. Re-running the
    script reuses an existing tab instead of creating duplicate sessions.
#>

[CmdletBinding()]
param(
    [switch]$ValidateOnly,

    [ValidateNotNullOrEmpty()]
    [string]$WorkspaceId = $env:HERDR_WORKSPACE_ID,

    [ValidateNotNullOrEmpty()]
    [string]$TabLabel = "OMO Combo",

    [switch]$Focus
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

if ($env:HERDR_ENV -ne "1") {
    throw "This launcher must run inside a Herdr-managed pane (HERDR_ENV=1)."
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$requiredPaths = @(
    (Join-Path $repoRoot ".omo\omo.jsonc"),
    (Join-Path $repoRoot ".omo\prompts\plan.md"),
    (Join-Path $repoRoot ".omo\prompts\build.md"),
    (Join-Path $repoRoot ".omo\prompts\verify.md"),
    (Join-Path $repoRoot "AGENTS.md")
)

foreach ($requiredPath in $requiredPaths) {
    if (-not (Test-Path -LiteralPath $requiredPath)) {
        throw "Required combo file not found: $requiredPath"
    }
}

$configPath = Join-Path $repoRoot ".omo\omo.jsonc"
$config = Get-Content -LiteralPath $configPath -Raw | ConvertFrom-Json
if ($null -eq $config."[senpi]") {
    throw "Project OMO configuration is missing the [senpi] block: $configPath"
}

if ($ValidateOnly) {
    Write-Output "Herdr combo configuration is valid."
    return
}

# Runtime commands are only required when changing the Herdr layout.
$herdrCommand = Get-Command herdr -ErrorAction Stop
$null = Get-Command omo -ErrorAction Stop

if ([string]::IsNullOrWhiteSpace($WorkspaceId)) {
    throw "HERDR_WORKSPACE_ID is not available. Pass -WorkspaceId explicitly."
}

function Invoke-HerdrJson {
    param(
        [Parameter(Mandatory)]
        [string[]]$Arguments
    )

    $output = & $herdrCommand.Source @Arguments 2>&1
    $exitCode = $LASTEXITCODE
    $text = $output -join [Environment]::NewLine
    if ($exitCode -ne 0) {
        throw "Herdr command failed ($exitCode): herdr $($Arguments -join ' ')`n$text"
    }

    return $text | ConvertFrom-Json
}

$tabList = Invoke-HerdrJson -Arguments @("tab", "list", "--workspace", $WorkspaceId)
$existingTab = @($tabList.result.tabs) |
    Where-Object { $_.label -eq $TabLabel } |
    Select-Object -First 1

if ($null -ne $existingTab) {
    if ($Focus) {
        $null = Invoke-HerdrJson -Arguments @("tab", "focus", $existingTab.tab_id)
    }

    [pscustomobject]@{
        Created = $false
        TabId = $existingTab.tab_id
        Label = $existingTab.label
        PaneCount = $existingTab.pane_count
    }
    return
}

$createdTabId = $null
try {
    $createArguments = @(
        "tab", "create",
        "--workspace", $WorkspaceId,
        "--cwd", $repoRoot,
        "--label", $TabLabel
    )
    if ($Focus) {
        $createArguments += "--focus"
    }
    else {
        $createArguments += "--no-focus"
    }

    $tabResult = Invoke-HerdrJson -Arguments $createArguments
    $createdTabId = $tabResult.result.tab.tab_id
    $planPaneId = $tabResult.result.root_pane.pane_id

    $null = Invoke-HerdrJson -Arguments @("pane", "rename", $planPaneId, "PLAN - SOL")

    $buildResult = Invoke-HerdrJson -Arguments @(
        "pane", "split", $planPaneId,
        "--direction", "down",
        "--ratio", "0.5",
        "--cwd", $repoRoot,
        "--no-focus"
    )
    $buildPaneId = $buildResult.result.pane.pane_id
    $null = Invoke-HerdrJson -Arguments @("pane", "rename", $buildPaneId, "BUILD - LUNA")

    $verifyResult = Invoke-HerdrJson -Arguments @(
        "pane", "split", $buildPaneId,
        "--direction", "right",
        "--ratio", "0.5",
        "--cwd", $repoRoot,
        "--no-focus"
    )
    $verifyPaneId = $verifyResult.result.pane.pane_id
    $null = Invoke-HerdrJson -Arguments @("pane", "rename", $verifyPaneId, "VERIFY - SOL")

    $planCommand = 'omo --approve --model openai-codex/gpt-5.6-sol --thinking high --permission-preset workspace --name "HWPX PLAN"'
    $buildCommand = 'omo --approve --model openai-codex/gpt-5.6-luna --thinking xhigh --permission-preset workspace --name "HWPX BUILD"'
    $verifyCommand = 'omo --approve --model openai-codex/gpt-5.6-sol --thinking xhigh --permission-preset read-only --name "HWPX VERIFY"'

    $null = Invoke-HerdrJson -Arguments @("pane", "send-text", $planPaneId, $planCommand)
    $null = Invoke-HerdrJson -Arguments @("pane", "send-text", $buildPaneId, $buildCommand)
    $null = Invoke-HerdrJson -Arguments @("pane", "send-text", $verifyPaneId, $verifyCommand)

    [pscustomobject]@{
        Created = $true
        CommandsStaged = $true
        TabId = $createdTabId
        PlanPaneId = $planPaneId
        BuildPaneId = $buildPaneId
        VerifyPaneId = $verifyPaneId
    }
}
catch {
    if ($createdTabId) {
        try {
            $null = Invoke-HerdrJson -Arguments @("tab", "close", $createdTabId)
        }
        catch {
            Write-Warning "Could not close partially created Herdr tab '$createdTabId'."
        }
    }
    throw
}
