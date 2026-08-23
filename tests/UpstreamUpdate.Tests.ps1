BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:UpdaterPath = Join-Path $script:RepoRoot "tools/update-upstreams.ps1"
    $script:PowerShellPath = (Get-Command powershell.exe -ErrorAction Stop).Source
}

Describe "Upstream update workflow" {
    It "has a read-only-by-default updater" {
        Test-Path -LiteralPath $script:UpdaterPath | Should -BeTrue
        $source = Get-Content -LiteralPath $script:UpdaterPath -Raw

        $source | Should -Match '\[switch\]\$Apply'
        $source | Should -Match "Review the report"
        $source | Should -Match "No files or external clones were changed"
    }

    It "requires Apply before accepting non-interactive approval" {
        $output = & $script:PowerShellPath -NoProfile `
            -ExecutionPolicy Bypass -File $script:UpdaterPath -Yes 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "-Yes requires -Apply"
    }

    It "reports an absent skill clone without modifying the repository" {
        $missingPath = Join-Path $TestDrive "missing-hwpx-skill"
        $output = & $script:PowerShellPath -NoProfile `
            -ExecutionPolicy Bypass -File $script:UpdaterPath `
            -Component HwpxSkill -HwpxSkillPath $missingPath 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        ($output -join [Environment]::NewLine) |
            Should -Match "not installed"
        Test-Path -LiteralPath $missingPath | Should -BeFalse
    }

    It "uses fast-forward-only skill updates and explicit version pins" {
        $source = Get-Content -LiteralPath $script:UpdaterPath -Raw

        $source | Should -Match "pull.*--ff-only"
        $source | Should -Match '"install", "--prefix"'
        $source | Should -Match "save-exact"
        $source | Should -Match "ShouldProcess"
    }
}
