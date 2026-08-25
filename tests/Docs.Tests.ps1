BeforeDiscovery {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $docsRoot = Join-Path $script:RepoRoot "docs"
    $script:KordocVersionDocuments = @(
        Get-Item -LiteralPath (Join-Path $script:RepoRoot "README.md")
        Get-Item -LiteralPath (Join-Path $script:RepoRoot "HWPX_wiz_easy_guide.html")
        Get-ChildItem -LiteralPath $docsRoot -File -Filter "*.md" |
            Where-Object {
                $_.Name -notin @("MAINTENANCE_HANDOFF.md", "ROADMAP.md")
            }
    ) | ForEach-Object {
        @{
            Name = $_.Name
            FullName = $_.FullName
        }
    }
}

BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $packagePath = Join-Path $script:RepoRoot "tools\kordoc\package.json"
    $package = Get-Content -LiteralPath $packagePath -Raw | ConvertFrom-Json
    $script:ExpectedKordocVersion = [version]$package.dependencies.kordoc
}

Describe "Kordoc documentation version pins" -Tag "Static" {
    It "<Name> matches tools/kordoc/package.json" -ForEach $script:KordocVersionDocuments {
        $content = Get-Content -LiteralPath $FullName -Raw
        $versions = @(
            [regex]::Matches($content, "\b4\.\d+\.\d+\b") |
                ForEach-Object { [version]$_.Value }
        )

        foreach ($version in $versions) {
            $version | Should -Be $script:ExpectedKordocVersion
        }
    }
}

Describe "Static verification workflow" -Tag "Static" {
    It "runs the static tier on Windows" {
        $workflowPath = Join-Path $script:RepoRoot ".github\workflows\verify-static.yml"

        Test-Path -LiteralPath $workflowPath | Should -BeTrue
        $workflow = Get-Content -LiteralPath $workflowPath -Raw
        $workflow | Should -Match "windows-latest"
        $workflow | Should -Match "astral-sh/setup-uv"
        $workflow | Should -Match "npm ci --prefix tools/kordoc"
        $workflow | Should -Match "verify\.ps1 -Tier Static"
    }
}
