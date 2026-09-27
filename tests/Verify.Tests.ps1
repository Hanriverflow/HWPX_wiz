BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:VerifierPath = Join-Path $script:RepoRoot "tools\verify.ps1"
    $script:PowerShellPath = (Get-Command powershell.exe -ErrorAction Stop).Source
}

Describe "Repository verifier" -Tag "Static" {
    It "has a verifier entry point" {
        Test-Path -LiteralPath $script:VerifierPath | Should -BeTrue
    }

    It "passes its prerequisite-only seam without running the suite" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -Tier Static -PrerequisiteCheckOnly 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        ($output -join [Environment]::NewLine) |
            Should -Match "Prerequisite check passed"
    }

    It "reports missing modules without installing them" {
        $isolatedModulePath = Join-Path $TestDrive "isolated-modules"
        [void](New-Item -ItemType Directory -Path $isolatedModulePath -Force)
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly `
            -SimulateMissing Pester,PSScriptAnalyzer 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        $text = $output -join [Environment]::NewLine
        $text | Should -Match "Install-Module Pester -MinimumVersion 6.1.0 -Scope CurrentUser"
        $text | Should -Match "Install-Module PSScriptAnalyzer -MinimumVersion 1.25.0 -Scope CurrentUser"
        Test-Path -LiteralPath (Join-Path $isolatedModulePath "Pester") |
            Should -BeFalse
    }

    It "reports a missing uv executable without installing it" {
        $previousPath = $env:PATH
        try {
            $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
                -File $script:VerifierPath -PrerequisiteCheckOnly `
                -SimulateMissing uv 6>&1 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $env:PATH = $previousPath
        }

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) | Should -Match "uv"
    }

    It "reports a missing npm executable without installing it" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly `
            -SimulateMissing npm 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) | Should -Match "npm"
    }

    It "reports unavailable Word COM automation explicitly" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly `
            -SimulateMissing Office 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "Word.Application|Office integration"
    }

    It "lets the static tier run without Word COM" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -Tier Static -PrerequisiteCheckOnly `
            -SimulateMissing Office 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        ($output -join [Environment]::NewLine) |
            Should -Match "Prerequisite check passed"
    }

    It "reports missing Node or local Kordoc installation guidance" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly `
            -SimulateMissing Node,Kordoc 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "npm ci --prefix \.\\tools\\kordoc"
    }

    It "reports missing HWPX skill Python imports" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly `
            -SimulateMissing HwpxPython 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "hwpx.*lxml.*win32com"
    }

    It "reads the locked Kordoc version from the npm lockfile" {
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null
        $nodePath = (Get-Command node.exe -ErrorAction Stop).Source
        $lockPath = Join-Path $script:RepoRoot "tools\kordoc\package-lock.json"
        $packagePath = Join-Path $script:RepoRoot "tools\kordoc\package.json"
        $package = Get-Content -LiteralPath $packagePath -Raw | ConvertFrom-Json

        Get-LockedKordocVersion -LockPath $lockPath -NodePath $nodePath |
            Should -Be ([version]$package.dependencies.kordoc)
    }

    It "uses package.json as the Kordoc version source" {
        $verifierSource = Get-Content -LiteralPath $script:VerifierPath -Raw

        $verifierSource | Should -Not -Match "RequiredKordocVersion"
        $verifierSource | Should -Match "package\.dependencies\.kordoc"
    }

    It "warns when Codex Kordoc MCP does not use the local CLI" {
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null
        $configPath = Join-Path $TestDrive "config.toml"
        $expectedCliPath = Join-Path $script:RepoRoot `
            "tools\kordoc\node_modules\kordoc\dist\cli.js"
        $tomlCliPath = $expectedCliPath.Replace("\", "/")
        @"
[mcp_servers.kordoc]
command = "node.exe"
args = ["$tomlCliPath", "mcp"]
"@ | Set-Content -LiteralPath $configPath -Encoding UTF8

        Get-KordocMcpConfigurationWarning `
            -ConfigPath $configPath `
            -ExpectedCliPath $expectedCliPath |
            Should -BeNullOrEmpty

        @'
[mcp_servers.kordoc]
command = "npx.cmd"
args = ["-y", "kordoc@4.14.0", "mcp"]
'@ | Set-Content -LiteralPath $configPath -Encoding UTF8

        Get-KordocMcpConfigurationWarning `
            -ConfigPath $configPath `
            -ExpectedCliPath $expectedCliPath |
            Should -Match "local Kordoc CLI"
    }

    It "fails a requested HWPX gate when the skill is missing" {
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null
        $missingSkillRoot = Join-Path $TestDrive "missing-hwpx-skill"
        $uvPath = (Get-Command uv -ErrorAction Stop).Source

        $output = @(
            Invoke-HwpxValidationSmoke `
                -SkillRoot $missingSkillRoot `
                -UvPath $uvPath 6>&1
        )

        $output[-1] | Should -Not -Be 0
        ($output -join [Environment]::NewLine) | Should -Match '"ok": false'
    }

    It "rejects a legacy template-only skill without the required integration" {
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null
        $skillRoot = Join-Path $TestDrive "failing-hwpx-skill"
        $scriptsPath = Join-Path $skillRoot "scripts"
        $assetsPath = Join-Path $skillRoot "assets"
        [void](New-Item -ItemType Directory -Path $scriptsPath, $assetsPath)
        "raise SystemExit(23)" |
            Set-Content -LiteralPath (Join-Path $scriptsPath "validate.py") -Encoding UTF8
        "fixture" |
            Set-Content -LiteralPath (Join-Path $assetsPath "report-template.hwpx") -Encoding ASCII
        $uvPath = (Get-Command uv -ErrorAction Stop).Source

        $output = @(
            Invoke-HwpxValidationSmoke `
                -SkillRoot $skillRoot `
                -UvPath $uvPath 3>&1 6>&1
        )

        $output[-1] | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "Installed skill lacks"
    }

    It "restores the caller location when dot-sourced" {
        $before = (Get-Location).Path
        Push-Location $TestDrive
        try {
            $result = . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly
            (Get-Location).Path | Should -Be $TestDrive
            $result | Should -Be 0
        }
        finally {
            Pop-Location
        }

        (Get-Location).Path | Should -Be $before
    }

    It "returns a nonzero code for a failed external child command" {
        $fakeCommand = Join-Path $TestDrive "failed-child.cmd"
        Set-Content -LiteralPath $fakeCommand -Value "@exit /b 23" -NoNewline
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null

        $exitCode = Invoke-VerifyExternalCommand -FilePath $fakeCommand -Arguments @()

        $exitCode | Should -Be 23
    }

    It "returns only the exit code when a child emits output" {
        $fakeCommand = Join-Path $TestDrive "chatty-child.cmd"
        Set-Content -LiteralPath $fakeCommand -Value "@echo child-output`r`n@exit /b 23" -NoNewline
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null

        $exitCode = Invoke-VerifyExternalCommand -FilePath $fakeCommand -Arguments @()

        $exitCode | Should -Be 23
    }

    It "accepts successful child commands that write to stderr" {
        $fakeCommand = Join-Path $TestDrive "stderr-child.cmd"
        Set-Content -LiteralPath $fakeCommand -Value "@echo child-status 1>&2`r`n@exit /b 0" -NoNewline
        . $script:VerifierPath -Tier Static -PrerequisiteCheckOnly | Out-Null

        $exitCode = Invoke-VerifyExternalCommand -FilePath $fakeCommand -Arguments @()

        $exitCode | Should -Be 0
    }
}
