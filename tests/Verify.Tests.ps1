BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:VerifierPath = Join-Path $script:RepoRoot "tools\verify.ps1"
    $script:PowerShellPath = (Get-Command powershell.exe -ErrorAction Stop).Source
}

Describe "Repository verifier" {
    It "has a verifier entry point" {
        Test-Path -LiteralPath $script:VerifierPath | Should -BeTrue
    }

    It "passes its prerequisite-only seam without running the suite" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly 6>&1 2>&1
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

    It "reports missing Node or local Kordoc installation guidance" {
        $output = & $script:PowerShellPath -NoProfile -ExecutionPolicy Bypass `
            -File $script:VerifierPath -PrerequisiteCheckOnly `
            -SimulateMissing Node,Kordoc 6>&1 2>&1
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "npm ci --prefix \.\\tools\\kordoc"
    }

    It "reads the locked Kordoc version from the npm lockfile" {
        . $script:VerifierPath -PrerequisiteCheckOnly | Out-Null
        $nodePath = (Get-Command node.exe -ErrorAction Stop).Source
        $lockPath = Join-Path $script:RepoRoot "tools\kordoc\package-lock.json"

        Get-LockedKordocVersion -LockPath $lockPath -NodePath $nodePath |
            Should -Be ([version]"4.9.0")
    }

    It "restores the caller location when dot-sourced" {
        $before = (Get-Location).Path
        Push-Location $TestDrive
        try {
            $result = . $script:VerifierPath -PrerequisiteCheckOnly
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
        . $script:VerifierPath -PrerequisiteCheckOnly | Out-Null

        $exitCode = Invoke-VerifyExternalCommand -FilePath $fakeCommand -Arguments @()

        $exitCode | Should -Be 23
    }

    It "returns only the exit code when a child emits output" {
        $fakeCommand = Join-Path $TestDrive "chatty-child.cmd"
        Set-Content -LiteralPath $fakeCommand -Value "@echo child-output`r`n@exit /b 23" -NoNewline
        . $script:VerifierPath -PrerequisiteCheckOnly | Out-Null

        $exitCode = Invoke-VerifyExternalCommand -FilePath $fakeCommand -Arguments @()

        $exitCode | Should -Be 23
    }

    It "accepts successful child commands that write to stderr" {
        $fakeCommand = Join-Path $TestDrive "stderr-child.cmd"
        Set-Content -LiteralPath $fakeCommand -Value "@echo child-status 1>&2`r`n@exit /b 0" -NoNewline
        . $script:VerifierPath -PrerequisiteCheckOnly | Out-Null

        $exitCode = Invoke-VerifyExternalCommand -FilePath $fakeCommand -Arguments @()

        $exitCode | Should -Be 0
    }
}
