BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
}

Describe "Batch launcher input handling" {
    It "<Launcher> ignores a missing-path argument and does not execute metacharacters" -ForEach @(
        @{ Launcher = "convert-doc-to-docx.bat" }
        @{ Launcher = "convert-doc-to-md.bat" }
    ) {
        $launcherPath = Join-Path $script:RepoRoot $Launcher
        $sentinel = "BATCH_INJECTION_$([guid]::NewGuid().ToString('N'))"
        $command = '"{0}" "Z:\missing&echo {1}&rem"' -f $launcherPath, $sentinel

        $outputLines = & cmd.exe /d /s /c $command 2>&1
        $exitCode = $LASTEXITCODE
        $output = $outputLines -join [Environment]::NewLine

        $exitCode | Should -Be 0
        $output | Should -Not -Match ([regex]::Escape($sentinel))
    }

    It "<Launcher> does not execute a balanced-quote payload" -ForEach @(
        @{ Launcher = "convert-doc-to-docx.bat" }
        @{ Launcher = "convert-doc-to-md.bat" }
    ) {
        $launcherPath = Join-Path $script:RepoRoot $Launcher
        $sentinel = "QUOTED_INJECTION_$([guid]::NewGuid().ToString('N'))"
        $payload = 'foo"=="/?" echo NO & echo {0} & rem "' -f $sentinel

        $outputLines = & $launcherPath $payload 2>&1
        $exitCode = $LASTEXITCODE
        $output = $outputLines -join [Environment]::NewLine

        $output | Should -Not -Match ([regex]::Escape($sentinel))
        $exitCode | Should -Be 0
    }

    It "<Launcher> processes the project inbox when called without arguments" -ForEach @(
        @{ Launcher = "convert-doc-to-docx.bat" }
        @{ Launcher = "convert-doc-to-md.bat" }
    ) {
        $launcherPath = Join-Path $script:RepoRoot $Launcher

        & $launcherPath 2>&1 | Out-Null

        $LASTEXITCODE | Should -Be 0
    }

    It "<Script> accepts an explicit directory containing an ampersand" -ForEach @(
        @{ Script = "tools/doc-to-docx/convert-doc-to-docx.ps1" }
        @{ Script = "tools/doc-to-docx/convert-doc-to-md.ps1" }
    ) {
        $scriptPath = Join-Path $script:RepoRoot $Script
        $inputDirectory = Join-Path $TestDrive "input&documents"
        [void](New-Item -ItemType Directory -Path $inputDirectory -Force)

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Path $inputDirectory -Recurse 2>&1 |
            Out-Null

        $LASTEXITCODE | Should -Be 0
    }
}
