BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot

    function Initialize-IsolatedLauncherRepository {
        param(
            [Parameter(Mandatory)]
            [string]$Launcher
        )

        $isolatedRoot = Join-Path $TestDrive ([System.IO.Path]::GetFileNameWithoutExtension($Launcher))
        $toolsDirectory = Join-Path $isolatedRoot "tools\doc-to-docx"
        $inboxPath = Join-Path $isolatedRoot "inbox"
        $capturePath = Join-Path $isolatedRoot "captured-args.json"
        [void](New-Item -ItemType Directory -Path $toolsDirectory -Force)
        [void](New-Item -ItemType Directory -Path $inboxPath -Force)
        Copy-Item -LiteralPath (Join-Path $script:RepoRoot $Launcher) `
            -Destination (Join-Path $isolatedRoot $Launcher)

        $stub = @'
param(
    [string]$Path,
    [switch]$Recurse
)

[pscustomobject]@{
    Path = $Path
    Recurse = $Recurse.IsPresent
} | ConvertTo-Json -Compress |
    Set-Content -LiteralPath $env:HWPX_WIZ_LAUNCHER_CAPTURE -Encoding UTF8
'@
        $stub | Set-Content -LiteralPath (Join-Path $toolsDirectory "convert-doc-to-docx.ps1") -Encoding UTF8
        $stub | Set-Content -LiteralPath (Join-Path $toolsDirectory "convert-doc-to-md.ps1") -Encoding UTF8

        return [pscustomobject]@{
            LauncherPath = Join-Path $isolatedRoot $Launcher
            InboxPath = $inboxPath
            CapturePath = $capturePath
        }
    }

    function Assert-IsolatedLauncherInvocation {
        param(
            [Parameter(Mandatory)]
            [pscustomobject]$Repository
        )

        Test-Path -LiteralPath $Repository.CapturePath | Should -BeTrue
        $captured = Get-Content -LiteralPath $Repository.CapturePath -Raw | ConvertFrom-Json
        $captured.Path | Should -BeExactly $Repository.InboxPath
        $captured.Recurse | Should -BeTrue
    }
}

Describe "Batch launcher input handling" -Tag "Static" {
    It "<Launcher> ignores a missing-path argument and does not execute metacharacters" -ForEach @(
        @{ Launcher = "convert-doc-to-docx.bat" }
        @{ Launcher = "convert-doc-to-md.bat" }
    ) {
        $repository = Initialize-IsolatedLauncherRepository -Launcher $Launcher
        $sentinel = "BATCH_INJECTION_$([guid]::NewGuid().ToString('N'))"
        $command = '"{0}" "Z:\missing&echo {1}&rem"' -f $repository.LauncherPath, $sentinel

        $originalCapture = $env:HWPX_WIZ_LAUNCHER_CAPTURE
        try {
            $env:HWPX_WIZ_LAUNCHER_CAPTURE = $repository.CapturePath
            $outputLines = & cmd.exe /d /s /c $command 2>&1
            $exitCode = $LASTEXITCODE
            $output = $outputLines -join [Environment]::NewLine
        }
        finally {
            $env:HWPX_WIZ_LAUNCHER_CAPTURE = $originalCapture
        }

        $exitCode | Should -Be 0
        $output | Should -Not -Match ([regex]::Escape($sentinel))
        Assert-IsolatedLauncherInvocation -Repository $repository
    }

    It "<Launcher> does not execute a balanced-quote payload" -ForEach @(
        @{ Launcher = "convert-doc-to-docx.bat" }
        @{ Launcher = "convert-doc-to-md.bat" }
    ) {
        $repository = Initialize-IsolatedLauncherRepository -Launcher $Launcher
        $sentinel = "QUOTED_INJECTION_$([guid]::NewGuid().ToString('N'))"
        $payload = 'foo"=="/?" echo NO & echo {0} & rem "' -f $sentinel

        $originalCapture = $env:HWPX_WIZ_LAUNCHER_CAPTURE
        try {
            $env:HWPX_WIZ_LAUNCHER_CAPTURE = $repository.CapturePath
            $outputLines = & $repository.LauncherPath $payload 2>&1
            $exitCode = $LASTEXITCODE
            $output = $outputLines -join [Environment]::NewLine
        }
        finally {
            $env:HWPX_WIZ_LAUNCHER_CAPTURE = $originalCapture
        }

        $output | Should -Not -Match ([regex]::Escape($sentinel))
        $exitCode | Should -Be 0
        Assert-IsolatedLauncherInvocation -Repository $repository
    }

    It "<Launcher> processes the project inbox when called without arguments" -ForEach @(
        @{ Launcher = "convert-doc-to-docx.bat" }
        @{ Launcher = "convert-doc-to-md.bat" }
    ) {
        $repository = Initialize-IsolatedLauncherRepository -Launcher $Launcher

        $originalCapture = $env:HWPX_WIZ_LAUNCHER_CAPTURE
        try {
            $env:HWPX_WIZ_LAUNCHER_CAPTURE = $repository.CapturePath
            & $repository.LauncherPath 2>&1 | Out-Null
            $exitCode = $LASTEXITCODE
        }
        finally {
            $env:HWPX_WIZ_LAUNCHER_CAPTURE = $originalCapture
        }

        $exitCode | Should -Be 0
        Assert-IsolatedLauncherInvocation -Repository $repository
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
