BeforeAll {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:DocxConverter = Join-Path $script:RepoRoot "tools/doc-to-docx/convert-doc-to-docx.ps1"
    $script:MarkdownConverter = Join-Path $script:RepoRoot "tools/doc-to-docx/convert-doc-to-md.ps1"

    function Write-LegacyDocFixture {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $rtf = "{\rtf1\ansi HWPX_wiz regression fixture\par}"
        [System.IO.File]::WriteAllText($Path, $rtf, [System.Text.Encoding]::ASCII)
    }
}

Describe "DOC converter safety" {
    It "emits a structured skipped record and preserves dot-source callers" {
        $sourcePath = Join-Path $TestDrive "structured-skip.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($targetPath, "EXISTING_DOCX", [System.Text.Encoding]::ASCII)

        $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File $script:DocxConverter -Path $sourcePath 2>&1
        $exitCode = $LASTEXITCODE
        $output = $outputLines -join [Environment]::NewLine

        $exitCode | Should -Be 0 -Because $output
        $output | Should -Match "Status"
        $output | Should -Match "Skipped"
        $output | Should -Not -Match "TotalFiles"

        $dotCommand = '$result = . "{0}" -Path "{1}"; Write-Output "SENTINEL"; $result | ConvertTo-Json -Compress' -f `
            $script:DocxConverter, $sourcePath
        $dotOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $dotCommand 2>&1
        $dotOutput -join [Environment]::NewLine | Should -Match "SENTINEL"
    }

    It "captures a nonzero Word process ID" {
        $sourcePath = Join-Path $TestDrive "word-pid.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        Write-LegacyDocFixture -Path $sourcePath

        $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:DocxConverter -Path $sourcePath 2>&1
        $exitCode = $LASTEXITCODE
        $output = $outputLines -join [Environment]::NewLine

        $exitCode | Should -Be 0 -Because $output
        $output | Should -Match "Word COM instance created \(PID [1-9][0-9]*, provenance Owned\)"
        $output | Should -Not -Match "Force-killing orphaned Word process"
        (Test-Path -LiteralPath $targetPath) | Should -BeTrue
    }

    It "does not terminate an independently opened Word process" {
        if (-not ([System.Management.Automation.PSTypeName]"HwpxWiz.User32").Type) {
            Add-Type -Namespace HwpxWiz -Name User32 -MemberDefinition @"
                [System.Runtime.InteropServices.DllImport("user32.dll")]
                public static extern uint GetWindowThreadProcessId(System.IntPtr hWnd, out uint lpdwProcessId);
"@
        }

        $sourcePath = Join-Path $TestDrive "owned-word.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        Write-LegacyDocFixture -Path $sourcePath

        $userWord = $null
        $probeDocument = $null
        $userProcessId = 0
        try {
            $userWord = New-Object -ComObject Word.Application
            $userWord.Visible = $false
            $probeDocument = $userWord.Documents.Add()
            [uint32]$resolvedProcessId = 0
            $null = [HwpxWiz.User32]::GetWindowThreadProcessId(
                [IntPtr]$probeDocument.ActiveWindow.Hwnd,
                [ref]$resolvedProcessId
            )
            $userProcessId = [int]$resolvedProcessId
            $userProcessId | Should -BeGreaterThan 0

            $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:DocxConverter -Path $sourcePath 2>&1
            $exitCode = $LASTEXITCODE
            $output = $outputLines -join [Environment]::NewLine

            $exitCode | Should -Be 0 -Because $output
            $output | Should -Match "Word COM instance created \(PID [1-9][0-9]*, provenance Owned\)"
            $output | Should -Not -Match ("PID {0}" -f $userProcessId)
            (Get-Process -Id $userProcessId -ErrorAction Stop).Id | Should -Be $userProcessId
            (Test-Path -LiteralPath $targetPath) | Should -BeTrue
        }
        finally {
            if ($null -ne $probeDocument) {
                $probeDocument.Close([ref]0)
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($probeDocument)
            }
            if ($null -ne $userWord) {
                $userWord.Quit([ref]0)
                [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($userWord)
            }
        }
    }

    It "preserves the previous DOCX when overwrite conversion fails" {
        $sourcePath = Join-Path $TestDrive "locked-source.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $sentinel = "KEEP_PREVIOUS_DOCX"
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($targetPath, $sentinel, [System.Text.Encoding]::ASCII)
        $sourceLock = [System.IO.File]::Open(
            $sourcePath,
            [System.IO.FileMode]::Open,
            [System.IO.FileAccess]::ReadWrite,
            [System.IO.FileShare]::None
        )

        try {
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:DocxConverter -Path $sourcePath -Overwrite 2>&1 |
                Out-Null
            $exitCode = $LASTEXITCODE
        }
        finally {
            $sourceLock.Dispose()
        }

        $exitCode | Should -Not -Be 0
        [System.IO.File]::ReadAllText($targetPath, [System.Text.Encoding]::ASCII) |
            Should -BeExactly $sentinel
    }

    It "atomically replaces an existing DOCX after successful conversion" {
        $sourcePath = Join-Path $TestDrive "overwrite-success.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $sentinel = "OLD_DOCX"
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($targetPath, $sentinel, [System.Text.Encoding]::ASCII)

        $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:DocxConverter -Path $sourcePath -Overwrite 2>&1
        $exitCode = $LASTEXITCODE
        $output = $outputLines -join [Environment]::NewLine

        $exitCode | Should -Be 0 -Because $output
        [System.IO.File]::ReadAllText($targetPath, [System.Text.Encoding]::ASCII) |
            Should -Not -BeExactly $sentinel
        {
            $archive = [System.IO.Compression.ZipFile]::OpenRead($targetPath)
            $archive.Dispose()
        } | Should -Not -Throw
    }
}

Describe "Markdown converter safety" {
    It "does not terminate a dot-sourced Markdown caller" {
        $emptyDirectory = Join-Path $TestDrive "empty-markdown-input"
        [void](New-Item -ItemType Directory -Path $emptyDirectory -Force)
        $dotCommand = '$result = . "{0}" -Path "{1}"; Write-Output "SENTINEL"; $result | ConvertTo-Json -Compress' -f `
            $script:MarkdownConverter, $emptyDirectory

        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass -Command $dotCommand 2>&1

        ($output -join [Environment]::NewLine) | Should -Match "SENTINEL"
    }

    It "uses the repository-local locked Kordoc CLI" {
        $kordocPackage = Join-Path $script:RepoRoot "tools/kordoc/package.json"
        $kordocLock = Join-Path $script:RepoRoot "tools/kordoc/package-lock.json"
        $converterSource = Get-Content -LiteralPath $script:MarkdownConverter -Raw

        Test-Path -LiteralPath $kordocPackage | Should -BeTrue
        Test-Path -LiteralPath $kordocLock | Should -BeTrue
        $converterSource | Should -Match "KordocCliPath"
        $converterSource | Should -Not -Match "npx\.cmd"
        $converterSource | Should -Match "node_modules[\\/]+\.bin"
    }

    It "rejects stale Markdown instead of silently skipping it" {
        $sourcePath = Join-Path $TestDrive "stale-markdown.doc"
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($markdownPath, "STALE", [System.Text.Encoding]::UTF8)
        [System.IO.File]::SetLastWriteTimeUtc($markdownPath, [datetime]::UtcNow.AddHours(-1))
        [System.IO.File]::SetLastWriteTimeUtc($sourcePath, [datetime]::UtcNow)

        & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter -Path $sourcePath 2>&1 |
            Out-Null
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Not -Be 0
        [System.IO.File]::ReadAllText($markdownPath, [System.Text.Encoding]::UTF8) |
            Should -BeExactly "STALE"
    }

    It "uses the validated Kordoc version and a temporary Markdown output" {
        $sourcePath = Join-Path $TestDrive "atomic-markdown.doc"
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        $fakeBin = Join-Path $TestDrive "fake-bin"
        $capturedArgsPath = Join-Path $TestDrive "kordoc-args.txt"
        $capturedOutputPath = Join-Path $TestDrive "kordoc-output.txt"
        [void](New-Item -ItemType Directory -Path $fakeBin)
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($docxPath, "DOCX_FIXTURE", [System.Text.Encoding]::ASCII)
        [System.IO.File]::SetLastWriteTimeUtc($sourcePath, [datetime]::UtcNow.AddMinutes(-1))
        [System.IO.File]::SetLastWriteTimeUtc($docxPath, [datetime]::UtcNow)
        @'
@echo off
> "%FAKE_NPX_ARGS%" echo %*
:parse
if "%~1"=="" exit /b 2
if /i "%~1"=="-o" (
    > "%FAKE_NPX_OUTPUT%" echo %~2
    > "%~2" echo fake markdown
    exit /b 0
)
shift
goto parse
'@ | Set-Content -LiteralPath (Join-Path $fakeBin "kordoc.cmd") -Encoding ASCII

        $originalPath = $env:PATH
        $originalArgsCapture = $env:FAKE_NPX_ARGS
        $originalOutputCapture = $env:FAKE_NPX_OUTPUT
        try {
            $env:PATH = "$fakeBin;$originalPath"
            $env:FAKE_NPX_ARGS = $capturedArgsPath
            $env:FAKE_NPX_OUTPUT = $capturedOutputPath

            $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
                -Path $sourcePath -KordocCliPath (Join-Path $fakeBin "kordoc.cmd") 2>&1
            $exitCode = $LASTEXITCODE
            $output = $outputLines -join [Environment]::NewLine
        }
        finally {
            $env:PATH = $originalPath
            $env:FAKE_NPX_ARGS = $originalArgsCapture
            $env:FAKE_NPX_OUTPUT = $originalOutputCapture
        }

        $exitCode | Should -Be 0
        $output | Should -Match "Status"
        $output | Should -Match "Converted"
        $output | Should -Not -Match "TotalFiles"
        (Test-Path -LiteralPath $markdownPath) | Should -BeTrue
        [System.IO.File]::ReadAllText($capturedArgsPath) | Should -Match "--silent"
        $kordocOutputPath = [System.IO.File]::ReadAllText($capturedOutputPath).Trim()
        $kordocOutputPath | Should -Not -Be $markdownPath
        (Split-Path -Parent $kordocOutputPath) | Should -Be (Split-Path -Parent $markdownPath)
    }

    It "restores the previous DOCX when Kordoc overwrite conversion fails" {
        $sourcePath = Join-Path $TestDrive "kordoc-fail-rollback.doc"
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        $fakeBin = Join-Path $TestDrive "fake-bin-fail"
        $docxSentinel = "KEEP_PREVIOUS_DOCX"
        $markdownSentinel = "KEEP_PREVIOUS_MD"
        [void](New-Item -ItemType Directory -Path $fakeBin)
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($docxPath, $docxSentinel, [System.Text.Encoding]::ASCII)
        [System.IO.File]::WriteAllText($markdownPath, $markdownSentinel, [System.Text.Encoding]::UTF8)
        @'
@echo off
exit /b 17
'@ | Set-Content -LiteralPath (Join-Path $fakeBin "kordoc.cmd") -Encoding ASCII

        $originalPath = $env:PATH
        try {
            $env:PATH = "$fakeBin;$originalPath"
            $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
                -Path $sourcePath -Overwrite -KordocCliPath (Join-Path $fakeBin "kordoc.cmd") 2>&1
            $exitCode = $LASTEXITCODE
            $output = $outputLines -join [Environment]::NewLine
        }
        finally {
            $env:PATH = $originalPath
        }

        $exitCode | Should -Not -Be 0
        $output | Should -Match "Status"
        $output | Should -Match "Failed"
        [System.IO.File]::ReadAllText($docxPath, [System.Text.Encoding]::ASCII) |
            Should -BeExactly $docxSentinel
        [System.IO.File]::ReadAllText($markdownPath, [System.Text.Encoding]::UTF8) |
            Should -BeExactly $markdownSentinel
        @(Get-ChildItem -LiteralPath $TestDrive -Force |
            Where-Object { $_.Name -match "^\.kordoc-fail-rollback\." }).Count |
            Should -Be 0
    }

    It "removes a newly created DOCX when Kordoc fails" {
        $sourcePath = Join-Path $TestDrive "kordoc-fail-new.docx-source.doc"
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        $fakeBin = Join-Path $TestDrive "fake-bin-fail-new"
        [void](New-Item -ItemType Directory -Path $fakeBin)
        Write-LegacyDocFixture -Path $sourcePath
        @'
@echo off
exit /b 17
'@ | Set-Content -LiteralPath (Join-Path $fakeBin "kordoc.cmd") -Encoding ASCII

        $originalPath = $env:PATH
        try {
            $env:PATH = "$fakeBin;$originalPath"
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
                -Path $sourcePath -KordocCliPath (Join-Path $fakeBin "kordoc.cmd") 2>&1 |
                Out-Null
            $exitCode = $LASTEXITCODE
        }
        finally {
            $env:PATH = $originalPath
        }

        $exitCode | Should -Not -Be 0
        (Test-Path -LiteralPath $docxPath) | Should -BeFalse
        (Test-Path -LiteralPath $markdownPath) | Should -BeFalse
    }
}
