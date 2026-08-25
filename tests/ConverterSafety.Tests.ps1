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

    function Write-LargeLegacyDocFixture {
        param(
            [Parameter(Mandatory)]
            [string]$Path
        )

        $builder = [System.Text.StringBuilder]::new(16000000)
        [void]$builder.Append("{\rtf1\ansi ")
        for ($index = 0; $index -lt 600000; $index++) {
            [void]$builder.Append("timeout fixture\par ")
        }
        [void]$builder.Append("}")
        [System.IO.File]::WriteAllText(
            $Path,
            $builder.ToString(),
            [System.Text.Encoding]::ASCII
        )
    }
}

Describe "DOC converter safety" -Tag "Office" {
    It "synchronizes timeout firing and cancellation" {
        $emptyDirectory = Join-Path $TestDrive "timeout-guard-load"
        [void](New-Item -ItemType Directory -Path $emptyDirectory)
        . $script:DocxConverter -Path $emptyDirectory -OutputFormat Json | Out-Null

        $timeoutEventName = "HwpxWizTimeout_$([guid]::NewGuid().ToString('N'))"
        $timeoutEvent = [System.Threading.EventWaitHandle]::new(
            $false,
            [System.Threading.EventResetMode]::ManualReset,
            $timeoutEventName
        )
        $timeoutCommand = '$event = [Threading.EventWaitHandle]::OpenExisting("{0}"); [void]$event.WaitOne()' -f `
            $timeoutEventName
        $timeoutEncoded = [Convert]::ToBase64String(
            [Text.Encoding]::Unicode.GetBytes($timeoutCommand)
        )
        $timeoutProcess = Start-Process powershell.exe `
            -ArgumentList @("-NoProfile", "-EncodedCommand", $timeoutEncoded) `
            -PassThru
        $timeoutGuard = [HwpxWiz.ProcessTimeoutGuard]::new()

        try {
            $timeoutGuard.Arm($timeoutProcess, 10)
            $timeoutProcess.WaitForExit(5000) | Should -BeTrue
            $timeoutGuard.Dispose()
            $timeoutGuard.TimedOut | Should -BeTrue
        }
        finally {
            $timeoutGuard.Dispose()
            if (-not $timeoutProcess.HasExited) {
                $timeoutProcess.Kill()
            }
            $timeoutEvent.Dispose()
        }

        $cancelEventName = "HwpxWizCancel_$([guid]::NewGuid().ToString('N'))"
        $cancelEvent = [System.Threading.EventWaitHandle]::new(
            $false,
            [System.Threading.EventResetMode]::ManualReset,
            $cancelEventName
        )
        $cancelCommand = '$event = [Threading.EventWaitHandle]::OpenExisting("{0}"); [void]$event.WaitOne()' -f `
            $cancelEventName
        $cancelEncoded = [Convert]::ToBase64String(
            [Text.Encoding]::Unicode.GetBytes($cancelCommand)
        )
        $cancelProcess = Start-Process powershell.exe `
            -ArgumentList @("-NoProfile", "-EncodedCommand", $cancelEncoded) `
            -PassThru
        $cancelGuard = [HwpxWiz.ProcessTimeoutGuard]::new()

        try {
            $cancelGuard.Arm($cancelProcess, 60000)
            $cancelGuard.Dispose()
            $cancelGuard.TimedOut | Should -BeFalse
            [void]$cancelEvent.Set()
            $cancelProcess.WaitForExit(5000) | Should -BeTrue
        }
        finally {
            $cancelGuard.Dispose()
            if (-not $cancelProcess.HasExited) {
                $cancelProcess.Kill()
            }
            $cancelEvent.Dispose()
        }
    }

    It "rejects wildcard LogPath without mutating a source document" {
        $sourcePath = Join-Path $TestDrive "wildcard-log-source.doc"
        $wildcardLogPath = Join-Path $TestDrive "*.doc"
        Write-LegacyDocFixture -Path $sourcePath
        $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash

        $outputLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:DocxConverter -Path $sourcePath `
                -LogPath $wildcardLogPath -OutputFormat Json
        )
        $exitCode = $LASTEXITCODE
        $jsonOutput = $outputLines -join [Environment]::NewLine
        $jsonOutput | Should -Not -BeNullOrEmpty
        $records = @($jsonOutput | ConvertFrom-Json)

        $exitCode | Should -Not -Be 0
        $records[0].Reason | Should -BeExactly "LogPath must be literal"
        (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash |
            Should -BeExactly $sourceHash
    }

    It "resolves SecureString and environment document keys" {
        $emptyDirectory = Join-Path $TestDrive "password-resolution"
        [void](New-Item -ItemType Directory -Path $emptyDirectory)
        . $script:DocxConverter -Path $emptyDirectory | Out-Null
        $secret = "secure-value-$([guid]::NewGuid().ToString('N'))"
        $secureSecret = [System.Security.SecureString]::new()
        foreach ($character in $secret.ToCharArray()) {
            $secureSecret.AppendChar($character)
        }
        $secureSecret.MakeReadOnly()
        $originalEnvironmentPassword = $env:HWPX_WIZ_DOC_PASSWORD

        try {
            Resolve-DocumentKey -Value $secureSecret | Should -BeExactly $secret
            $env:HWPX_WIZ_DOC_PASSWORD = $secret
            Resolve-DocumentKey -Value $null | Should -BeExactly $secret
        }
        finally {
            $env:HWPX_WIZ_DOC_PASSWORD = $originalEnvironmentPassword
        }
    }

    It "emits one-line JSON records without human summary output" {
        $sourcePath = Join-Path $TestDrive "json-skip.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($targetPath, "EXISTING_DOCX", [System.Text.Encoding]::ASCII)

        $outputLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:DocxConverter -Path $sourcePath -OutputFormat Json
        )
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Be 0
        $outputLines.Count | Should -Be 1
        $outputLines[0] | Should -Not -Match "CONVERSION SUMMARY"
        $jsonOutput = $outputLines -join [Environment]::NewLine
        $jsonOutput | Should -Not -BeNullOrEmpty
        $records = @($jsonOutput | ConvertFrom-Json)
        $records.Count | Should -Be 1
        $records[0].Status | Should -BeExactly "Skipped"
        $records[0].Source | Should -BeExactly $sourcePath
    }

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

    It "records a timeout and leaves no owned Word process" {
        $sourcePath = Join-Path $TestDrive "timeout-large.doc"
        $targetPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        Write-LargeLegacyDocFixture -Path $sourcePath
        $wordIdsBefore = @(
            Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                ForEach-Object Id
        )

        $outputLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:DocxConverter -Path $sourcePath `
                -TimeoutSeconds 1 -OutputFormat Json
        )
        $exitCode = $LASTEXITCODE
        $timeoutJson = $outputLines -join [Environment]::NewLine
        $timeoutJson | Should -Not -BeNullOrEmpty
        $records = @($timeoutJson | ConvertFrom-Json)
        $newWordProcesses = @(
            Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -notin $wordIdsBefore }
        )

        $exitCode | Should -Not -Be 0
        $records.Count | Should -Be 1
        $records[0].Status | Should -BeExactly "Failed"
        $records[0].Reason | Should -BeExactly "Timeout" `
            -Because ($outputLines -join [Environment]::NewLine)
        (Test-Path -LiteralPath $targetPath) | Should -BeFalse
        $newWordProcesses.Count | Should -Be 0
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

Describe "Markdown converter safety" -Tag "Office" {
    It "rejects LogPath that collides with source or output" {
        $sourcePath = Join-Path $TestDrive "markdown-log-conflict.doc"
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($docxPath, "DOCX_FIXTURE", [System.Text.Encoding]::ASCII)
        [System.IO.File]::WriteAllText($markdownPath, "MARKDOWN_FIXTURE", [System.Text.Encoding]::UTF8)
        $sourceHash = (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash

        $outputLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:MarkdownConverter -Path $sourcePath `
                -LogPath $sourcePath -OutputFormat Json
        )
        $exitCode = $LASTEXITCODE
        $records = @($outputLines[0] | ConvertFrom-Json)

        $exitCode | Should -Not -Be 0
        $records[0].Reason | Should -BeExactly "LogPath conflicts with a source or output"
        (Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash |
            Should -BeExactly $sourceHash
        [System.IO.File]::ReadAllText($docxPath, [System.Text.Encoding]::ASCII) |
            Should -BeExactly "DOCX_FIXTURE"
        [System.IO.File]::ReadAllText($markdownPath, [System.Text.Encoding]::UTF8) |
            Should -BeExactly "MARKDOWN_FIXTURE"
    }

    It "accepts the shared document key options" {
        $command = Get-Command -Name $script:MarkdownConverter

        $command.Parameters.ContainsKey("DocumentKey") | Should -BeTrue
        @($command.Parameters["DocumentKey"].Aliases) | Should -Contain "Password"
    }

    It "propagates a DOC conversion timeout" {
        $sourcePath = Join-Path $TestDrive "timeout-markdown.doc"
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        Write-LargeLegacyDocFixture -Path $sourcePath
        $wordIdsBefore = @(
            Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                ForEach-Object Id
        )

        $outputLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:MarkdownConverter -Path $sourcePath `
                -TimeoutSeconds 1 -OutputFormat Json
        )
        $exitCode = $LASTEXITCODE
        $timeoutJson = $outputLines -join [Environment]::NewLine
        $timeoutJson | Should -Not -BeNullOrEmpty
        $records = @($timeoutJson | ConvertFrom-Json)
        $newWordProcesses = @(
            Get-Process -Name WINWORD -ErrorAction SilentlyContinue |
                Where-Object { $_.Id -notin $wordIdsBefore }
        )

        $exitCode | Should -Not -Be 0
        $records.Count | Should -Be 1
        $records[0].Reason | Should -BeExactly "Timeout" `
            -Because ($outputLines -join [Environment]::NewLine)
        (Test-Path -LiteralPath $docxPath) | Should -BeFalse
        (Test-Path -LiteralPath $markdownPath) | Should -BeFalse
        $newWordProcesses.Count | Should -Be 0
    }

    It "emits one-line JSON and writes the requested log" {
        $sourcePath = Join-Path $TestDrive "json-markdown.doc"
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        $logPath = Join-Path $TestDrive "json-markdown.log"
        Write-LegacyDocFixture -Path $sourcePath
        [System.IO.File]::WriteAllText($docxPath, "DOCX_FIXTURE", [System.Text.Encoding]::ASCII)
        [System.IO.File]::WriteAllText($markdownPath, "MARKDOWN_FIXTURE", [System.Text.Encoding]::UTF8)
        [System.IO.File]::SetLastWriteTimeUtc($sourcePath, [datetime]::UtcNow.AddMinutes(-2))
        [System.IO.File]::SetLastWriteTimeUtc($docxPath, [datetime]::UtcNow.AddMinutes(-1))
        [System.IO.File]::SetLastWriteTimeUtc($markdownPath, [datetime]::UtcNow)

        $outputLines = @(
            & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:MarkdownConverter -Path $sourcePath `
                -OutputFormat Json -LogPath $logPath
        )
        $exitCode = $LASTEXITCODE

        $exitCode | Should -Be 0
        $outputLines.Count | Should -Be 1
        $records = @($outputLines[0] | ConvertFrom-Json)
        $records.Count | Should -Be 1
        $records[0].Status | Should -BeExactly "Skipped"
        Test-Path -LiteralPath $logPath | Should -BeTrue
        Get-Content -LiteralPath $logPath -Raw | Should -Match "Markdown is up to date"
    }

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
        $converterSource | Should -Match "node_modules[\\/]+kordoc[\\/]+dist[\\/]+cli\.js"
        $converterSource | Should -Not -Match "node_modules[\\/]+\.bin"
    }

    It "converts Markdown under an ampersand path without cmd reparsing" {
        $inputDirectory = Join-Path $TestDrive "a&b"
        $sourcePath = Join-Path $inputDirectory "ampersand.doc"
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        [void](New-Item -ItemType Directory -Path $inputDirectory)
        Write-LegacyDocFixture -Path $sourcePath

        $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
            -Path $sourcePath 2>&1
        $exitCode = $LASTEXITCODE
        $output = $outputLines -join [Environment]::NewLine

        $exitCode | Should -Be 0 -Because $output
        (Test-Path -LiteralPath $markdownPath) | Should -BeTrue
        [System.IO.File]::ReadAllText($markdownPath) | Should -Match "regression fixture"
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
const fs = require("fs");
const args = process.argv.slice(2);
fs.writeFileSync(process.env.FAKE_KORDOC_ARGS, JSON.stringify(args));
const outputIndex = args.indexOf("-o");
if (outputIndex < 0 || !args[outputIndex + 1]) {
    process.exit(2);
}
fs.writeFileSync(process.env.FAKE_KORDOC_OUTPUT, args[outputIndex + 1]);
fs.writeFileSync(args[outputIndex + 1], "fake markdown");
'@ | Set-Content -LiteralPath (Join-Path $fakeBin "kordoc.js") -Encoding ASCII

        $originalPath = $env:PATH
        $originalArgsCapture = $env:FAKE_KORDOC_ARGS
        $originalOutputCapture = $env:FAKE_KORDOC_OUTPUT
        try {
            $env:PATH = "$fakeBin;$originalPath"
            $env:FAKE_KORDOC_ARGS = $capturedArgsPath
            $env:FAKE_KORDOC_OUTPUT = $capturedOutputPath

            $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
                -Path $sourcePath -KordocCliPath (Join-Path $fakeBin "kordoc.js") 2>&1
            $exitCode = $LASTEXITCODE
            $output = $outputLines -join [Environment]::NewLine
        }
        finally {
            $env:PATH = $originalPath
            $env:FAKE_KORDOC_ARGS = $originalArgsCapture
            $env:FAKE_KORDOC_OUTPUT = $originalOutputCapture
        }

        $exitCode | Should -Be 0
        $output | Should -Match "Status"
        $output | Should -Match "Converted"
        $output | Should -Not -Match "TotalFiles"
        (Test-Path -LiteralPath $markdownPath) | Should -BeTrue
        $capturedArgs = [System.IO.File]::ReadAllText($capturedArgsPath) | ConvertFrom-Json
        $capturedArgs | Should -Contain "--silent"
        $capturedArgs | Should -Contain $docxPath
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
        'process.exit(17);' |
            Set-Content -LiteralPath (Join-Path $fakeBin "kordoc.js") -Encoding ASCII

        $originalPath = $env:PATH
        try {
            $env:PATH = "$fakeBin;$originalPath"
            $outputLines = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
                -Path $sourcePath -Overwrite -KordocCliPath (Join-Path $fakeBin "kordoc.js") 2>&1
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
        'process.exit(17);' |
            Set-Content -LiteralPath (Join-Path $fakeBin "kordoc.js") -Encoding ASCII

        $originalPath = $env:PATH
        try {
            $env:PATH = "$fakeBin;$originalPath"
            & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $script:MarkdownConverter `
                -Path $sourcePath -KordocCliPath (Join-Path $fakeBin "kordoc.js") 2>&1 |
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
