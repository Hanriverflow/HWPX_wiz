BeforeAll {
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:DocxConverter = Join-Path $script:RepoRoot "tools/doc-to-docx/convert-doc-to-docx.ps1"
    $script:MarkdownConverter = Join-Path $script:RepoRoot "tools/doc-to-docx/convert-doc-to-md.ps1"
    $script:LocalKordoc = Join-Path $script:RepoRoot "tools/kordoc/node_modules/.bin/kordoc.cmd"
    $script:FixtureRoot = Join-Path $PSScriptRoot "fixtures/doc"
    $script:ExpectedFixtureHashes = @{
        "special-name.doc.base64" = "2a7c5606fb6f7aeed9b88e6a72f162665bc7c5849ba0d475b5b9479a0dabbedf"
        "password-protected.doc.base64" = "bb3c76697092123e2b21165b0e2021aa353f355dc23995f4ab2ea8ba0c972435"
        "corrupt.doc.base64" = "570e5ac66722e9169fee82ef08036e6eebde2741ba0b91f26fdf3050f7bddcc1"
    }

    function Restore-DocFixture {
        param(
            [Parameter(Mandatory)]
            [string]$Base64Name,

            [Parameter(Mandatory)]
            [string]$Destination
        )

        $encoded = Get-Content -LiteralPath (Join-Path $script:FixtureRoot $Base64Name) -Raw
        [System.IO.File]::WriteAllBytes($Destination, [Convert]::FromBase64String($encoded.Trim()))
        (Get-Sha256 -Path $Destination) | Should -BeExactly $script:ExpectedFixtureHashes[$Base64Name]
    }

    function Get-Sha256 {
        param([Parameter(Mandatory)][string]$Path)
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
    }
}

Describe "Real DOC integration boundaries" {
    It "converts a special-name DOC through Word and local Kordoc" {
        $koreanLabel = [string]::Concat([char]0xBCF4, [char]0xACE0, [char]0xC11C)
        $inputDirectory = Join-Path $TestDrive "$koreanLabel & final"
        [void](New-Item -ItemType Directory -Path $inputDirectory -Force)
        $sourcePath = Join-Path $inputDirectory "$koreanLabel & final.doc"
        Restore-DocFixture -Base64Name "special-name.doc.base64" -Destination $sourcePath
        $sourceHash = Get-Sha256 -Path $sourcePath

        $docxOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File $script:DocxConverter -Path $sourcePath 2>&1
        $docxExit = $LASTEXITCODE
        $docxText = $docxOutput -join [Environment]::NewLine

        $docxExit | Should -Be 0 -Because $docxText
        $docxPath = [System.IO.Path]::ChangeExtension($sourcePath, ".docx")
        Test-Path -LiteralPath $docxPath | Should -BeTrue
        $docxArchive = [System.IO.Compression.ZipFile]::OpenRead($docxPath)
        try {
            @($docxArchive.Entries | ForEach-Object FullName) |
                Should -Contain "[Content_Types].xml"
            @($docxArchive.Entries | ForEach-Object FullName) |
                Should -Contain "word/document.xml"
        }
        finally {
            $docxArchive.Dispose()
        }

        $markdownOutput = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File $script:MarkdownConverter -Path $sourcePath -KordocCliPath $script:LocalKordoc 2>&1
        $markdownExit = $LASTEXITCODE
        $markdownText = $markdownOutput -join [Environment]::NewLine

        $markdownExit | Should -Be 0 -Because $markdownText
        $markdownPath = [System.IO.Path]::ChangeExtension($sourcePath, ".md")
        Test-Path -LiteralPath $markdownPath | Should -BeTrue
        (Get-Content -LiteralPath $markdownPath -Raw) |
            Should -Match "synthetic binary DOC fixture"
        (Get-Sha256 -Path $sourcePath) | Should -BeExactly $sourceHash
        @(Get-ChildItem -LiteralPath $inputDirectory -Force |
            Where-Object { $_.Name -match "^\..+\.(docx|md|backup|rollback)" }).Count |
            Should -Be 0
    }

    It "returns a typed failure for corrupt input" {
        $sourcePath = Join-Path $TestDrive "corrupt.doc"
        Restore-DocFixture -Base64Name "corrupt.doc.base64" -Destination $sourcePath
        $sourceHash = Get-Sha256 -Path $sourcePath

        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File $script:DocxConverter -Path $sourcePath 2>&1
        $exitCode = $LASTEXITCODE
        $text = $output -join [Environment]::NewLine

        $exitCode | Should -Not -Be 0
        $text | Should -Match "Status"
        $text | Should -Match "Failed"
        (Get-Sha256 -Path $sourcePath) | Should -BeExactly $sourceHash
        (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($sourcePath, ".docx"))) |
            Should -BeFalse
    }

    It "returns a typed failure for a wrong password without changing the source" {
        $sourcePath = Join-Path $TestDrive "password-protected.doc"
        Restore-DocFixture -Base64Name "password-protected.doc.base64" -Destination $sourcePath
        $sourceHash = Get-Sha256 -Path $sourcePath

        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File $script:DocxConverter -Path $sourcePath -Password "incorrect-password" 2>&1
        $exitCode = $LASTEXITCODE
        $text = $output -join [Environment]::NewLine

        $exitCode | Should -Not -Be 0
        $text | Should -Match "Status"
        $text | Should -Match "Failed"
        (Get-Sha256 -Path $sourcePath) | Should -BeExactly $sourceHash
        (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($sourcePath, ".docx"))) |
            Should -BeFalse
    }

    It "fails before Word starts when the LogPath parent is absent" {
        $sourcePath = Join-Path $TestDrive "missing-log-parent.doc"
        Restore-DocFixture -Base64Name "special-name.doc.base64" -Destination $sourcePath
        $missingLog = Join-Path $TestDrive "absent-log-parent\conversion.log"

        $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
            -File $script:DocxConverter -Path $sourcePath -LogPath $missingLog 2>&1
        $exitCode = $LASTEXITCODE
        $text = $output -join [Environment]::NewLine

        $exitCode | Should -Not -Be 0
        $text | Should -Match "Status"
        $text | Should -Match "Failed"
        $text | Should -Match "LogPath"
        (Test-Path -LiteralPath ([System.IO.Path]::ChangeExtension($sourcePath, ".docx"))) |
            Should -BeFalse
    }
}
