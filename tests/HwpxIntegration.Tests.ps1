Describe "HWPX integration contracts" -Tag "Static" {
    It "rejects incomplete builds and stale page evidence" {
        $repoRoot = Split-Path -Parent $PSScriptRoot
        $output = & uv run --project $repoRoot python -X utf8 `
            (Join-Path $PSScriptRoot "test_hwpx_integration.py") 2>&1
        $LASTEXITCODE | Should -Be 0 -Because ($output -join [Environment]::NewLine)
    }
}
