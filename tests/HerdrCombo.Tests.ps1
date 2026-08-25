BeforeAll {
    $script:RepoRoot = Split-Path -Parent $PSScriptRoot
    $script:ConfigPath = Join-Path $script:RepoRoot ".omo\omo.jsonc"
    $script:LauncherPath = Join-Path $script:RepoRoot "tools\herdr\start-combo.ps1"
}

Describe "Project-local OMO routing" -Tag "Static" {
    It "routes planning and review to Sol and execution to Luna" {
        Test-Path -LiteralPath $script:ConfigPath | Should -BeTrue

        $config = Get-Content -LiteralPath $script:ConfigPath -Raw | ConvertFrom-Json
        $senpi = $config."[senpi]"

        $senpi.models."sol-high".model | Should -BeExactly "openai-codex/gpt-5.6-sol"
        $senpi.models."sol-xhigh".reasoning | Should -BeExactly "xhigh"
        $senpi.models."luna-low".model | Should -BeExactly "openai-codex/gpt-5.6-luna"
        $senpi.models."luna-xhigh".reasoning | Should -BeExactly "xhigh"

        $senpi.categories.quick.model | Should -BeExactly "luna-low"
        $senpi.categories.deep.model | Should -BeExactly "luna-xhigh"
        $senpi.categories."unspecified-low".model | Should -BeExactly "luna-xhigh"
        $senpi.categories."unspecified-high".model | Should -BeExactly "luna-xhigh"
        $senpi.categories.ultrabrain.model | Should -BeExactly "sol-xhigh"
        $senpi.categories.writing.model | Should -BeExactly "sol-high"

        $senpi.agents.explore.model | Should -BeExactly "luna-low"
        $senpi.agents.librarian.model | Should -BeExactly "luna-low"
        $senpi.agents.metis.model | Should -BeExactly "sol-high"
        $senpi.agents.momus.model | Should -BeExactly "sol-xhigh"

        $senpi.task.default_execution_mode | Should -BeExactly "in-process"
        $senpi.task.default_concurrency | Should -Be 4
        $senpi.task.max_depth | Should -Be 1
        $senpi.task.resume_children | Should -BeTrue
        $senpi.task.reattach_on_reconcile | Should -BeTrue
        $senpi.task.wait.default_ms | Should -Be 90000
    }
}

Describe "Herdr combo launcher" -Tag "Static" {
    It "rejects execution outside a Herdr-managed pane" {
        Test-Path -LiteralPath $script:LauncherPath | Should -BeTrue

        $previousHerdrEnv = $env:HERDR_ENV
        try {
            Remove-Item Env:HERDR_ENV -ErrorAction SilentlyContinue
            $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:LauncherPath -ValidateOnly 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $env:HERDR_ENV = $previousHerdrEnv
        }

        $exitCode | Should -Not -Be 0
        ($output -join [Environment]::NewLine) |
            Should -Match "Herdr-managed pane"
    }

    It "validates the checked-in combo configuration without changing layout" {
        Test-Path -LiteralPath $script:LauncherPath | Should -BeTrue

        $previousHerdrEnv = $env:HERDR_ENV
        try {
            $env:HERDR_ENV = "1"
            $output = & powershell.exe -NoProfile -ExecutionPolicy Bypass `
                -File $script:LauncherPath -ValidateOnly 2>&1
            $exitCode = $LASTEXITCODE
        }
        finally {
            $env:HERDR_ENV = $previousHerdrEnv
        }

        $exitCode | Should -Be 0 -Because ($output -join [Environment]::NewLine)
        ($output -join [Environment]::NewLine) |
            Should -Match "Herdr combo configuration is valid"
    }

    It "pre-approves the project configuration for every OMO lane" {
        $launcher = Get-Content -LiteralPath $script:LauncherPath -Raw
        ([regex]::Matches($launcher, "omo --approve --model")).Count |
            Should -Be 3
    }

    It "stages OMO commands without starting autonomous turns" {
        $launcher = Get-Content -LiteralPath $script:LauncherPath -Raw
        $launcher | Should -Not -Match "--append-system-prompt"
        ([regex]::Matches($launcher, '"pane", "send-text"')).Count |
            Should -Be 3
        ([regex]::Matches($launcher, '"pane", "run"')).Count |
            Should -Be 0
    }
}
