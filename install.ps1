<#
.SYNOPSIS
  Install `academic-report-docx` skill into Claude Code, OpenAI Codex, dsh (DeepSeek Harness), ZCode, WorkBuddy, and Gemini/Antigravity.

.EXAMPLE
  .\install.ps1
  .\install.ps1 -Target claude,codex,dsh,zcode,workbuddy
#>
param(
    [string[]]$Target = @("all")
)

$ErrorActionPreference = "Stop"
$SkillName = "academic-report-docx"
$RepoRoot = Split-Path -Parent $MyInvocation.MyCommand.Path

$AllTargets = [ordered]@{
    "claude"    = @(Join-Path $HOME ".claude\skills\$SkillName")
    "codex"     = @(
        (Join-Path $HOME ".codex\skills\$SkillName"),
        (Join-Path $HOME ".agents\skills\$SkillName")
    )
    "dsh"       = @(Join-Path $HOME ".dsh\skills\$SkillName")
    "zcode"     = @(Join-Path $HOME ".zcode\skills\$SkillName")
    "workbuddy" = @(Join-Path $HOME ".workbuddy\skills\$SkillName")
    "gemini"    = @(Join-Path $HOME ".gemini\config\skills\$SkillName")
}

# If ~/.workbuddy-ai exists on this machine, mirror to it as well under workbuddy target
if (Test-Path (Join-Path $HOME ".workbuddy-ai")) {
    $AllTargets["workbuddy"] += (Join-Path $HOME ".workbuddy-ai\skills\$SkillName")
}

$SelectedKeys = if ($Target -contains "all") {
    $AllTargets.Keys
} else {
    $Target | ForEach-Object { $_.ToLower().Trim() }
}

$BundleItems = @("SKILL.md", "agents", "configs", "scripts", "references", "examples", "assets")

foreach ($key in $SelectedKeys) {
    if (-not $AllTargets.Contains($key)) {
        Write-Warning "Unknown target '$key' (supported: all, claude, codex, dsh, zcode, workbuddy, gemini)"
        continue
    }
    foreach ($dest in $AllTargets[$key]) {
        New-Item -ItemType Directory -Force -Path $dest | Out-Null
        foreach ($item in $BundleItems) {
            $srcPath = Join-Path $RepoRoot $item
            if (Test-Path $srcPath) {
                Copy-Item -Recurse -Force -Path $srcPath -Destination $dest
            }
        }
        Write-Host "[OK] Installed $SkillName -> $dest ($key)" -ForegroundColor Green
    }
}
