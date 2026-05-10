param(
    [string]$ConfigPath = $env:DAILY_REPORT_CONFIG,
    [string]$OutputDir = $env:DAILY_REPORT_OUTPUT_DIR
)

$repoRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
if ([string]::IsNullOrWhiteSpace($ConfigPath)) {
    # Prefer production config; fall back to example if it doesn't exist yet
    $prodConfig = Join-Path $repoRoot "config\daily-report-sources.json"
    $ConfigPath = if (Test-Path $prodConfig) { $prodConfig } else { Join-Path $repoRoot "config\daily-report-sources.example.json" }
} elseif (-not [System.IO.Path]::IsPathRooted($ConfigPath)) {
    $ConfigPath = Join-Path $repoRoot $ConfigPath
}

if (-not [string]::IsNullOrWhiteSpace($OutputDir) -and -not [System.IO.Path]::IsPathRooted($OutputDir)) {
    $OutputDir = Join-Path $repoRoot $OutputDir
}

$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
if (-not (Test-Path $python)) {
    $python = "python"
}

$scriptPath = Join-Path $repoRoot "scripts\build_daily_report.py"
$args = @($scriptPath, "--config", $ConfigPath)
if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
    $args += @("--output-dir", $OutputDir)
}

Write-Host "Running daily report build using config: $ConfigPath"
if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
    Write-Host "Output directory: $OutputDir"
}

& $python @args
exit $LASTEXITCODE
