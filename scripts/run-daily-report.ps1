param(
    [string]$ConfigPath = $env:DAILY_REPORT_CONFIG,
    [string]$OutputDir = $env:DAILY_REPORT_OUTPUT_DIR,
    [switch]$ForceCloseExcel = ($env:DAILY_REPORT_FORCE_CLOSE_EXCEL -in @('1', 'true', 'TRUE', 'yes', 'YES')),
    [int]$LockRetryCount = 6,
    [int]$LockRetryDelaySeconds = 5
)

function Test-FileUnlocked {
    param([string]$Path)
    if (-not (Test-Path $Path)) {
        return $true
    }

    try {
        $stream = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        $stream.Close()
        return $true
    } catch {
        return $false
    }
}

function Resolve-WorkbookPath {
    param(
        [string]$ResolvedConfigPath,
        [string]$ResolvedOutputDir,
        [string]$RepoRoot
    )

    $configJson = Get-Content -Path $ResolvedConfigPath -Raw | ConvertFrom-Json
    $effectiveOutputDir = $ResolvedOutputDir
    if ([string]::IsNullOrWhiteSpace($effectiveOutputDir)) {
        $effectiveOutputDir = $configJson.output_dir
    }

    if ([string]::IsNullOrWhiteSpace($effectiveOutputDir)) {
        $effectiveOutputDir = "outputs/daily_report"
    }

    if (-not [System.IO.Path]::IsPathRooted($effectiveOutputDir)) {
        $effectiveOutputDir = Join-Path $RepoRoot $effectiveOutputDir
    }

    return Join-Path $effectiveOutputDir "Ametek_SAP_S4_Impl_Daily_Report.xlsx"
}

function Ensure-WorkbookUnlocked {
    param(
        [string]$WorkbookPath,
        [bool]$ForceClose,
        [int]$Retries,
        [int]$DelaySeconds
    )

    $lockName = "~$" + [System.IO.Path]::GetFileName($WorkbookPath)
    $lockFile = Join-Path ([System.IO.Path]::GetDirectoryName($WorkbookPath)) $lockName

    for ($attempt = 1; $attempt -le $Retries; $attempt++) {
        if (Test-Path $lockFile) {
            Remove-Item -Path $lockFile -Force -ErrorAction SilentlyContinue
        }

        if (Test-FileUnlocked -Path $WorkbookPath) {
            if ($attempt -gt 1) {
                Write-Host "Workbook lock released on attempt $attempt."
            }
            return
        }

        if ($ForceClose -and $attempt -eq 1) {
            Write-Host "Workbook is locked. Force-closing Excel processes because ForceCloseExcel is enabled..."
            Get-Process | Where-Object { $_.ProcessName -match '^EXCEL$' } | Stop-Process -Force -ErrorAction SilentlyContinue
        }

        Write-Host "Workbook is locked (attempt $attempt/$Retries). Waiting $DelaySeconds seconds before retry..."
        Start-Sleep -Seconds $DelaySeconds
    }

    $hint = if ($ForceClose) {
        "Could not unlock workbook even after force-closing Excel."
    } else {
        "Workbook appears to be in use. Re-run with -ForceCloseExcel to automatically close Excel locks."
    }

    throw "Unable to acquire lock for workbook: $WorkbookPath. $hint"
}

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

$targetWorkbook = Resolve-WorkbookPath -ResolvedConfigPath $ConfigPath -ResolvedOutputDir $OutputDir -RepoRoot $repoRoot
Ensure-WorkbookUnlocked -WorkbookPath $targetWorkbook -ForceClose $ForceCloseExcel -Retries $LockRetryCount -DelaySeconds $LockRetryDelaySeconds

Write-Host "Running daily report build using config: $ConfigPath"
if (-not [string]::IsNullOrWhiteSpace($OutputDir)) {
    Write-Host "Output directory: $OutputDir"
}

& $python @args
exit $LASTEXITCODE
