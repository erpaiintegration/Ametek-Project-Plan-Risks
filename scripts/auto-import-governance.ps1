
# Auto-import and run governance views via VBA
# Usage: powershell -ExecutionPolicy Bypass -File auto-import-governance.ps1

$basPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-simple.bas"
$draftPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\AMETEK SAP S4 Master Project Plan May Full Replan Draft.mpp"
$moduleName = "GovernanceSimple"

# Connect to MS Project
Write-Host "Connecting to MS Project..."
try {
    $msProj = [Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
} catch {
    Write-Host "ERROR: MS Project not running" -ForegroundColor Red
    exit 1
}

$proj = $msProj.ActiveProject
if (-not $proj) {
    Write-Host "No active project. Opening draft plan..."
    $msProj.FileOpen($draftPath)
    Start-Sleep -Milliseconds 2000
    $proj = $msProj.ActiveProject
}

if (-not $proj) {
    Write-Host "ERROR: Could not open project" -ForegroundColor Red
    exit 1
}

Write-Host "Project: $($proj.Name)" -ForegroundColor Green

# Import VBA module
Write-Host "Importing governance module..."
try {
    $vbProj = $msProj.VBE.ActiveVBProject
    $codModule = $vbProj.VBComponents.Import($basPath)
    Write-Host "Module imported: $($codModule.Name)" -ForegroundColor Green
} catch {
    Write-Host "ERROR importing module: $_" -ForegroundColor Red
    exit 1
}

# Run the macro
Write-Host "Running RunGovernanceViews()..." -ForegroundColor Cyan
try {
    $msProj.RunMacro("$($moduleName).RunGovernanceViews")
    Write-Host "[OK] Governance views applied successfully!" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "ERROR running macro: $_" -ForegroundColor Red
    exit 1
}
