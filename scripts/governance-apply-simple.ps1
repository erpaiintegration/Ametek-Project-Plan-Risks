# Simple version: assumes draft plan already open in MS Project
# Just imports the module and runs it

$basPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-simple.bas"

Write-Host "Connecting to MS Project..." -ForegroundColor Cyan

try {
    $msProj = [Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
} catch {
    Write-Host "ERROR: MS Project not running or draft plan not open" -ForegroundColor Red
    exit 1
}

# Verify active project
if ($msProj.ActiveProject -eq $null) {
    Write-Host "ERROR: No active project. Please open the draft plan in MS Project first." -ForegroundColor Red
    exit 1
}

Write-Host "Found: $($msProj.ActiveProject.Name)" -ForegroundColor Green

# Make visible
$msProj.Visible = $true

# Import module
Write-Host "Importing governance module..." -ForegroundColor Cyan
try {
    $vbProj = $msProj.VBE.ActiveVBProject
    $codModule = $vbProj.VBComponents.Import($basPath)
    Write-Host "Imported: $($codModule.Name)" -ForegroundColor Green
    Start-Sleep -Milliseconds 1000
} catch {
    Write-Host "ERROR importing: $_" -ForegroundColor Red
    exit 1
}

# Run macro
Write-Host "Running governance macro..." -ForegroundColor Cyan
try {
    $msProj.RunMacro("GovernanceSimple.RunGovernanceViews")
    Write-Host "[SUCCESS] Governance views applied!" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "ERROR running macro: $_" -ForegroundColor Red
    exit 1
}
