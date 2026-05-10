# Robust version: finds the draft plan among all open projects

$basPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-simple.bas"

Write-Host "Connecting to MS Project..." -ForegroundColor Cyan

try {
    $msProj = [Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
} catch {
    Write-Host "ERROR: MS Project not running" -ForegroundColor Red
    exit 1
}

Write-Host "Connected" -ForegroundColor Green

# Try to find draft plan in open projects
$draftProj = $null

# First check active project
if ($msProj.ActiveProject -ne $null) {
    if ($msProj.ActiveProject.Name -like "*Draft*") {
        $draftProj = $msProj.ActiveProject
        Write-Host "Found draft in active project: $($draftProj.Name)" -ForegroundColor Green
    }
}

# If not found, search through all open projects
if ($draftProj -eq $null) {
    Write-Host "Searching through open projects..." -ForegroundColor Yellow
    try {
        foreach ($proj in $msProj.Projects) {
            if ($proj.Name -like "*Draft*" -or $proj.Name -like "*Replan*") {
                $draftProj = $proj
                Write-Host "Found: $($proj.Name)" -ForegroundColor Green
                # Make it active
                $msProj.Projects($proj.Name).Activate()
                Start-Sleep -Milliseconds 1000
                break
            }
        }
    } catch {
        Write-Host "Error searching projects" -ForegroundColor Yellow
    }
}

# Final fallback: just use whatever is active
if ($draftProj -eq $null) {
    $draftProj = $msProj.ActiveProject
    if ($draftProj -ne $null) {
        Write-Host "Using active project: $($draftProj.Name)" -ForegroundColor Yellow
    }
}

if ($draftProj -eq $null) {
    Write-Host "ERROR: Could not find any open project" -ForegroundColor Red
    exit 1
}

Write-Host "Working with: $($draftProj.Name)" -ForegroundColor Green

# Make visible
$msProj.Visible = $true
Start-Sleep -Milliseconds 500

# Import module
Write-Host "Importing governance module..." -ForegroundColor Cyan
try {
    $vbProj = $msProj.VBE.ActiveVBProject
    if ($vbProj -eq $null) {
        Write-Host "ERROR: Could not access VBA editor" -ForegroundColor Red
        exit 1
    }
    
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
