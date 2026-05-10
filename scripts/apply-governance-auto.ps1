# Complete auto-import-and-run for governance views
# No user interaction needed

$basPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\scripts\governance-simple.bas"
$draftPath = "C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\AMETEK SAP S4 Master Project Plan May Full Replan Draft.mpp"

Write-Host "Connecting to MS Project..." -ForegroundColor Cyan

# Get or create MS Project instance
try {
    $msProj = [Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
    Write-Host "Connected to running MS Project" -ForegroundColor Green
} catch {
    Write-Host "Launching MS Project..." -ForegroundColor Yellow
    $msProj = New-Object -ComObject MSProject.Application
    Start-Sleep -Milliseconds 3000
}

# Make visible
$msProj.Visible = $true
Start-Sleep -Milliseconds 1000

# Check for active project
$activeProj = $msProj.ActiveProject
if ($activeProj -ne $null) {
    Write-Host "Project open: $($activeProj.Name)" -ForegroundColor Green
} else {
    Write-Host "Opening draft plan..." -ForegroundColor Yellow
    try {
        # Try to open the file with detailed error handling
        $result = $msProj.FileOpen($draftPath)
        Write-Host "FileOpen returned: $result" -ForegroundColor Cyan
        
        # Wait longer for file to open
        Start-Sleep -Milliseconds 15000
        
        # Try to get active project multiple times
        $attempts = 0
        while ($attempts -lt 5) {
            $activeProj = $msProj.ActiveProject
            if ($activeProj -ne $null) {
                break
            }
            $attempts++
            Write-Host "Waiting for project... attempt $attempts/5" -ForegroundColor Yellow
            Start-Sleep -Milliseconds 2000
        }
        
        if ($activeProj -ne $null) {
            Write-Host "Draft plan opened: $($activeProj.Name)" -ForegroundColor Green
        } else {
            Write-Host "ERROR: File opened but project not accessible" -ForegroundColor Red
            Write-Host "Please ensure the draft plan is open in MS Project and try again" -ForegroundColor Yellow
            exit 1
        }
    } catch {
        Write-Host "ERROR opening file: $_" -ForegroundColor Red
        exit 1
    }
}

# Import VBA module
Write-Host "Importing governance module..." -ForegroundColor Cyan
try {
    $vbProj = $msProj.VBE.ActiveVBProject
    $codModule = $vbProj.VBComponents.Import($basPath)
    Write-Host "Module imported successfully: $($codModule.Name)" -ForegroundColor Green
    Start-Sleep -Milliseconds 500
} catch {
    Write-Host "ERROR: Could not import: $_" -ForegroundColor Red
    exit 1
}

# Run macro
Write-Host "Running governance views..." -ForegroundColor Cyan
try {
    $msProj.RunMacro("GovernanceSimple.RunGovernanceViews")
    Write-Host "[SUCCESS] Governance views applied!" -ForegroundColor Green
    exit 0
} catch {
    Write-Host "ERROR: Could not run macro: $_" -ForegroundColor Red
    exit 1
}
