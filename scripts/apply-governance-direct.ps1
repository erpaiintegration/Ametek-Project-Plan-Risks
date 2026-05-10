# apply-governance-direct.ps1
# Connects to running MS Project via COM and applies all governance views directly
# WITHOUT injecting VBA — uses PowerShell COM calls for everything.

Add-Type -AssemblyName Microsoft.Office.Interop.MSProject -ErrorAction SilentlyContinue

$filePath = 'C:\Users\jsw73\OneDrive\Ametek Project Plan Risks\AMETEK SAP S4 Master Project Plan May Full Replan Draft.mpp'

# Connect to running MS Project (requires Windows PowerShell 5.1)
Write-Host "Connecting to MS Project..."
$msProj = [Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
$msProj.Visible = $true

# Check if file already open
$proj = $null
try { $proj = $msProj.ActiveProject } catch {}

if (-not $proj -or $proj.Tasks.Count -eq 0) {
    Write-Host "Opening file..."
    $msProj.FileOpen($filePath) | Out-Null
    
    # Wait until project loads (poll for up to 30s)
    $waited = 0
    do {
        Start-Sleep -Seconds 2
        $waited += 2
        try { $proj = $msProj.ActiveProject } catch {}
        if ($proj) { Write-Host "  Waiting... Tasks: $($proj.Tasks.Count)" }
    } while (($proj -eq $null -or $proj.Tasks.Count -eq 0) -and $waited -lt 30)
}

if (-not $proj -or $proj.Tasks.Count -eq 0) {
    Write-Error "Could not load project after 30s. Is there a password dialog or read-only prompt in MS Project?"
    exit 1
}

Write-Host "Project loaded: $($proj.Name)  Tasks: $($proj.Tasks.Count)"

# ============================================================
# GOVERNANCE GATE DEFINITIONS
# ============================================================
$gates = @(
    @{UID=686;  WBS=$null;       Phase="ITC-1";    Cat="TESTING";   Mandate=[datetime]"5/14/2026"; Field="Finish"; Label="ITC-1 Team Testing"},
    @{UID=688;  WBS=$null;       Phase="ITC-1";    Cat="DEFECT";    Mandate=[datetime]"5/14/2026"; Field="Finish"; Label="ITC-1 Defect Correction"},
    @{UID=689;  WBS=$null;       Phase="ITC-1";    Cat="GATE";      Mandate=[datetime]"5/14/2026"; Field="Finish"; Label="ITC-1 Gate Review"},
    @{UID=849;  WBS=$null;       Phase="ITC-2";    Cat="REFRESH";   Mandate=[datetime]"5/17/2026"; Field="Finish"; Label="QS4 ITC-2 Refresh"},
    @{UID=932;  WBS=$null;       Phase="ITC-2";    Cat="TRR";       Mandate=[datetime]"5/17/2026"; Field="Finish"; Label="TRR ITC-1 (feeds ITC-2)"},
    @{UID=1284; WBS=$null;       Phase="ITC-3";    Cat="REFRESH";   Mandate=[datetime]"6/4/2026";  Field="Finish"; Label="QS5 ITC-3 Refresh"},
    @{UID=931;  WBS=$null;       Phase="ITC-3";    Cat="TRR";       Mandate=[datetime]"6/4/2026";  Field="Finish"; Label="TRR ITC-3"},
    @{UID=0;   WBS="1.3.3.7.6"; Phase="ITC-3";    Cat="KICKOFF";   Mandate=[datetime]"7/12/2026"; Field="Start";  Label="ITC-3 Test Kickoff (MSO)"},
    @{UID=884;  WBS=$null;       Phase="UAT";      Cat="REFRESH";   Mandate=[datetime]"6/21/2026"; Field="Finish"; Label="QS4 UAT Refresh"},
    @{UID=0;   WBS="1.2.4.1";   Phase="UAT";      Cat="DATA GATE"; Mandate=[datetime]"6/21/2026"; Field="Start";  Label="PMO Auth UAT Data Loads"},
    @{UID=290;  WBS=$null;       Phase="UAT";      Cat="TESTING";   Mandate=[datetime]"8/27/2026"; Field="Finish"; Label="UAT with Roles"},
    @{UID=690;  WBS=$null;       Phase="UAT";      Cat="TRR";       Mandate=[datetime]"8/27/2026"; Field="Finish"; Label="TRR UAT"},
    @{UID=294;  WBS=$null;       Phase="GO-LIVE";  Cat="MILESTONE"; Mandate=[datetime]"10/22/2026";Field="Finish"; Label="Go-Live Oct 22"}
)

# ============================================================
# STEP 1: Stamp fields on governance gate tasks
# ============================================================
Write-Host "`nSTEP 1: Stamping governance fields..."

# Clear Flag1 on all tasks first
foreach ($t in $proj.Tasks) {
    if ($t -ne $null) {
        $t.Flag1 = $false
        $t.Number1 = 0
        $t.Text20 = ""
    }
}

$stamped = 0
$results = @()

foreach ($gate in $gates) {
    $found = $false
    foreach ($t in $proj.Tasks) {
        if ($t -eq $null) { continue }
        
        $match = $false
        if ($gate.UID -gt 0) {
            if ($t.UniqueID -eq $gate.UID) { $match = $true }
        } else {
            if ($t.WBS -eq $gate.WBS) { $match = $true }
        }
        
        if ($match) {
            $planDate = if ($gate.Field -eq "Start") { [datetime]$t.Start } else { [datetime]$t.Finish }
            $delta = [int]($planDate - $gate.Mandate).TotalDays
            
            $t.Flag1 = $true
            $t.Number1 = $delta
            
            $statusLabel = if ($delta -gt 0) { "+${delta}d LATE" } elseif ($delta -eq 0) { "On mandate" } else { "$([Math]::Abs($delta))d buffer" }
            $t.Text20 = "$($gate.Phase) | $($gate.Cat) | Mandate: $($gate.Mandate.ToString('MMM d')) | $statusLabel"
            
            $statusIcon = if ($t.PercentComplete -ge 100) { "DONE" } elseif ($delta -gt 30) { "RED  >30d LATE" } elseif ($delta -gt 0) { "AMBER $delta d late" } elseif ($delta -ge -4) { "AMBER on line" } else { "GREEN $([Math]::Abs($delta))d buf" }
            
            $results += [PSCustomObject]@{
                Phase   = $gate.Phase
                Cat     = $gate.Cat
                UID     = $t.UniqueID
                Task    = $t.Name.Substring(0, [Math]::Min(35, $t.Name.Length))
                PlanDate = $planDate.ToString("MMM d, yyyy")
                Mandate  = $gate.Mandate.ToString("MMM d, yyyy")
                Delta    = $delta
                Status   = $statusIcon
                Pct      = $t.PercentComplete
            }
            
            $stamped++
            $found = $true
            break
        }
    }
    if (-not $found) {
        Write-Warning "  NOT FOUND: $($gate.Label) (UID=$($gate.UID), WBS=$($gate.WBS))"
    }
}

Write-Host "  Stamped $stamped governance gate tasks."
Write-Host ""
$results | Format-Table -AutoSize

# ============================================================
# STEP 2: Create table "Governance CPM Check"
# ============================================================
Write-Host "STEP 2: Creating table 'Governance CPM Check'..."

# pjTask* constants
$pjLeft   = 0
$pjCenter = 1
$pjTaskWBS               = 188743739
$pjTaskName              = 188743694
$pjTaskPercentComplete   = 188743712
$pjTaskStart             = 188743715
$pjTaskFinish            = 188743711
$pjTaskBaselineTenStart  = 188744091
$pjTaskBaselineTenFinish = 188744092
$pjTaskText20            = 188743773
$pjTaskTotalSlack        = 188743718
$pjTaskConstraintType    = 188743704
$pjTaskCritical          = 188743705

try { $proj.TaskTables("Governance CPM Check").Delete() } catch {}
$tbl = $proj.TaskTables.Add("Governance CPM Check")
$tbl.ShowInMenu = $true
$tbl.LockFirstColumn = $true

$f = $tbl.TableFields.Add($pjTaskWBS, $pjLeft);            $f.Width = 15; $f.Title = "WBS"
$f = $tbl.TableFields.Add($pjTaskName, $pjLeft);           $f.Width = 40; $f.Title = "Task Name"
$f = $tbl.TableFields.Add($pjTaskPercentComplete, $pjCenter); $f.Width = 7;  $f.Title = "% Done"
$f = $tbl.TableFields.Add($pjTaskStart, $pjLeft);          $f.Width = 12; $f.Title = "Plan Start"
$f = $tbl.TableFields.Add($pjTaskFinish, $pjLeft);         $f.Width = 12; $f.Title = "Plan Finish"
$f = $tbl.TableFields.Add($pjTaskBaselineTenStart, $pjLeft);  $f.Width = 12; $f.Title = "B10 Start"
$f = $tbl.TableFields.Add($pjTaskBaselineTenFinish, $pjLeft); $f.Width = 12; $f.Title = "B10 Finish"
$f = $tbl.TableFields.Add($pjTaskText20, $pjLeft);         $f.Width = 42; $f.Title = "Mandate | Delta"
$f = $tbl.TableFields.Add($pjTaskTotalSlack, $pjCenter);   $f.Width = 9;  $f.Title = "Slack"
$f = $tbl.TableFields.Add($pjTaskConstraintType, $pjLeft); $f.Width = 12; $f.Title = "Constraint"
$f = $tbl.TableFields.Add($pjTaskCritical, $pjCenter);     $f.Width = 7;  $f.Title = "Critical"

Write-Host "  Table created."

# ============================================================
# STEP 3: Create filter "Governance Gates Only" (Flag1=Yes)
# ============================================================
Write-Host "STEP 3: Creating filter 'Governance Gates Only'..."

$pjIsTrue = 38

try { $proj.TaskFilters("Governance Gates Only").Delete() } catch {}
$flt = $proj.TaskFilters.Add("Governance Gates Only")
$flt.ShowInMenu = $true
$flt.ShowRelatedSummaryRows = $true
$fc = $flt.FilterCriteria.Add()
$fc.FieldName = "Flag1"
$fc.Test = $pjIsTrue

Write-Host "  Filter created."

# ============================================================
# STEP 4: Apply view
# ============================================================
Write-Host "STEP 4: Applying Gantt Chart with governance table + filter..."

$msProj.ViewApply("Gantt Chart")
$msProj.TableApply("Governance CPM Check")
$msProj.FilterApply("Governance Gates Only")

Write-Host "  View applied."

# ============================================================
# STEP 5: Color code task rows
# ============================================================
Write-Host "STEP 5: Applying color coding..."

$colored = 0
foreach ($t in $proj.Tasks) {
    if ($t -eq $null) { continue }
    if ($t.Flag1 -eq $true) {
        $d = [int]$t.Number1
        $t.Font.Bold = $true
        
        if ($t.PercentComplete -ge 100) {
            # Blue = done
            $t.Font.Color = [long](0 + 112*256 + 192*65536)   # RGB(192,112,0) — wait, COM uses BGR
            # MS Project Font.Color uses standard RGB integer: RGB(r,g,b) = r + g*256 + b*65536
            $t.Font.Color = 192 + (112 * 256) + (0 * 65536)   # actually let's use decimal
        } elseif ($d -gt 30) {
            $t.Font.Color = 192           # RGB(192,0,0) = dark red
        } elseif ($d -gt 0) {
            $t.Font.Color = 192 + (90*256) + (17*65536)   # RGB(192,90,17) = orange/brown? no...
            # RGB(197,90,17): r=197 g=90 b=17 => 197 + 90*256 + 17*65536
            $t.Font.Color = 197 + (90*256) + (17*65536)
        } elseif ($d -ge -4) {
            # Amber: RGB(180,100,0)
            $t.Font.Color = 180 + (100*256) + (0*65536)
        } else {
            # Green: RGB(0,130,0)
            $t.Font.Color = 0 + (130*256) + (0*65536)
        }
        $colored++
    }
}

Write-Host "  Colored $colored governance gate rows."

Write-Host ""
Write-Host "============================================================"
Write-Host "GOVERNANCE VIEWS APPLIED SUCCESSFULLY"
Write-Host "============================================================"
Write-Host "Table:  'Governance CPM Check' is now active"
Write-Host "Filter: 'Governance Gates Only' (Flag1=Yes) is now active"
Write-Host "Colors: RED=>30d late  ORANGE=1-30d late  AMBER=on line  GREEN=buffer"
Write-Host ""
Write-Host "To show all tasks:  View > Filter > [No Filter]"
Write-Host "To reset table:     View > Tables > Entry"
Write-Host "============================================================"
