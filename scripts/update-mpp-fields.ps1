param(
  [Parameter(Mandatory = $true)]
  [string]$MppPath,

  [string]$SupplementalCsv = "imports/staging/Project plan with resources and busines.csv",

  [switch]$UseActiveProject,

  [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

function Resolve-AbsolutePath([string]$pathValue) {
  if ([System.IO.Path]::IsPathRooted($pathValue)) {
    return $pathValue
  }

  return Join-Path (Get-Location) $pathValue
}

function Parse-DelimitedRow([string]$line) {
  if ($null -eq $line) { return @() }

  $values = New-Object System.Collections.Generic.List[string]
  $current = New-Object System.Text.StringBuilder
  $inQuotes = $false

  for ($i = 0; $i -lt $line.Length; $i++) {
    $ch = $line[$i]

    if ($ch -eq '"') {
      if ($inQuotes -and $i + 1 -lt $line.Length -and $line[$i + 1] -eq '"') {
        [void]$current.Append('"')
        $i++
      }
      else {
        $inQuotes = -not $inQuotes
      }
      continue
    }

    if ($ch -eq ',' -and -not $inQuotes) {
      $values.Add($current.ToString())
      $null = $current.Clear()
      continue
    }

    [void]$current.Append($ch)
  }

  $values.Add($current.ToString())
  return $values.ToArray()
}

function Parse-NullableInt([string]$value) {
  if ([string]::IsNullOrWhiteSpace($value)) { return $null }
  $tmp = 0
  if ([int]::TryParse($value.Trim(), [ref]$tmp)) {
    return $tmp
  }
  return $null
}

function Parse-SupplementalRows([string]$csvPath) {
  $lines = Get-Content -LiteralPath $csvPath
  if ($lines.Count -lt 3) {
    throw "Supplemental CSV appears empty: $csvPath"
  }

  $headerIndex = -1
  for ($i = 0; $i -lt $lines.Count; $i++) {
    $row = Parse-DelimitedRow $lines[$i]
    if ($row -contains 'Unique ID' -and $row -contains 'Task description') {
      $headerIndex = $i
      break
    }
  }

  if ($headerIndex -lt 0) {
    throw "Could not find supplemental header row with 'Unique ID' and 'Task description'."
  }

  $header = Parse-DelimitedRow $lines[$headerIndex]

  $idxUnique = [Array]::IndexOf($header, 'Unique ID')
  $idxTask = [Array]::IndexOf($header, 'Task description')
  $idxResources = [Array]::IndexOf($header, 'Resources')
  $idxBusinessValidation = [Array]::IndexOf($header, 'Business Validation')
  $idxWorkstream = [Array]::IndexOf($header, 'Workstream')

  if ($idxUnique -lt 0) {
    throw "Supplemental CSV header missing 'Unique ID'."
  }

  $map = @{}

  for ($i = $headerIndex + 2; $i -lt $lines.Count; $i++) {
    $row = Parse-DelimitedRow $lines[$i]
    if ($row.Length -eq 0) { continue }

    $uid = Parse-NullableInt ($row[$idxUnique])
    if ($null -eq $uid) { continue }

    $entry = [ordered]@{
      uid = $uid
      taskDescription = if ($idxTask -ge 0 -and $idxTask -lt $row.Length) { [string]$row[$idxTask] } else { '' }
      resourceNames = if ($idxResources -ge 0 -and $idxResources -lt $row.Length) { [string]$row[$idxResources] } else { '' }
      businessValidation = if ($idxBusinessValidation -ge 0 -and $idxBusinessValidation -lt $row.Length) { [string]$row[$idxBusinessValidation] } else { '' }
      workstream = if ($idxWorkstream -ge 0 -and $idxWorkstream -lt $row.Length) { [string]$row[$idxWorkstream] } else { '' }
    }

    $map[$uid] = $entry
  }

  return $map
}

function Try-FieldConstant($app, [string[]]$aliases) {
  foreach ($name in $aliases) {
    try {
      $constant = $app.FieldNameToFieldConstant($name)
      if ($null -ne $constant -and [int]$constant -ne 0) {
        return [ordered]@{ name = $name; constant = [int]$constant }
      }
    }
    catch {
      continue
    }
  }

  return $null
}

function Invoke-ComRetry {
  param(
    [Parameter(Mandatory = $true)]
    [scriptblock]$Action,

    [int]$MaxAttempts = 10,

    [int]$DelayMs = 750
  )

  for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
    try {
      return & $Action
    }
    catch {
      $msg = [string]$_.Exception.Message
      $isRetryable = $msg -like '*Call was rejected by callee*' -or $msg -like '*RPC_E_CALL_REJECTED*' -or $msg -like '*The remote procedure call failed*'
      if (-not $isRetryable -or $attempt -eq $MaxAttempts) {
        throw
      }

      Start-Sleep -Milliseconds $DelayMs
    }
  }
}

function Get-ProjectPath($project) {
  try {
    $path = $project.FullName
    if ([string]::IsNullOrWhiteSpace([string]$path)) {
      return $null
    }

    return [string]$path
  }
  catch {
    return $null
  }
}

function Resolve-ProjectContext($app, [string]$sourcePath, [bool]$preferActive) {
  if ($preferActive) {
    $project = Invoke-ComRetry { $app.ActiveProject }
    if ($null -eq $project) {
      throw 'No active MS Project file is open. Open the target .mpp in Project, then retry with -UseActiveProject.'
    }

    $expectedLeaf = [System.IO.Path]::GetFileName($sourcePath)
    $projectPath = Get-ProjectPath $project
    $actualLeaf = if ($projectPath) { [System.IO.Path]::GetFileName($projectPath) } else { [string]$project.Name }

    if (-not [string]::Equals($expectedLeaf, $actualLeaf, [System.StringComparison]::OrdinalIgnoreCase)) {
      throw "Active Project file '$actualLeaf' does not match expected '$expectedLeaf'."
    }

    return [ordered]@{
      Project = $project
      OpenedByAutomation = $false
      OpenMode = 'active-project'
      OpenedFrom = $projectPath
    }
  }

  $opened = $false
  try {
    Invoke-ComRetry { $null = $app.FileOpen($sourcePath) } | Out-Null
    $opened = $true
  }
  catch {
    Invoke-ComRetry { $null = $app.FileOpenEx($sourcePath) } | Out-Null
    $opened = $true
  }

  if (-not $opened) {
    throw "MS Project did not report successful open for '$sourcePath'."
  }

  $openedProject = Invoke-ComRetry { $app.ActiveProject }
  if ($null -eq $openedProject) {
    $projectsCount = 0
    try { $projectsCount = [int]$app.Projects.Count } catch { $projectsCount = 0 }
    if ($projectsCount -gt 0) {
      try { $openedProject = $app.Projects.Item(1) } catch { $openedProject = $null }
    }
  }

  return [ordered]@{
    Project = $openedProject
    OpenedByAutomation = $true
    OpenMode = 'file-open'
    OpenedFrom = $sourcePath
  }
}

$mppFull = Resolve-AbsolutePath $MppPath
$suppFull = Resolve-AbsolutePath $SupplementalCsv

if (-not (Test-Path -LiteralPath $mppFull)) {
  throw "MPP file not found: $mppFull"
}

if (-not (Test-Path -LiteralPath $suppFull)) {
  throw "Supplemental CSV not found: $suppFull"
}

$supplementalByUid = Parse-SupplementalRows -csvPath $suppFull
if ($supplementalByUid.Count -eq 0) {
  throw "No supplemental rows with Unique ID were found in: $suppFull"
}

Write-Host "Loaded supplemental rows: $($supplementalByUid.Count)"

$app = $null
$opened = $false
$project = $null
$openedByAutomation = $false
$openMode = $null
$openedFrom = $null

try {
  $app = New-Object -ComObject MSProject.Application
  try { $app.Visible = $false } catch {}
  try { $app.DisplayAlerts = 0 } catch {}

  $context = Resolve-ProjectContext -app $app -sourcePath $mppFull -preferActive:$UseActiveProject
  $project = $context.Project
  $openedByAutomation = [bool]$context.OpenedByAutomation
  $openMode = [string]$context.OpenMode
  $openedFrom = [string]$context.OpenedFrom
  $opened = $openedByAutomation

  if ($null -eq $project) {
    throw "MS Project did not return an active project after opening file."
  }

  Write-Host "Open mode: $openMode"
  if (-not [string]::IsNullOrWhiteSpace($openedFrom)) {
    Write-Host "Opened from: $openedFrom"
  }

  $resourceField = Try-FieldConstant -app $app -aliases @('Resource Names')
  $workstreamField = Try-FieldConstant -app $app -aliases @('Workstream', 'Text1')
  $businessField = Try-FieldConstant -app $app -aliases @('Business Validation Owner', 'Business Validation', 'Text2')

  if ($null -eq $resourceField) {
    throw "Could not resolve field constant for 'Resource Names'."
  }

  Write-Host "Field map:"
  Write-Host "  Resource Names => $($resourceField.name) [$($resourceField.constant)]"
  if ($workstreamField) { Write-Host "  Workstream      => $($workstreamField.name) [$($workstreamField.constant)]" } else { Write-Host "  Workstream      => NOT FOUND (skipped)" }
  if ($businessField) { Write-Host "  Business Val.   => $($businessField.name) [$($businessField.constant)]" } else { Write-Host "  Business Val.   => NOT FOUND (skipped)" }

  $stats = [ordered]@{
    matchedTasks = 0
    resourceUpdated = 0
    workstreamUpdated = 0
    businessUpdated = 0
    skippedNoSupplemental = 0
  }

  foreach ($task in $project.Tasks) {
    if ($null -eq $task) { continue }

    $uidValue = $null
    try { $uidValue = [int]$task.UniqueID } catch { $uidValue = $null }
    if ($null -eq $uidValue) { continue }

    if (-not $supplementalByUid.ContainsKey($uidValue)) {
      $stats.skippedNoSupplemental++
      continue
    }

    $supp = $supplementalByUid[$uidValue]
    $stats.matchedTasks++

    if (-not [string]::IsNullOrWhiteSpace($supp.resourceNames)) {
      if (-not $DryRun) { $null = $task.SetField($resourceField.constant, $supp.resourceNames) }
      $stats.resourceUpdated++
    }

    if ($workstreamField -and -not [string]::IsNullOrWhiteSpace($supp.workstream)) {
      if (-not $DryRun) { $null = $task.SetField($workstreamField.constant, $supp.workstream) }
      $stats.workstreamUpdated++
    }

    if ($businessField -and -not [string]::IsNullOrWhiteSpace($supp.businessValidation)) {
      if (-not $DryRun) { $null = $task.SetField($businessField.constant, $supp.businessValidation) }
      $stats.businessUpdated++
    }
  }

  if (-not $DryRun) {
    $project.Save()
  }

  Write-Host "Update complete."
  Write-Host ("Matched tasks: {0}" -f $stats.matchedTasks)
  Write-Host ("Resource Names updated: {0}" -f $stats.resourceUpdated)
  Write-Host ("Workstream updated: {0}" -f $stats.workstreamUpdated)
  Write-Host ("Business Validation updated: {0}" -f $stats.businessUpdated)
  if ($DryRun) {
    Write-Host "Dry run mode: no changes were saved."
  }
}
finally {
  if ($app -ne $null) {
    try {
      if ($openedByAutomation) {
        try { $app.FileCloseAll([int]0) } catch {}
        $app.Quit()
      }
    }
    catch {}

    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
  }

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
