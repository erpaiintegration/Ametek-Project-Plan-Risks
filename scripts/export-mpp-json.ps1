param(
  [Parameter(Mandatory = $true)]
  [string]$MppPath,

  [Parameter(Mandatory = $true)]
  [string]$OutJson,

  [switch]$UseActiveProject
)

$ErrorActionPreference = 'Stop'

function Format-DateIso($value) {
  if ($null -eq $value) { return $null }
  try {
    $d = [datetime]$value
    return $d.ToString('yyyy-MM-dd')
  }
  catch {
    return $null
  }
}

function Safe-Get($obj, $propName) {
  try {
    return $obj.$propName
  }
  catch {
    return $null
  }
}

function Open-MppProject($app, $sourcePath) {
  $resolved = (Resolve-Path $sourcePath).Path

  $candidates = @($resolved)

  try {
    $attrs = (Get-Item -LiteralPath $resolved).Attributes
    if (($attrs -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
      $tmpDir = Join-Path ([System.IO.Path]::GetTempPath()) 'ametek-mpp'
      if (-not (Test-Path $tmpDir)) {
        New-Item -ItemType Directory -Path $tmpDir -Force | Out-Null
      }

      $tmpPath = Join-Path $tmpDir ([System.IO.Path]::GetFileName($resolved))
      Copy-Item -LiteralPath $resolved -Destination $tmpPath -Force
      $candidates += $tmpPath
    }
  }
  catch {
    # Ignore candidate-copy failures and still attempt the original path.
  }

  $errors = @()
  foreach ($candidate in ($candidates | Select-Object -Unique)) {
    try {
      $null = $app.FileOpen($candidate)
      return $candidate
    }
    catch {
      $errors += "FileOpen($candidate): $($_.Exception.Message)"
    }

    try {
      $null = $app.FileOpenEx($candidate)
      return $candidate
    }
    catch {
      $errors += "FileOpenEx($candidate): $($_.Exception.Message)"
    }
  }

  throw "Unable to open MPP file in MS Project. Attempts: $($errors -join ' | ')"
}

function Get-ProjectPath($project) {
  $path = Safe-Get $project 'FullName'
  if ([string]::IsNullOrWhiteSpace([string]$path)) {
    return $null
  }

  return [string]$path
}

function Resolve-ProjectContext($app, $sourcePath, $preferActive) {
  if ($preferActive) {
    $project = Safe-Get $app 'ActiveProject'
    if ($null -eq $project) {
      throw 'No active MS Project file is open. Open the .mpp in Project first, then retry with -UseActiveProject.'
    }

    $projectPath = Get-ProjectPath $project
    $actualLeaf = if ($projectPath) { [System.IO.Path]::GetFileName($projectPath) } else { [string](Safe-Get $project 'Name') }
    Write-Host "Reading active project: $actualLeaf"

    return [ordered]@{
      Project = $project
      OpenedFrom = $projectPath
      OpenMode = 'active-project'
      OpenedByAutomation = $false
    }
  }

  $openedFrom = Open-MppProject -app $app -sourcePath $sourcePath
  return [ordered]@{
    Project = $app.ActiveProject
    OpenedFrom = $openedFrom
    OpenMode = 'file-open'
    OpenedByAutomation = $true
  }
}

if (-not (Test-Path $MppPath)) {
  throw "MPP file not found: $MppPath"
}

$outDir = Split-Path -Parent $OutJson
if (-not (Test-Path $outDir)) {
  New-Item -ItemType Directory -Path $outDir -Force | Out-Null
}

$app = $null
$project = $null
$openedByAutomation = $false
$openMode = $null
$openedFrom = $null

try {
  $app = New-Object -ComObject MSProject.Application
  $app.Visible = $false
  try { $app.DisplayAlerts = 0 } catch {}

  $context = Resolve-ProjectContext -app $app -sourcePath $MppPath -preferActive:$UseActiveProject
  $project = $context.Project
  $openedFrom = $context.OpenedFrom
  $openMode = $context.OpenMode
  $openedByAutomation = [bool]$context.OpenedByAutomation

  $tasks = @()

  foreach ($task in $project.Tasks) {
    if ($null -eq $task) { continue }

    $uid = Safe-Get $task 'UniqueID'
    $name = Safe-Get $task 'Name'
    if ($null -eq $uid -or [string]::IsNullOrWhiteSpace([string]$name)) { continue }

    $taskObj = [ordered]@{
      id = [string](Safe-Get $task 'ID')
      uid = [int]$uid
      name = [string]$name
      wbs = [string](Safe-Get $task 'WBS')
      outlineNumber = [string](Safe-Get $task 'OutlineNumber')
      outlineLevel = [int](Safe-Get $task 'OutlineLevel')
      summary = [bool](Safe-Get $task 'Summary')
      milestone = [bool](Safe-Get $task 'Milestone')
      critical = [bool](Safe-Get $task 'Critical')
      percentComplete = [double](Safe-Get $task 'PercentComplete')
      start = Format-DateIso (Safe-Get $task 'Start')
      finish = Format-DateIso (Safe-Get $task 'Finish')
      baselineStart = Format-DateIso (Safe-Get $task 'BaselineStart')
      baselineFinish = Format-DateIso (Safe-Get $task 'BaselineFinish')
      constraintType = [string](Safe-Get $task 'ConstraintType')
      constraintDate = Format-DateIso (Safe-Get $task 'ConstraintDate')
      totalSlack = [string](Safe-Get $task 'TotalSlack')
      freeSlack = [string](Safe-Get $task 'FreeSlack')
      predecessors = [string](Safe-Get $task 'Predecessors')
      successors = [string](Safe-Get $task 'Successors')
      resourceNames = [string](Safe-Get $task 'ResourceNames')
      notes = [string](Safe-Get $task 'Notes')
      text20 = [string](Safe-Get $task 'Text20')
    }

    $tasks += $taskObj
  }

  $payload = [ordered]@{
    exportedAt = (Get-Date).ToUniversalTime().ToString('o')
    sourceFile = (Resolve-Path $MppPath).Path
    openedFrom = $openedFrom
    openMode = $openMode
    projectName = [string](Safe-Get $project 'Name')
    taskCount = $tasks.Count
    tasks = $tasks
  }

  $json = $payload | ConvertTo-Json -Depth 8
  $json | Out-File -FilePath $OutJson -Encoding utf8
}
finally {
  if ($app -ne $null) {
    if ($openedByAutomation) {
      try { $app.FileCloseAll([int]0) } catch {}
      try { $app.Quit() } catch {}
    }
    try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch {}
  }

  [System.GC]::Collect()
  [System.GC]::WaitForPendingFinalizers()
}
