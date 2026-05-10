try {
    $app = [System.Runtime.InteropServices.Marshal]::GetActiveObject("MSProject.Application")
    Write-Host "Found open MSProject instance"
    foreach ($p in $app.Projects) {
        Write-Host ("File: " + $p.FullName + " | Tasks: " + $p.Tasks.Count + " | Name: " + $p.Name)
    }
} catch {
    Write-Host ("Not running: " + $_.Exception.Message)
}
