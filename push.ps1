# Get all subdirectories that contain a .clasp.json file
$projects = Get-ChildItem -Path . -Filter .clasp.json -Recurse | ForEach-Object { $_.Directory }

foreach ($project in $projects) {
    Write-Host ">>> Entering $($project.FullName)"
    Set-Location -Path $project.FullName  # Go into the project directory

    Write-Host ">>> Pushing $($project.Name)..."
    clasp push

    Write-Host "" # Add a blank line for readability
    Set-Location -Path ".."               # Go back to the root directory
}

Write-Host "✅ All projects pushed."