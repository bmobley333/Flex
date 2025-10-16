# ⚠️ WARNING! This is a destructive script. ⚠️
# It will forcibly overwrite the code on your Google Apps Script projects with your local versions.
# This is useful for reverting to old commits, but can cause data loss if used incorrectly.
# Make sure you are on the correct Git branch before running this.

# Get all subdirectories that contain a .clasp.json file
$projects = Get-ChildItem -Path . -Filter .clasp.json -Recurse | ForEach-Object { $_.Directory }

foreach ($project in $projects) {
    Write-Host ">>> Entering $($project.FullName)"
    Set-Location -Path $project.FullName  # Go into the project directory

    Write-Host ">>> FORCE Pushing $($project.Name)..."
    clasp push --force

    Write-Host "" # Add a blank line for readability
    Set-Location -Path ".."               # Go back to the root directory
}

Write-Host "✅ All projects have been force pushed."