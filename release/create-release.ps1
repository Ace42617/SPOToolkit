# Run this after: gh auth login
# Creates GitHub release "SPO Toolkit Alpha 1.0" with the zip and release notes.
$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot
gh release create "v1.0.0-alpha" --title "SPO Toolkit Alpha 1.0" --notes-file "RELEASE_NOTES.md" "SPOToolkit-Alpha-1.0.zip"
