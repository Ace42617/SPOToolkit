# Create a GitHub release

## Prerequisites

1. Bump `manifest.json` → `version`.
2. Add `release/RELEASE_NOTES_<version>.md`.
3. Commit and push `main`.
4. `gh auth login` (one-time).

## Pack + publish

From the repo root:

```powershell
cd "C:\Temp\Git\SPOToolkit"
pwsh -File release\create-release.ps1
# Script reads version from manifest.json and prints the gh command, e.g.:
git tag "v1.1.0.0"
git push origin main
git push origin "v1.1.0.0"
gh release create "v1.1.0.0" --title "SPO Developer Toolkit 1.1.0.0" --notes-file "release\RELEASE_NOTES_1.1.0.0.md" "release\SPOToolkit-1.1.0.0.zip"
```

Ensure `levelup/` exists first (`npm run sync-levelup` if needed).
