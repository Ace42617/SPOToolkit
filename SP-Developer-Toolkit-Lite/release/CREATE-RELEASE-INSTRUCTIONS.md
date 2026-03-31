# Create the Alpha 1.0 release on GitHub

The tag **v1.0.0-alpha** is already pushed. To publish the release (with the zip and installation instructions), use one of these options.

## Option A: GitHub CLI (recommended)

1. Open a **new** PowerShell or terminal (so it picks up `gh`).
2. Log in (one-time):
   ```powershell
   gh auth login
   ```
   Follow the prompts (browser or token).
3. From this folder (`release`), run:
   ```powershell
   cd "c:\Users\Alex\SharePointCSVExport\release"
   .\create-release.ps1
   ```
   Or run the same command from the project root:
   ```powershell
   cd "c:\Users\Alex\SharePointCSVExport"
   gh release create "v1.0.0-alpha" --title "SPO Toolkit Alpha 1.0" --notes-file "release/RELEASE_NOTES.md" "release/SPOToolkit-Alpha-1.0.zip"
   ```

## Option B: GitHub website

1. Go to **https://github.com/Ace42617/SPOToolkit/releases**
2. Click **Draft a new release**
3. **Choose a tag:** select **v1.0.0-alpha**
4. **Release title:** `SPO Toolkit Alpha 1.0`
5. **Description:** paste the contents of `release/RELEASE_NOTES.md` (includes installation instructions)
6. Attach **SPOToolkit-Alpha-1.0.zip** from the `release` folder (drag and drop or **Attach binaries**)
7. Click **Publish release**

The zip is at: `c:\Users\Alex\SharePointCSVExport\release\SPOToolkit-Alpha-1.0.zip`
