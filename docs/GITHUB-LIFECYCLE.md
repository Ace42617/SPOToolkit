# GitHub lifecycle: development, testing, and publishing

A simple way to think about working with this repo: **edit locally → test → commit → push**.

---

## 1. Development (on your machine)

- You edit code in **Cursor** (or any editor) in `c:\Users\Alex\SharePointCSVExport`.
- Those files are your **working copy**. Nothing is on GitHub until you **commit** and **push**.

**Git only tracks changes when you tell it to** (add + commit). Saving a file in the editor does **not** update GitHub.

---

## 2. The basic loop: add → commit → push

| Step | What you do | What it means |
|------|-------------|---------------|
| **Edit** | Change files (e.g. `popup.js`, `exportCSV.js`) | Changes exist only on your PC. |
| **Stage** | `git add .` (or specific files) | Marks which changes you want in the next “save point”. |
| **Commit** | `git commit -m "Short description"` | Creates a save point **locally** with a message. |
| **Push** | `git push origin main` | Sends your commits to GitHub so the repo on the web is up to date. |

After a successful push, **github.com/Ace42617/SPOToolkit** shows the same history as your local repo.

---

## 3. Testing before you push

- **Load the extension** in Chrome/Edge (e.g. **Load unpacked** → this folder) and try your changes.
- Fix bugs and repeat: edit → save → reload extension → test.
- When you’re happy, **then** run: `git add .` → `git commit -m "..."` → `git push origin main`.

You can commit and push as often as you like (e.g. “Add button” then “Fix typo” then “Update README”). Each commit is a step in the history.

---

## 4. Publishing a “release” (optional)

- **Normal work:** You push to the **main** branch. That’s the main line of development; it’s already “published” in the sense that the code is on GitHub.
- **Releases** (e.g. “SPO Toolkit Alpha 1.0”) are **snapshots** you create when you want to offer a downloadable version (e.g. a .zip) and optional release notes.
  - You create a **tag** (e.g. `v1.0.0-alpha`) and then **create a release** from that tag and attach the zip.
  - Day-to-day fixes and features usually go to **main**; you only create a new release when you want to ship a new version to users.

So:

- **Development/testing:** Edit → test in browser → commit → push to **main**.
- **Publishing a version:** When ready, tag, build the zip, and create a release (as in `release/CREATE-RELEASE-INSTRUCTIONS.md`).

---

## 5. Quick command reference

From a terminal in the project folder (`c:\Users\Alex\SharePointCSVExport`):

```powershell
# See what files you changed
git status

# Stage all changes
git add .

# Create a commit with a message
git commit -m "Describe what you did in a few words"

# Send commits to GitHub
git push origin main
```

If you get “Updates were rejected” or “pull first”, run:

```powershell
git pull origin main
```

Then fix any conflicts (if Git says so), and run `git push origin main` again.

---

## 6. Summary

| Phase | Where it happens | When you’re “done” |
|-------|------------------|--------------------|
| **Development** | Your PC (this folder) | You’re satisfied with the code. |
| **Testing** | Browser (Load unpacked extension) | Behavior looks good. |
| **Publishing to GitHub** | Terminal: `add` → `commit` → `push` | `main` on GitHub matches your PC. |
| **Publishing a release** | GitHub Releases (tag + zip + notes) | When you want to ship a version to users. |

You’re “publishing changes to GitHub” every time you **push** to **main**. Releases are an extra, optional step for shipping a specific version (e.g. Alpha 1.0) with a zip and notes.
