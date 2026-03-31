# Upload this project to GitHub

Git wasn’t found on your machine. Use one of the options below to push this folder as a **new** GitHub repo.

## Option A: Git from command line

1. **Install Git** (if needed): https://git-scm.com/download/win  
   Then close and reopen PowerShell/terminal.

2. **Create a new repo on GitHub**  
   - Go to https://github.com/new  
   - Set repository name (e.g. `SharePoint-Toolkit`), leave it empty (no README, no .gitignore).  
   - Click **Create repository**.

3. **Push this folder** (replace `YOUR_USERNAME` and `REPO_NAME` with your GitHub username and repo name):

   ```powershell
   cd "c:\Users\Alex\SharePointCSVExport"
   git init
   git add .
   git commit -m "Initial commit: SharePoint Toolkit extension"
   git branch -M main
   git remote add origin https://github.com/YOUR_USERNAME/REPO_NAME.git
   git push -u origin main
   ```

4. If GitHub prompts for auth, use a **Personal Access Token** as the password (Settings → Developer settings → Personal access tokens).

## Option B: GitHub Desktop

1. Install **GitHub Desktop**: https://desktop.github.com/

2. **File → Add local repository** → choose `c:\Users\Alex\SharePointCSVExport`.  
   If it says “not a Git repository”, click **create a repository** and use this folder.

3. **Publish repository** → choose your GitHub account and set the repo name → **Publish**.

After the first push, you can remove this file (`GITHUB.md`) if you don’t need it anymore.
