# Deploy helpers

## `refresh-deploy-token.bat`

Windows one-click script that refreshes the clasp OAuth token GitHub Actions
uses to deploy the Apps Script project. Run this whenever GitHub Actions
fails with an `invalid_rapt` or `invalid_grant` error (usually once a month,
caused by the NF6 Capital Google Workspace reauth policy).

### What it does

1. Logs out your old clasp session
2. Opens your browser so you can sign in with `amanda.lusk@nf6capital.com`
3. Uploads the fresh token to the GitHub `CLASPRC_JSON` secret via `gh`
4. Triggers a test deploy

Total time: ~30 seconds, almost all of it the Google sign-in.

### One-time setup (first run only)

The script walks you through this automatically the first time, but here's
the summary:

1. **Install GitHub CLI** — open an admin Command Prompt and run:
   ```
   winget install GitHub.cli
   ```
   (Or download the MSI from https://cli.github.com)

2. **Sign in to GitHub CLI** — the script will do this for you on first run,
   or you can do it ahead of time:
   ```
   gh auth login
   ```
   Pick GitHub.com → HTTPS → "Login with a web browser", copy the code,
   paste in the browser.

3. **Install clasp** (if you don't have it already) — this comes with the
   repo setup; needs Node from https://nodejs.org then:
   ```
   npm install -g @google/clasp
   ```

### How to run

Double-click `refresh-deploy-token.bat` from Windows Explorer. Or from
Command Prompt:

```
cd path\to\NF6\tools
refresh-deploy-token.bat
```

### Pin it to your taskbar

Right-click the `.bat` file → Send to → Desktop (create shortcut). Then drag
the desktop shortcut to the taskbar. One click, you're refreshing.
