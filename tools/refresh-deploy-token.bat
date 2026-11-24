@echo off
setlocal EnableDelayedExpansion

REM =====================================================================
REM  NF6 Deploy Token Refresh
REM
REM  Run this whenever GitHub Actions fails with "invalid_rapt" or
REM  "invalid_grant" (usually once a month because of your Google
REM  Workspace reauth policy).
REM
REM  What it does, in order:
REM    1. Logs out your old clasp session
REM    2. Opens browser so you can sign in with amanda.lusk@nf6capital.com
REM    3. Uploads the fresh token to GitHub as the CLASPRC_JSON secret
REM    4. Triggers a test deploy
REM    5. Prints the Actions URL so you can watch it go green
REM
REM  Total time: about 30 seconds (most of it is the Google sign-in).
REM =====================================================================

echo.
echo ===============================================
echo  NF6 Deploy Token Refresh
echo ===============================================
echo.

REM --- One-time setup check: GitHub CLI --------------------------------
where gh >nul 2>&1
if errorlevel 1 (
    echo [FIRST-TIME SETUP] The GitHub CLI isn't installed yet.
    echo.
    echo Please run this one command in an admin Command Prompt:
    echo.
    echo     winget install GitHub.cli
    echo.
    echo Then close this window, open a NEW Command Prompt, and run
    echo this script again.
    echo.
    echo Alternative: download from https://cli.github.com and install.
    echo.
    pause
    exit /b 1
)

REM --- One-time setup check: gh authenticated --------------------------
gh auth status >nul 2>&1
if errorlevel 1 (
    echo [FIRST-TIME SETUP] GitHub CLI is installed but not signed in.
    echo.
    echo Running 'gh auth login' now - follow the prompts:
    echo   - Pick GitHub.com
    echo   - Pick HTTPS
    echo   - Pick 'Login with a web browser'
    echo   - Copy the one-time code, press Enter, paste it in the browser
    echo.
    pause
    gh auth login
    if errorlevel 1 (
        echo.
        echo gh auth login did not complete. Re-run this script when ready.
        pause
        exit /b 1
    )
)

REM --- One-time setup check: clasp installed ---------------------------
where clasp >nul 2>&1
if errorlevel 1 (
    echo [FIRST-TIME SETUP] clasp isn't installed. Install Node from
    echo https://nodejs.org, then run:
    echo.
    echo     npm install -g @google/clasp
    echo.
    pause
    exit /b 1
)

REM --- Step 1: Log out old clasp session -------------------------------
echo Step 1/4: Logging out old clasp session...
call clasp logout >nul 2>&1
echo   done.
echo.

REM --- Step 2: Log back in ---------------------------------------------
echo Step 2/4: Opening browser to sign in...
echo (Use amanda.lusk@nf6capital.com - the account that owns the Apps Script)
echo.
call clasp login
if errorlevel 1 (
    echo.
    echo ERROR: clasp login did not complete. See the message above.
    pause
    exit /b 1
)
if not exist "%USERPROFILE%\.clasprc.json" (
    echo.
    echo ERROR: clasp login finished but the token file wasn't created.
    echo Expected at: %USERPROFILE%\.clasprc.json
    pause
    exit /b 1
)
echo.
echo   done. Fresh token saved locally.
echo.

REM --- Step 3: Upload the token to GitHub ------------------------------
echo Step 3/4: Uploading fresh token to GitHub secret CLASPRC_JSON...
gh secret set CLASPRC_JSON --repo amandalusk-hub/NF6 < "%USERPROFILE%\.clasprc.json"
if errorlevel 1 (
    echo.
    echo ERROR: Couldn't upload the secret. Make sure gh is still signed
    echo in: gh auth status
    pause
    exit /b 1
)
echo   done.
echo.

REM --- Step 4: Trigger a test deploy -----------------------------------
echo Step 4/4: Triggering a test deploy on the working branch...
gh workflow run "Auto-deploy to Apps Script" --repo amandalusk-hub/NF6 --ref claude/debug-recurring-error-m6tVm >nul 2>&1
if errorlevel 1 (
    echo   Couldn't auto-trigger a workflow run - that's OK. The next commit
    echo   from Claude will deploy automatically.
) else (
    echo   deploy kicked off.
)
echo.

echo ===============================================
echo  All done!
echo ===============================================
echo.
echo Watch the deploy go green at:
echo   https://github.com/amandalusk-hub/NF6/actions
echo.
echo After it shows a green check, hard-refresh the web app
echo (Ctrl+Shift+R) and you're on the latest version.
echo.
pause
