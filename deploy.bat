@echo off
REM One-command deploy for Windows: pull latest from git, push to Apps Script,
REM update the live deployment.
REM
REM Usage:  deploy                # deploys with auto-timestamp label
REM         deploy "my notes"     # deploys with custom label
REM
REM Requires clasp installed + logged in (`clasp login` once).
REM .clasp.json at the repo root tells clasp where to send the code.
REM
REM DEPLOYMENT_ID is Amanda's active web-app deployment. Passing it to
REM `clasp deploy` UPDATES that deployment instead of creating a new one,
REM which:
REM   1) keeps the /exec URL stable — everyone using it stays working, and
REM   2) avoids hitting the 20-deployment cap Apps Script enforces.

setlocal
set "DEPLOYMENT_ID=AKfycbz4ySRzlSamI9bkgnEjlQKh1ypawUX5I25D8I9j1Uy4aSJNpa7tRDrYPgczVcpJypFLGw"

set "LABEL=%~1"
if "%LABEL%"=="" (
  for /f "tokens=1-4 delims=/ " %%a in ("%date%") do set "TODAY=%%c-%%a-%%b"
  set "LABEL=Auto-deploy %TODAY% %time:~0,5%"
)

echo ==^> Pulling latest from git...
git pull --ff-only
if errorlevel 1 (
  echo Git pull failed. Fix conflicts or use "git status" to see what happened.
  exit /b 1
)

echo ==^> Pushing code to Apps Script...
call clasp push -f
if errorlevel 1 (
  echo clasp push failed. Screenshot the error and send it.
  exit /b 1
)

echo ==^> Updating deployment: %LABEL%
call clasp deploy --deploymentId %DEPLOYMENT_ID% --description "%LABEL%"
if errorlevel 1 (
  echo clasp deploy failed.
  exit /b 1
)

echo.
echo Done. Hard-refresh the dashboard to see the new code.
endlocal
