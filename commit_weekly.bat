@echo off
REM Commits and pushes any new "All Salons MM.DD.YY.xlsx" file(s) in this folder.
REM Also detects the "committed locally but never pushed" case that has bitten
REM us before and auto-pushes any pending commits.
REM Double-click after dropping a new weekly Power BI export.

cd /d "%~dp0"

REM ── Step 1: Fetch origin so we know if we're ahead/behind ────────────────
echo Fetching origin...
git fetch origin main >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo ERROR: could not reach GitHub. Check your internet connection.
    pause
    exit /b 1
)

REM ── Step 2: Stage any new/modified weekly files ──────────────────────────
git add "All Salons *.xlsx"

REM ── Step 3: Decide what to do based on state ─────────────────────────────
REM  a) staged changes → commit + push
REM  b) no staged changes BUT local ahead of origin → push existing commits
REM  c) no staged changes AND up to date → nothing to do

git diff --cached --quiet
set STAGED=%errorlevel%

for /f %%i in ('git rev-list --count origin/main..HEAD 2^>nul') do set AHEAD=%%i
if not defined AHEAD set AHEAD=0

if %STAGED% equ 0 (
    if %AHEAD% equ 0 (
        echo.
        echo No new files, and local matches GitHub. Nothing to do.
        pause
        exit /b 0
    ) else (
        echo.
        echo No new files to stage, but %AHEAD% local commit(s) never pushed.
        echo Pushing those now...
        goto PUSH
    )
)

REM Show what's staged.
echo Staged files:
git diff --cached --name-only
echo.

REM Commit with today's date in the message.
for /f "tokens=2 delims==" %%i in ('"wmic os get localdatetime /value"') do set dt=%%i
set today=%dt:~4,2%/%dt:~6,2%/%dt:~0,4%
git commit -m "Add weekly All Salons file(s) (%today%)"

:PUSH
REM Rebase against origin so our push is fast-forward.
git pull --rebase origin main
if %errorlevel% neq 0 (
    echo.
    echo ERROR: rebase failed. Resolve conflicts manually, then run 'git push'.
    pause
    exit /b 1
)

git push origin main
if %errorlevel% neq 0 (
    echo.
    echo ERROR: push failed.
    pause
    exit /b 1
)

echo.
echo Done. Dashboards will rebuild in ~2 minutes at:
echo   https://jmd0515.github.io/GC-Weekly-Stats/
echo.
pause
