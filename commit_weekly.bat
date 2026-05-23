@echo off
REM Commits and pushes any new "All Salons MM.DD.YY.xlsx" file(s) in this folder.
REM Double-click after dropping a new weekly Power BI export.

cd /d "%~dp0"

REM Stage all matching weekly files (untracked + modified).
git add "All Salons *.xlsx"

REM If nothing changed, exit cleanly.
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo No new or changed All Salons files to commit.
    pause
    exit /b 0
)

REM Show what's staged.
echo Staged files:
git diff --cached --name-only
echo.

REM Commit with a generic message including today's date.
for /f "tokens=2 delims==" %%i in ('"wmic os get localdatetime /value"') do set dt=%%i
set today=%dt:~4,2%/%dt:~6,2%/%dt:~0,4%
git commit -m "Add weekly All Salons file(s) (%today%)"

REM Pull (rebase) in case the bot pushed anything, then push.
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
