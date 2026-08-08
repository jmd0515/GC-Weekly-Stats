@echo off
REM ============================================================================
REM  Commits and pushes any new "All Salons MM.DD.YY.xlsx" file(s) in this
REM  folder. Double-click after dropping a new weekly Power BI export.
REM
REM  Also detects the "committed locally but never pushed" case that has bitten
REM  us before and auto-pushes any pending commits.
REM
REM  Every git command is retried (this repo lives in OneDrive, which
REM  intermittently holds a handle on .git\index or the .xlsx mid-sync) and
REM  all output is appended to commit_weekly.log so a failure is still
REM  readable after this window closes.
REM ============================================================================

setlocal enabledelayedexpansion
cd /d "%~dp0"

set "LOG=%~dp0commit_weekly.log"
set "GITOUT=%TEMP%\gc_git_out.txt"
set RETRIES=3

echo.>>"%LOG%"
echo ============================================================>>"%LOG%"
echo Run started %DATE% %TIME%>>"%LOG%"
echo ============================================================>>"%LOG%"

REM -- Today's date (MM/DD/YYYY) for the commit message. -----------------------
REM  Was 'wmic os get localdatetime' - WMIC is deprecated and being removed
REM  from Windows 11, so ask PowerShell instead.
set "TODAY="
for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "Get-Date -Format MM/dd/yyyy"`) do set "TODAY=%%i"
if not defined TODAY set "TODAY=unknown date"

REM -- Step 0a: leftover rebase from a previous failed run? --------------------
if exist ".git\rebase-merge" goto REBASE_STUCK
if exist ".git\rebase-apply" goto REBASE_STUCK
goto NO_REBASE

:REBASE_STUCK
echo.
echo ============================================================
echo   A REBASE FROM A PREVIOUS RUN IS STILL IN PROGRESS
echo ============================================================
echo.
echo Finish it first, in this folder:
echo   git rebase --continue     ^(keep going^)
echo   git rebase --abort        ^(back out, nothing is lost^)
echo.
echo Rebase still in progress - bailed out.>>"%LOG%"
pause
exit /b 1

:NO_REBASE

REM -- Step 0b: stale git index lock (OneDrive / a killed git process) ---------
if not exist ".git\index.lock" goto NO_INDEX_LOCK
tasklist /FI "IMAGENAME eq git.exe" 2>nul | find /I "git.exe" >nul
if !errorlevel! equ 0 goto GIT_RUNNING
echo Cleaning up stale .git\index.lock ^(no git process running^).
echo Removed stale .git\index.lock>>"%LOG%"
del /f ".git\index.lock" >nul 2>&1
goto NO_INDEX_LOCK

:GIT_RUNNING
echo.
echo ERROR: another git process is running. Wait for it to finish, then
echo        re-run this script.
echo git.exe running, index.lock present - bailed out.>>"%LOG%"
pause
exit /b 1

:NO_INDEX_LOCK

REM -- Step 0c: Excel lock files (~$*.xlsx) -----------------------------------
REM  If Excel is NOT running the lock files are stale (Excel crashed or was
REM  killed) - safe to auto-delete. If Excel IS running, bail out.
set HAS_LOCK=0
for %%F in ("~$*.xlsx") do if exist "%%F" set HAS_LOCK=1
if !HAS_LOCK! equ 0 goto NO_XLS_LOCK

tasklist /FI "IMAGENAME eq EXCEL.EXE" 2>nul | find /I "EXCEL.EXE" >nul
if !errorlevel! neq 0 goto CLEAN_XLS_LOCK

echo.
echo ============================================================
echo   EXCEL IS RUNNING and has files open. Cannot proceed.
echo ============================================================
echo.
echo Lock file^(s^) detected:
for %%F in ("~$*.xlsx") do if exist "%%F" echo   %%F
echo.
echo To fix:
echo   1. Close the workbook in Excel ^(or the whole Excel app^).
echo   2. If lock files remain: Ctrl+Shift+Esc, End 'Microsoft Excel' task.
echo   3. Re-run this script.
echo.
echo Excel running with open workbooks - bailed out.>>"%LOG%"
pause
exit /b 1

:CLEAN_XLS_LOCK
echo Cleaning up stale lock files ^(Excel is not running^):
for %%F in ("~$*.xlsx") do if exist "%%F" echo   %%F& del /f "%%F" >nul 2>&1
echo.

:NO_XLS_LOCK

REM -- Step 1: fetch origin so we know if we're ahead/behind -------------------
echo Fetching origin...
call :GIT fetch origin main
if not "!GITRC!"=="0" goto FAIL_FETCH

REM -- Step 2: stage any new/modified weekly files -----------------------------
REM  Accept both naming patterns: 'All Salons MM.DD.YY.xlsx' (spaces + dots)
REM  and 'All_Salons_MM_DD_YY.xlsx' (underscores, Power BI default).
REM  Each add is guarded by 'if exist' because git add exits non-zero with
REM  "pathspec did not match any files" when a pattern matches nothing - which
REM  is the normal case for whichever naming style you didn't use this week.
if not exist "All Salons *.xlsx" goto SKIP_ADD_SPACED
call :GIT add "All Salons *.xlsx"
if not "!GITRC!"=="0" goto FAIL_ADD
:SKIP_ADD_SPACED

if not exist "All_Salons_*.xlsx" goto SKIP_ADD_USCORE
call :GIT add "All_Salons_*.xlsx"
if not "!GITRC!"=="0" goto FAIL_ADD
:SKIP_ADD_USCORE

REM -- Step 3: decide what to do based on state --------------------------------
REM  a) staged changes            -> commit + push
REM  b) nothing staged, but ahead -> push the existing commits
REM  c) nothing staged, in sync   -> nothing to do
git diff --cached --quiet
set STAGED=!errorlevel!

set AHEAD=0
for /f %%i in ('git rev-list --count origin/main..HEAD 2^>nul') do set AHEAD=%%i

if !STAGED! neq 0 goto DO_COMMIT
if !AHEAD! neq 0 goto PUSH_ONLY

echo.
echo No new files, and local matches GitHub. Nothing to do.
echo Nothing to do - no staged changes, not ahead of origin.>>"%LOG%"
pause
exit /b 0

:PUSH_ONLY
echo.
echo No new files to stage, but !AHEAD! local commit^(s^) never pushed.
echo Pushing those now...
goto PUSH

:DO_COMMIT
echo Staged files:
git diff --cached --name-only
git diff --cached --name-only >>"%LOG%"
echo.

call :GIT commit -m "Add weekly All Salons file(s) (%TODAY%)"
if not "!GITRC!"=="0" goto FAIL_COMMIT

:PUSH
REM  Rebase onto origin so the push is a fast-forward. Deliberately NOT
REM  retried: a half-finished rebase needs a human, and retrying it just
REM  produces a confusing "rebase in progress" error on top of the real one.
echo.
echo Rebasing onto origin/main...
set RETRIES=1
call :GIT pull --rebase origin main
set RETRIES=3
if not "!GITRC!"=="0" goto FAIL_REBASE

echo.
echo Pushing...
call :GIT push origin main
if not "!GITRC!"=="0" goto FAIL_PUSH

echo.
echo Done. Dashboards will rebuild in ~2 minutes at:
echo   https://jmd0515.github.io/GC-Weekly-Stats/
echo.
echo Run finished OK.>>"%LOG%"
pause
exit /b 0


REM ===========================  error exits  ==================================

:FAIL_FETCH
call :BANNER "COULD NOT REACH GITHUB"
echo The real git error is printed above and saved in the log.
echo Usual causes: no internet, VPN, or expired GitHub credentials.
goto BAIL

:FAIL_ADD
call :BANNER "COULD NOT STAGE THE WEEKLY FILE"
echo The real git error is printed above and saved in the log.
echo Usual cause: OneDrive or Excel still has the .xlsx open.
goto BAIL

:FAIL_COMMIT
call :BANNER "COMMIT FAILED - nothing was pushed"
echo The real git error is printed above and saved in the log.
echo Usual causes: OneDrive holding .git\index, or an unexpected git state.
echo Run 'git status' in this folder for details.
goto BAIL

:FAIL_REBASE
call :BANNER "REBASE FAILED - nothing was pushed"
echo The real git error is printed above and saved in the log.
echo Your commit is safe locally. Most likely Excel or OneDrive reopened a
echo staged file mid-rebase. Close Excel completely ^(Task Manager if needed^),
echo then run 'git rebase --continue' here, or just re-run this script.
goto BAIL

:FAIL_PUSH
call :BANNER "PUSH FAILED"
echo The real git error is printed above and saved in the log.
echo Your commit is safe locally - re-running this script will push it.
goto BAIL

:BAIL
echo.
echo Full log: %LOG%
echo.
pause
exit /b 1


REM ===========================  subroutines  ==================================

:BANNER
echo.
echo ============================================================
echo   %~1
echo ============================================================
echo.
echo FAILED: %~1>>"%LOG%"
exit /b 0

:GIT
REM  Runs 'git %*', echoing output to the console and the log.
REM  Retries up to %RETRIES% times, then sets GITRC to the last exit code.
echo ^> git %*>>"%LOG%"
set GIT_TRY=0
:GIT_RETRY
set /a GIT_TRY+=1
git %* >"%GITOUT%" 2>&1
set GITRC=!errorlevel!
type "%GITOUT%"
type "%GITOUT%" >>"%LOG%"
if "!GITRC!"=="0" exit /b 0
if !GIT_TRY! geq %RETRIES% exit /b !GITRC!
echo   [attempt !GIT_TRY! of %RETRIES% failed - waiting 5s and retrying]
echo   [attempt !GIT_TRY! failed, rc=!GITRC! - retrying]>>"%LOG%"
REM  ping, not timeout: 'timeout' aborts when stdin is redirected.
ping -n 6 127.0.0.1 >nul 2>&1
goto GIT_RETRY
