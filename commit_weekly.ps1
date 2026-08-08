# ============================================================================
#  Commits and pushes any new "All Salons MM.DD.YY.xlsx" file(s) in this
#  folder. Launched by double-clicking commit_weekly.bat.
#
#  Also detects the "committed locally but never pushed" case that has bitten
#  us before and auto-pushes any pending commits.
#
#  Why PowerShell and not the batch script this replaced:
#  cmd.exe streams a .bat from disk by byte offset, so if the file is rewritten
#  while it runs, execution resumes at a stale offset and starts executing
#  fragments of lines. This script's own job is to make git rewrite this repo -
#  'git pull --rebase' checks out commit_weekly.bat itself - which corrupted
#  the run (stray "'b' is not recognized" errors, the rebase running twice).
#  PowerShell reads the whole script into memory first, so it can't happen here.
#
#  Every git call is retried, because this repo lives in OneDrive, which
#  intermittently holds a handle on .git\index or the .xlsx mid-sync. All
#  output is appended to commit_weekly.log so a failure is still readable
#  after the window closes.
# ============================================================================

$ErrorActionPreference = 'Continue'
Set-Location -LiteralPath $PSScriptRoot

$Log      = Join-Path $PSScriptRoot 'commit_weekly.log'
$GitDir   = Join-Path $PSScriptRoot '.git'
$Retries  = 3

function Write-Log {
    param([string]$Text)
    Add-Content -LiteralPath $Log -Value $Text -Encoding utf8
}

function Say {
    param([string]$Text = '')
    Write-Host $Text
    Write-Log $Text
}

function Stop-Here {
    param([int]$Code, [string]$Banner)
    if ($Banner) {
        Say ''
        Say '============================================================'
        Say "  $Banner"
        Say '============================================================'
        Say ''
    }
    if ($Code -ne 0) {
        Say "Full log: $Log"
    }
    Write-Host ''
    Write-Host 'Press Enter to close...'
    [void](Read-Host)
    exit $Code
}

# Runs git, echoing real output to the console and the log. Retries transient
# failures. Returns the exit code of the last attempt.
function Invoke-Git {
    param([string[]]$GitArgs, [int]$Attempts = $Retries)

    Write-Log ('> git ' + ($GitArgs -join ' '))
    $code = 0
    for ($try = 1; $try -le $Attempts; $try++) {
        $out = & git @GitArgs 2>&1 | ForEach-Object { $_.ToString() }
        $code = $LASTEXITCODE
        foreach ($line in $out) { Say $line }
        if ($code -eq 0) { return 0 }
        if ($try -lt $Attempts) {
            Say "  [attempt $try of $Attempts failed (exit $code) - waiting 5s and retrying]"
            Say "  [this is usually OneDrive or Excel holding a file open]"
            Start-Sleep -Seconds 5
        }
    }
    return $code
}

Write-Log ''
Write-Log '============================================================'
Write-Log ("Run started " + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'))
Write-Log '============================================================'

# -- Step 0a: leftover rebase from a previous failed run? --------------------
if ((Test-Path (Join-Path $GitDir 'rebase-merge')) -or (Test-Path (Join-Path $GitDir 'rebase-apply'))) {
    Say 'Finish it first, in this folder:'
    Say '  git rebase --continue     (keep going)'
    Say '  git rebase --abort        (back out, nothing is lost)'
    Stop-Here 1 'A REBASE FROM A PREVIOUS RUN IS STILL IN PROGRESS'
}

# -- Step 0b: stale git index lock (OneDrive, or a killed git process) -------
$indexLock = Join-Path $GitDir 'index.lock'
if (Test-Path $indexLock) {
    if (Get-Process -Name 'git' -ErrorAction SilentlyContinue) {
        Say 'Wait for it to finish, then re-run this script.'
        Stop-Here 1 'ANOTHER GIT PROCESS IS RUNNING'
    }
    Say 'Cleaning up stale .git\index.lock (no git process running).'
    Remove-Item -LiteralPath $indexLock -Force -ErrorAction SilentlyContinue
}

# -- Step 0c: Excel lock files (~$*.xlsx) -----------------------------------
#  If Excel is NOT running the lock files are stale (Excel crashed or was
#  killed) - safe to auto-delete. If Excel IS running, bail out.
$locks = @(Get-ChildItem -LiteralPath $PSScriptRoot -Filter '~$*.xlsx' -File -Force -ErrorAction SilentlyContinue)
if ($locks.Count -gt 0) {
    if (Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue) {
        Say 'Lock file(s) detected:'
        foreach ($f in $locks) { Say "  $($f.Name)" }
        Say ''
        Say 'To fix:'
        Say '  1. Close the workbook in Excel (or the whole Excel app).'
        Say '  2. If lock files remain: Ctrl+Shift+Esc, End "Microsoft Excel" task.'
        Say '  3. Re-run this script.'
        Stop-Here 1 'EXCEL IS RUNNING and has files open. Cannot proceed.'
    }
    Say 'Cleaning up stale lock files (Excel is not running):'
    foreach ($f in $locks) {
        Say "  $($f.Name)"
        Remove-Item -LiteralPath $f.FullName -Force -ErrorAction SilentlyContinue
    }
    Say ''
}

# -- Step 1: fetch origin so we know if we're ahead/behind -------------------
Say 'Fetching origin...'
if ((Invoke-Git @('fetch', 'origin', 'main')) -ne 0) {
    Say 'The real git error is printed above and saved in the log.'
    Say 'Usual causes: no internet, VPN, or expired GitHub credentials.'
    Stop-Here 1 'COULD NOT REACH GITHUB'
}

# -- Step 2: stage any new/modified weekly files -----------------------------
#  Accept both naming patterns: 'All Salons MM.DD.YY.xlsx' (spaces + dots)
#  and 'All_Salons_MM_DD_YY.xlsx' (underscores, Power BI default).
#  Each add is guarded, because git add exits non-zero with "pathspec did not
#  match any files" when a pattern matches nothing - which is the normal case
#  for whichever naming style you didn't use this week.
foreach ($pattern in @('All Salons *.xlsx', 'All_Salons_*.xlsx')) {
    if (@(Get-ChildItem -LiteralPath $PSScriptRoot -Filter $pattern -File -ErrorAction SilentlyContinue).Count -eq 0) {
        continue
    }
    if ((Invoke-Git @('add', $pattern)) -ne 0) {
        Say 'The real git error is printed above and saved in the log.'
        Say 'Usual cause: OneDrive or Excel still has the .xlsx open.'
        Stop-Here 1 'COULD NOT STAGE THE WEEKLY FILE'
    }
}

# -- Step 3: decide what to do based on state --------------------------------
#  a) staged changes            -> commit + push
#  b) nothing staged, but ahead -> push the existing commits
#  c) nothing staged, in sync   -> nothing to do
& git diff --cached --quiet
$hasStaged = ($LASTEXITCODE -ne 0)

$ahead = 0
$aheadOut = & git rev-list --count origin/main..HEAD
if ($LASTEXITCODE -eq 0 -and $aheadOut) { $ahead = [int]($aheadOut | Select-Object -First 1) }

if ($hasStaged) {
    Say 'Staged files:'
    foreach ($line in (& git diff --cached --name-only)) { Say "  $line" }
    Say ''

    $today = Get-Date -Format 'MM/dd/yyyy'
    if ((Invoke-Git @('commit', '-m', "Add weekly All Salons file(s) ($today)")) -ne 0) {
        Say 'The real git error is printed above and saved in the log.'
        Say 'Usual causes: OneDrive holding .git\index, or an unexpected git state.'
        Say "Run 'git status' in this folder for details."
        Stop-Here 1 'COMMIT FAILED - nothing was pushed'
    }
}
elseif ($ahead -gt 0) {
    Say ''
    Say "No new files to stage, but $ahead local commit(s) never pushed."
    Say 'Pushing those now...'
}
else {
    Say ''
    Say 'No new files, and local matches GitHub. Nothing to do.'
    Stop-Here 0 ''
}

# -- Step 4: rebase onto origin so the push is a fast-forward -----------------
#  Deliberately NOT retried: a half-finished rebase needs a human, and
#  retrying it just produces a confusing "rebase in progress" error on top
#  of the real one.
Say ''
Say 'Rebasing onto origin/main...'
if ((Invoke-Git @('pull', '--rebase', 'origin', 'main') 1) -ne 0) {
    Say 'The real git error is printed above and saved in the log.'
    Say 'Your commit is safe locally. Most likely Excel or OneDrive reopened a'
    Say 'staged file mid-rebase. Close Excel completely (Task Manager if needed),'
    Say "then run 'git rebase --continue' here, or just re-run this script."
    Stop-Here 1 'REBASE FAILED - nothing was pushed'
}

# -- Step 5: push ------------------------------------------------------------
Say ''
Say 'Pushing...'
if ((Invoke-Git @('push', 'origin', 'main')) -ne 0) {
    Say 'The real git error is printed above and saved in the log.'
    Say 'Your commit is safe locally - re-running this script will push it.'
    Stop-Here 1 'PUSH FAILED'
}

Say ''
Say 'Done. Dashboards will rebuild in ~2 minutes at:'
Say '  https://jmd0515.github.io/GC-Weekly-Stats/'
Stop-Here 0 ''
