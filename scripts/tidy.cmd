@echo off
rem tidy.cmd -- tidy a Homer Tools project: the folder and the repository,
rem in one pass. It replaces two older scripts, which asked the same
rem question of two places and could disagree about the answer.
rem
rem It carries its plan out in the same run. A stray is moved into notes,
rem which is on this disk and never in git, so nothing it takes is lost; the
rem log names every move. (--do-it, from when a first run only planned, is
rem still accepted and changes nothing.)
rem
rem     tidy                      tidy the folder and the repository
rem     tidy --no-push            tidy, committing locally; do not push
rem     tidy --folder-only        leave git alone
rem     tidy --repo-only          leave the folder alone
rem     tidy --path C:\EdSharp    another project, without changing directory
rem
rem What belongs to the project is decided by the project's own
rem <App>_setup.iss and RepoFiles.txt, so there is no list in here to maintain.
rem Empty files go, duplicates and unnamed files move into notes\, files tracked
rem that do not belong are untracked and added to .gitignore, and anything large
rem in the history is reported with the command that would remove it.
rem
rem Writes tidy.log beside this script.
setlocal
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python was not found on the PATH. Install it from python.org
    echo        or with: winget install Python.Python.3.12
    endlocal
    exit /b 1
)
python "%~dp0tidy.py" %*
set exitCode=%errorlevel%
endlocal & exit /b %exitCode%
