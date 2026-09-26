@echo off
rem homerTidy.cmd -- tidy a Homer Tools project: the folder and the repository,
rem in one pass. It replaces homerTidy and homerTidy, which asked the same
rem question of two places and could disagree about the answer.
rem
rem Running it does nothing but look and report. Run it again with --do-it to
rem carry the plan out.
rem
rem     homerTidy                      survey and print the plan
rem     homerTidy --do-it              carry it out
rem     homerTidy --do-it --no-push    carry it out locally, do not push
rem     homerTidy --folder-only        leave git alone
rem     homerTidy --repo-only          leave the folder alone
rem     homerTidy --path C:\EdSharp    another project, without changing directory
rem
rem What belongs to the project is decided by the project's own
rem <App>_setup.iss and RepoFiles.txt, so there is no list in here to maintain.
rem Empty files go, duplicates and unnamed files move into notes\, files tracked
rem that do not belong are untracked and added to .gitignore, and anything large
rem in the history is reported with the command that would remove it.
rem
rem Writes homerTidy.log beside this script.
setlocal
where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python was not found on the PATH. Install it from python.org
    echo        or with: winget install Python.Python.3.12
    endlocal
    exit /b 1
)
python "%~dp0homerTidy.py" %*
set exitCode=%errorlevel%
endlocal & exit /b %exitCode%
