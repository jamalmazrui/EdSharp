@echo off
rem gitPush.cmd -- add everything the project names, commit, push, and show the
rem status. The everyday commit of every Homer app; part of the HomerDev kit,
rem refreshed into each app's scripts folder by its build.
rem
rem     gitPush                          commit with the default message, "Fix."
rem     gitPush "Add the prefix field."  commit with your own
rem
rem HOW IT FITS WITH THE OTHERS:
rem   build<App>      steps version.txt and builds; run it first, so the number
rem                   that goes up is the one in the installer.
rem   gitPush         commits what RepoFiles.txt names -- this script.
rem   homerTidy       the periodic clean: puts strays in place, deletes fetched
rem                   things, rewrites the whitelist .gitignore, commits. Same
rem                   whitelist as here, so the two never disagree.
rem   tagRelease      runs the checks, tags the pushed commit with the installer's
rem                   version and publishes the release.
rem
rem WHAT "git add -A" MEANS HERE. In a Homer project .gitignore is a WHITELIST
rem written from RepoFiles.txt: everything is ignored, then exactly the named
rem files are pushed, less what LocalFiles.txt names. So "add -A" means "add
rem everything the project has named", never "everything in the folder". On
rem 25 September 2026 a project with no RepoFiles.txt swept 480 fetched voice
rem files into a commit that way. So, in order:
rem   1. No RepoFiles.txt: stop. Nothing is staged.
rem   2. The whitelist .gitignore is rewritten from RepoFiles.txt first, so a
rem      line added to the list is honoured by this very push.
rem   3. Anything staged that is larger than 10 MB stops the commit and is
rem      named: something that size is fetched, built or recorded, and
rem      belongs in LocalFiles.txt, not in the history.
rem If a file you expected does not go up, add its line to RepoFiles.txt and
rem run gitPush again.
rem
rem The log is logs\<App>-push-yyyyMMdd-HHmmss.log in the project folder.
cls
setlocal enabledelayedexpansion
rem THE PROJECT IS THE FOLDER THIS IS RUN IN -- unless that folder is the
rem project's scripts or exec folder, in which case it is the parent. Every
rem Homer tool follows this one rule, so "cd scripts" and "gitPush" is the
rem same as running it at the top. On 25 Sep 2026 a push run from scripts
rem logged to scripts\logs and looked for RepoFiles.txt in scripts.
for %%I in ("%CD%") do set "sLeaf=%%~nxI"
if /i "%sLeaf%"=="scripts" (for %%I in ("%CD%\..") do cd /d "%%~fI")
if /i "%sLeaf%"=="exec" (for %%I in ("%CD%\..") do cd /d "%%~fI")
if not exist "%CD%\logs" mkdir "%CD%\logs"
for /f %%T in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd-HHmmss"') do set "sNow=%%T"
for %%I in ("%CD%") do set "sApp=%%~nxI"
set "log=%CD%\logs\%sApp%-push-%sNow%.log"
set "message=%~1"
if "%message%"=="" set "message=Fix."

> "%log%" echo gitPush started %date% %time%
>> "%log%" echo Script: %~f0
>> "%log%" echo Folder: %CD%
>> "%log%" echo Command line: %0 %*
>> "%log%" echo Message: %message%

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  echo This folder is not a git repository. Run create^<App^>Repo first.
  echo NOT A REPOSITORY>> "%log%"
  endlocal & exit /b 1
)

if not exist "%CD%\RepoFiles.txt" (
  echo No RepoFiles.txt here, so nothing was staged: without it, "git add -A" would
  echo take everything in the folder. RepoFiles.txt names what the repository
  echo carries; add it and run gitPush again. homerTidy --do-it makes the first commit.
  echo NO RepoFiles.txt: stopped before staging>> "%log%"
  endlocal & exit /b 1
)

rem The whitelist is rewritten on every push, so RepoFiles.txt and .gitignore
rem cannot drift apart, and a line just added to the list counts now.
if exist "%~dp0homerTidy.cmd" (
  call "%~dp0homerTidy.cmd" --gitignore >> "%log%" 2>&1
  if errorlevel 1 echo WARN: the whitelist .gitignore could not be rewritten; see the log.
)

git add -A >> "%log%" 2>&1
>> "%log%" echo ---- staged ----
git diff --cached --stat >> "%log%" 2>&1

rem Anything staged larger than 10 MB stops the commit.
set "sBig="
for /f "delims=" %%F in ('git diff --cached --name-only --diff-filter=AM') do (
  if exist "%%F" if %%~zF GTR 10485760 set "sBig=!sBig! "%%F""
)
if defined sBig (
  echo Stopped: something staged is larger than 10 MB, which means it was fetched,
  echo built or recorded rather than written, and belongs in LocalFiles.txt:
  for %%F in (!sBig!) do echo    %%~F
  echo Nothing was committed. Add the line to LocalFiles.txt and run gitPush again.
  echo LARGE STAGED, commit refused: !sBig!>> "%log%"
  git reset --quiet >> "%log%" 2>&1
  endlocal & exit /b 1
)

git commit -m "%message%" >> "%log%" 2>&1
if errorlevel 1 (
  echo Nothing to commit, so nothing was pushed.
  echo NOTHING TO COMMIT>> "%log%"
  git status --short --branch
  endlocal & exit /b 0
)
git push >> "%log%" 2>&1
if errorlevel 1 (
  echo The push failed. The log has why: %log%
  echo PUSH FAILED>> "%log%"
  endlocal & exit /b 1
)
git status --short --branch
>> "%log%" echo gitPush finished %date% %time%
echo Pushed: %message%
endlocal & exit /b 0
