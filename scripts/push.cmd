@echo off
rem push.cmd -- add everything the project names, commit, push, and show the
rem status. The everyday commit of every Homer app; part of the HomerDev kit,
rem refreshed into each app's scripts folder by its build.
rem
rem     push                          commit with the default message, "Fix."
rem     push "Add the prefix field."  commit with your own
rem
rem HOW IT FITS WITH THE OTHERS:
rem   build           steps version.txt and builds; run it first, so the number
rem                   that goes up is the one in the installer.
rem   push         commits what RepoFiles.txt names -- this script.
rem   tidy       the periodic clean: puts strays in place, deletes fetched
rem                   things, rewrites the whitelist .gitignore, commits. Same
rem                   whitelist as here, so the two never disagree.
rem   release      runs the checks, tags the pushed commit with the installer's
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
rem run push again.
rem
rem The log is logs\<App>-push-yyyyMMdd-HHmmss.log in the project folder.
cls
setlocal enabledelayedexpansion
rem THE PROJECT IS THE FOLDER THIS IS RUN IN -- unless that folder is the
rem project's scripts or exec folder, in which case it is the parent. Every
rem Homer tool follows this one rule, so "cd scripts" and "push" is the
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

for /f "usebackq delims=" %%i in (`powershell -NoProfile -Command "Get-Date -Format 'yyyy-MM-ddTHH:mm:ss.fffzzz'"`) do set "sIso=%%i"
> "%log%" echo %sIso% INFO  push start
>> "%log%" echo Script: %~f0
>> "%log%" echo Folder: %CD%
>> "%log%" echo Command line: %0 %*
>> "%log%" echo Message: %message%

rem FOUR KINDS OF HOMER RESOURCE (1.45.0). kind.py says whether this is an
rem app, a collection, the kit or a page. A page project that is not a
rem repository is published by post, so push has nothing to do there.
set "sKind=app"
set "sKindPy=%~dp0kind.py"
if not exist "%sKindPy%" if defined HomerDev set "sKindPy=%HomerDev%\scripts\kind.py"
if not exist "%sKindPy%" (
  set "sUp=%CD%"
  call :findKindUp
)
if not exist "%sKindPy%" for %%L in (C D E F G H I J K L M N O P Q R S T U V W X Y Z) do if exist "%%L:\HomerDev\scripts\kind.py" set "sKindPy=%%L:\HomerDev\scripts\kind.py"
if exist "%sKindPy%" for /f "usebackq delims=" %%K in (`python "%sKindPy%" "%CD%" --word 2^>nul`) do set "sKind=%%K"
>> "%log%" echo Kind: !sKind! ^(from %sKindPy%^)

git rev-parse --is-inside-work-tree >nul 2>&1
if errorlevel 1 (
  if /i "!sKind!"=="page" (
    echo This is a page project, not a repository: post publishes it, so there is nothing to push.
    echo PAGE PROJECT, not a repository: nothing to push>> "%log%"
    endlocal & exit /b 0
  )
  echo This folder is not a git repository. Run create^<App^>Repo first.
  echo NOT A REPOSITORY>> "%log%"
  endlocal & exit /b 1
)

if not exist "%CD%\RepoFiles.txt" (
  echo No RepoFiles.txt here, so nothing was staged: without it, "git add -A" would
  echo take everything in the folder. RepoFiles.txt names what the repository
  echo carries; add it and run push again. tidy makes the first commit.
  echo NO RepoFiles.txt: stopped before staging>> "%log%"
  endlocal & exit /b 1
)

rem The whitelist is rewritten on every push, so RepoFiles.txt and .gitignore
rem cannot drift apart, and a line just added to the list counts now.
rem A WHITELIST THAT CANNOT BE REWRITTEN STOPS THE PUSH (1.63.3, from an audit
rem by another AI): a stale .gitignore can publish files never meant for it.
if exist "%~dp0tidy.cmd" (
  call "%~dp0tidy.cmd" --gitignore >> "%log%" 2>&1
  if errorlevel 1 (
    echo The whitelist .gitignore could not be rewritten, so nothing was pushed. The log has why: %log%
    echo WHITELIST FAILED>> "%log%"
    endlocal & exit /b 1
  )
)

rem A LOCAL FILE STOPS BEING TRACKED (10 October 2026): a file named in LocalFiles.txt
rem is kept on this disk and out of the repository, but Git goes on staging a file it
rem already tracks even once the .gitignore leaves it out, so DbDo's 77 MB RadioTrail
rem catalog, tracked from its 14-station start, was staged again and refused. Each file
rem LocalFiles.txt names that Git still tracks is untracked here; it stays on disk, and
rem only files LocalFiles.txt names are ever untracked.
if exist "%~dp0..\LocalFiles.txt" (
  for /f "usebackq eol=# delims=" %%L in ("%~dp0..\LocalFiles.txt") do (
    for /f "delims=" %%F in ('git ls-files -- "%%L" 2^>nul') do (
      git rm --cached --quiet -- "%%F" >> "%log%" 2>&1
      echo untracked for the whitelist to decide: %%F>> "%log%"
    )
  )
)

git add -A >> "%log%" 2>&1
rem WHAT ACTUALLY LEAVES (10 October 2026): the step above untracks every file a
rem LocalFiles.txt line names, and git add -A puts back each one the whitelist still
rem includes -- 29 kit sources were announced as "no longer pushed" and were not.
rem Only a file staged for removal that is still on this disk is named now.
for /f "delims=" %%F in ('git diff --cached --name-only --diff-filter=D') do (
  if exist "%%F" (
    echo No longer pushed, kept on this disk: %%F
    echo NOW LOCAL: %%F>> "%log%"
  )
)
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
  echo Nothing was committed. Add the line to LocalFiles.txt and run push again.
  echo LARGE STAGED, commit refused: !sBig!>> "%log%"
  git reset --quiet >> "%log%" 2>&1
  endlocal & exit /b 1
)

rem NOTHING STAGED IS NOT A FAILED COMMIT (1.62.5, from an audit by another AI):
rem every failed commit was called "nothing to commit" and ended with success,
rem hiding a sign-in, hook or lock failure; and commits already made but not
rem pushed were never pushed from a clean tree. Git is asked first whether
rem anything is staged; a commit that then fails is a failure.
git diff --cached --quiet >> "%log%" 2>&1
if errorlevel 1 goto :commit
set "sAhead=0"
for /f "delims=" %%c in ('git rev-list --count @{u}..HEAD 2^>nul') do set "sAhead=%%c"
if "%sAhead%"=="0" (
  echo Nothing to commit or push.
  echo NOTHING TO COMMIT OR PUSH>> "%log%"
  git status --short --branch
  endlocal & exit /b 0
)
echo Nothing new to commit; pushing %sAhead% earlier commits.
>> "%log%" echo NOTHING STAGED; PUSHING %sAhead% EARLIER COMMITS
goto :push
:commit
git commit -m "%message%" >> "%log%" 2>&1
if errorlevel 1 (
  echo The commit failed. The log has why: %log%
  echo COMMIT FAILED>> "%log%"
  endlocal & exit /b 1
)
:push
git push >> "%log%" 2>&1
if errorlevel 1 (
  echo The push failed. The log has why: %log%
  echo PUSH FAILED>> "%log%"
  endlocal & exit /b 1
)
rem ORIGIN FOLLOWS A MOVE (1.42.1): when GitHub says the repository moved,
rem point origin at the new location once, so it stops saying so. Placed
rem after the push failure check, which must see git push's own errorlevel.
powershell -NoProfile -Command "$l = @(Get-Content -LiteralPath '%log%'); for ($i = 0; $i -lt $l.Count - 1; $i++) { if ($l[$i] -match 'This repository moved') { $u = ($l[$i + 1] -replace '^remote:\s*', '').Trim(); if ($u -match '^https://') { git remote set-url origin $u; 'Origin now points at ' + $u }; break } }" >> "%log%" 2>&1
git status --short --branch
>> "%log%" echo push finished %date% %time%
echo Pushed: %message%
endlocal & exit /b 0

:findKindUp
rem Climbs from sUp to the top of its drive (1.46.0) for the kit's kind.py:
rem in a folder that is the kit, or in a HomerDev folder beside one above.
if exist "!sUp!\scripts\kind.py" if exist "!sUp!\exec\CSharp\Lbc.cs" (set "sKindPy=!sUp!\scripts\kind.py" & goto :eof)
if exist "!sUp!\HomerDev\scripts\kind.py" (set "sKindPy=!sUp!\HomerDev\scripts\kind.py" & goto :eof)
for %%I in ("!sUp!\..") do set "sNext=%%~fI"
if /i "!sNext!"=="!sUp!" goto :eof
set "sUp=!sNext!"
goto :findKindUp
