@echo off
rem makeTutorials.cmd -- the documents step on its own.
rem
rem buildTutorial is the command to use: it writes the documents AND speaks the
rem tutorials. This wrapper exists because makeTutorials.py is a script, and
rem every script here has a wrapper so nobody has to type the PowerShell or
rem Python invocation from memory. buildTutorial -docs does the same thing.
setlocal
rem %~dp0 IS CAPTURED BEFORE pushd. When this file is CALLed by a relative
rem path -- call "scripts\buildTutorials.cmd" -- cmd works %~dp0 out again
rem against the current folder every time it is read, so after pushd into
rem scripts it became scripts\scripts\ and the .ps1 could not be found. Taking
rem it once, first, fixes it however the file is called.
set "sHere=%~dp0"
pushd "%sHere%"
python "%sHere%makeTutorials.py" %*
set iResult=%ERRORLEVEL%
popd
endlocal & exit /b %iResult%
