@echo off
rem buildTutorials.cmd -- write the tutorials and speak them, in one command.
rem
rem   buildTutorials                    every Tutorial*.inix: documents, then audio
rem   buildTutorials Tutorial_Tagging   just that one
rem   buildTutorials -docs              documents and feed only, no speaking
rem   buildTutorials -sapi              use Windows voices; fetch nothing
rem   buildTutorials -live              perform it now through the screen reader, write no file
rem
rem The voices are Kokoro, through sherpa-onnx, when they can be fetched --
rem Apache 2.0, fine to publish under MIT -- and piper's kristin and john,
rem public domain, when they cannot. Both fetched once into scripts\voices.
rem The audio goes to help\tutorials, one .mp3 per script, with Tutorials.m3u.
rem The log is logs\<App>-tutorials-yyyyMMdd-HHmmss.log.
setlocal
rem %~dp0 IS CAPTURED BEFORE pushd. When this file is CALLed by a relative
rem path -- call "scripts\buildTutorials.cmd" -- cmd works %~dp0 out again
rem against the current folder every time it is read, so after pushd into
rem scripts it became scripts\scripts\ and the .ps1 could not be found. Taking
rem it once, first, fixes it however the file is called.
set "sHere=%~dp0"
pushd "%sHere%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%sHere%buildTutorials.ps1" %*
set iResult=%ERRORLEVEL%
popd
endlocal & exit /b %iResult%
