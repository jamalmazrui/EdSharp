@echo off
rem THE KIT'S QUESTION (10 October 2026): HomerComponents.iss asks "state jaws <file>" or "state nvda <file>", as
rem it asks the kit's installScreenReaderSupport.cmd; this answers through -sState and -pathStateFile.
if /i "%~1"=="state" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installJawsScripts.ps1" -sState "%~2" -pathStateFile "%~3"
  exit /b %errorlevel%
)
rem installJawsScripts.cmd -- runs installJawsScripts.ps1, passing arguments
rem through, so nobody types the PowerShell execution-policy parameters.
rem No pause anywhere: the log has everything, and the installer's Results
rem box reports the outcome.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installJawsScripts.ps1" %*
exit /b %errorlevel%
