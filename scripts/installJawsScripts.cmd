@echo off
rem installJawsScripts.cmd -- runs installJawsScripts.ps1, passing arguments
rem through, so nobody types the PowerShell execution-policy parameters.
rem No pause anywhere: the log has everything, and the installer's Results
rem box reports the outcome.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installJawsScripts.ps1" %*
exit /b %errorlevel%
