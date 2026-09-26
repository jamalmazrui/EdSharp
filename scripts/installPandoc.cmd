@echo off
rem installPandoc.cmd -- wrapper so the PowerShell script runs without typing
rem execution-policy parameters. Arguments are passed straight through.
rem Writing to EdSharp's Convert folder under Program Files needs an elevated
rem prompt when this is run by hand; the EdSharp installer runs it elevated.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0installPandoc.ps1" %*
exit /b %errorlevel%
