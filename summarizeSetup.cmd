@echo off
rem summarizeSetup.cmd -- runs summarizeSetup.ps1, which shows the single
rem Results box after every finish-page checkbox has run. The PowerShell
rem parameters live here so nobody has to type them, per the house rule that
rem a .ps1 always ships with a .cmd that calls it. Nothing pauses: the box
rem is the only thing seen, and the summary file and log keep the record.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0summarizeSetup.ps1" %*
exit /b 0
