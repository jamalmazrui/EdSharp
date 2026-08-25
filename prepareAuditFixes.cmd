@echo off
rem prepareAuditFixes.cmd -- two one-time jobs before the 23 August 2026
rem audit fixes are committed. Run once from C:\EdSharp; safe to run again.
rem   1. Marks the last commit as branch snapshotBeforeAuditFixes_20260823,
rem      the way back if any of the fixes is ever to be reverted.
rem   2. git add -f Scripts, so the JAWS scripts finally join the
rem      repository (forced, because the ignore rules would otherwise keep
rem      skipping the folder; files tracked once stay tracked).
rem Everything it does is recorded in prepareAuditFixes.log beside it.
cd /d "%~dp0"
set "log=%~dp0prepareAuditFixes.log"
echo prepareAuditFixes %date% %time% > "%log%"
echo. >> "%log%"
git branch snapshotBeforeAuditFixes_20260823 >> "%log%" 2>&1
echo Snapshot: branch snapshotBeforeAuditFixes_20260823 marks the last commit before the audit fixes. >> "%log%"
echo. >> "%log%"
git add -f Scripts >> "%log%" 2>&1
for /f %%c in ('git ls-files Scripts ^| find /c /v ""') do set "count=%%c"
if "%count%"=="1" (echo Scripts: %count% file is now tracked. >> "%log%") else (echo Scripts: %count% files are now tracked. >> "%log%")
echo. >> "%log%"
echo Done. Continue with buildEdSharp and the normal release steps. >> "%log%"
type "%log%"
exit /b 0
