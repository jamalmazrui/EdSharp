@echo off
rem pbw.cmd -- compile a PowerBASIC source with PB/Win for EdSharp's Compiler menu.
rem   %1 = source file   %2 = the .log file PB/Win writes (SourceDir\Root.log)
rem PB/Win writes its messages to the .log, not the console. EdSharp captures
rem stdout (the .ini line appends 2>&1) and parses errors, so after compiling we
rem echo the log to stdout. (The old script tried "copy %2 %3" with an empty %3,
rem which produced no captured output.) Set PBWIN_BIN / PBWIN_INC to override.
setlocal
if not defined PBWIN_BIN set "PBWIN_BIN=C:\PBWin10\bin"
if not defined PBWIN_INC set "PBWIN_INC=C:\PBWin10\WinAPI"
set "src=%~1"
set "log=%~2"
if exist "%log%" del /f /q "%log%"
"%PBWIN_BIN%\pbwin.exe" "/i%PBWIN_INC%" /l /q "%src%"
if exist "%log%" type "%log%"
endlocal
