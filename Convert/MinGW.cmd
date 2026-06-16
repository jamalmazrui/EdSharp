@echo off
rem MinGW.cmd -- compile a C++ source with MinGW g++ for EdSharp's Compiler menu.
rem   %1 = full path to the source WITHOUT extension (so "%~1.cpp" / "%~1.exe")
rem Compiler messages go to stdout/stderr; EdSharp captures them (the .ini line
rem appends 2>&1) and parses errors with its regex. Set MINGW_BIN to point at a
rem different toolchain; it defaults to C:\MinGW\bin.
setlocal
if not defined MINGW_BIN set "MINGW_BIN=C:\MinGW\bin"
set "PATH=%MINGW_BIN%;%PATH%"
set "base=%~1"
if exist "%base%.o" del /f /q "%base%.o"
g++.exe -c "%base%.cpp" -o "%base%.o"
if not exist "%base%.o" goto :done
g++.exe -o "%base%.exe" "%base%.o"
:done
endlocal
