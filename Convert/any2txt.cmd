@echo off
rem any2txt.cmd -- convert a supported document to plain text for EdSharp's Import,
rem using 2htm.exe in plain-text mode (-p).  Replaces the old GetText.exe and the
rem Office COM converters (WdVert/XlVert/PpVert) for text extraction.
rem   %1 = EdSharp program directory   %2 = source document   %3 = target text file
rem 2htm writes <outputdir>\<sourcebasename>.txt, so we point it at the target's
rem directory and then rename its output to the exact target EdSharp expects.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
set "outdir=%~dp3"
"%prog%\Convert\2htm\2htm.exe" "%src%" -p -f -o "%outdir%"
set "made=%outdir%%~n2.txt"
if exist "%made%" if /i not "%made%"=="%dst%" move /y "%made%" "%dst%" >nul
endlocal
