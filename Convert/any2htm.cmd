@echo off
rem any2htm.cmd -- convert a supported document to HTML for EdSharp's Import,
rem using 2htm.exe in its default HTML mode.  2htm reads Word (.doc .docx .rtf
rem .odt), PDF (through Word's PDF Reflow), Excel, PowerPoint, and CSV.
rem   %1 = EdSharp program directory   %2 = source document   %3 = target file
rem 2htm writes <outputdir>\<sourcebasename>.htm, so we point it at the target's
rem directory and then rename its output to the exact target EdSharp expects.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
set "outdir=%~dp3"
"%prog%\Convert\2htm\2htm.exe" "%src%" -f -o "%outdir%"
set "made=%outdir%%~n2.htm"
if exist "%made%" if /i not "%made%"=="%dst%" move /y "%made%" "%dst%" >nul
endlocal
