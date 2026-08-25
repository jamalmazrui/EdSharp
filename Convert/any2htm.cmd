@echo off
rem any2htm.cmd -- convert a supported document to HTML for EdSharp's Import,
rem using 2htm.exe in its default HTML mode.  2htm reads Word (.doc .docx .rtf
rem .odt), PDF (through Word's PDF Reflow), Excel, PowerPoint, and CSV.
rem   %1 = EdSharp program directory   %2 = source document   %3 = target file
rem 2htm writes <outputdir>\<sourcebasename>.htm, so we point it at the target's
rem directory and then rename its output to the exact target EdSharp expects.
rem 2htm's console output is captured beside the target as <target>.log, and a
rem failure exits nonzero, so EdSharp's run log and error dialog can say WHY a
rem conversion produced nothing instead of showing a bare command line.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
if exist "%dst%.log" del /f /q "%dst%.log"
set "outdir=%~dp3"
if not exist "%prog%\Convert\2htm\2htm.exe" (
  echo 2htm.exe was not found at %prog%\Convert\2htm\2htm.exe > "%dst%.log"
  echo Run restoreFetchConvertTools.cmd in the EdSharp folder to fetch the conversion tools. >> "%dst%.log"
  exit /b 8
)
"%prog%\Convert\2htm\2htm.exe" "%src%" -f -o "%outdir%" > "%dst%.log" 2>&1
set "toolExit=%errorlevel%"
set "made=%outdir%%~n2.htm"
if exist "%made%" if /i not "%made%"=="%dst%" move /y "%made%" "%dst%" >nul
if not exist "%dst%" (
  if "%toolExit%"=="0" set "toolExit=9"
  exit /b %toolExit%
)
del /f /q "%dst%.log" >nul 2>&1
exit /b 0
