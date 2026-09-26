@echo off
rem any2txt.cmd -- convert a supported document to plain text for EdSharp's
rem Import, using 2htm.exe in plain-text mode (-p): one step, no intermediary.
rem 2htm reads Word documents, PDF (through Word's PDF Reflow), slides,
rem spreadsheets, and CSV, and its -p mode writes .txt directly -- which is
rem why it replaced the older Office COM converters (WdVert, XlVert, PpVert)
rem for text extraction. Those OfficeConvert utilities still ship in the
rem Convert folder for command-line use.
rem   %1 = EdSharp program directory   %2 = source document   %3 = target file
rem 2htm writes <outputdir>\<sourcebasename>.txt, so we point it at the
rem target's directory and rename its output to the exact target. The tool's
rem console output is captured beside the target as <target>.log, and a
rem failure exits nonzero, so EdSharp's error dialog can quote the reason.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
if exist "%dst%.log" del /f /q "%dst%.log"
set "outdir=%~dp3"
rem %~dp3 ends with a backslash; backslash-quote hands the tool a literal
rem quote in the path. Trim it, as in any2htm.cmd.
if "%outdir:~-1%"=="\" set "outdir=%outdir:~0,-1%"
if not exist "%prog%\Convert\2htm\2htm.exe" (
  echo 2htm.exe was not found at %prog%\Convert\2htm\2htm.exe > "%dst%.log"
  echo Run restoreFetchConvertTools.cmd in the EdSharp folder to fetch the conversion tools. >> "%dst%.log"
  exit /b 8
)
"%prog%\Convert\2htm\2htm.exe" "%src%" -p -f -o "%outdir%" > "%dst%.log" 2>&1
set "toolExit=%errorlevel%"
set "made=%outdir%\%~n2.txt"
if exist "%made%" if /i not "%made%"=="%dst%" move /y "%made%" "%dst%" >nul
if not exist "%dst%" (
  if "%toolExit%"=="0" set "toolExit=9"
  exit /b %toolExit%
)
del /f /q "%dst%.log" >nul 2>&1
exit /b 0
