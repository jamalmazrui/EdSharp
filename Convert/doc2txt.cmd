@echo off
rem doc2txt.cmd -- convert a Word document to plain text for EdSharp's Import.
rem   %1 = EdSharp program directory   %2 = source document   %3 = target text file
rem EdSharp reads the target file after this runs, so the job is to WRITE %3.
rem Tries the Word COM converter (WdVert) first; if it produced nothing, falls
rem back to GetText. All paths are quoted so spaces and 8.3-disabled volumes are
rem handled.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
"%prog%\Convert\OfficeConvert\WdVert.exe" "%src%" "%dst%"
if exist "%dst%" goto :done
"%prog%\Convert\GetText\GetText.exe" "%src%" "%dst%"
:done
endlocal
