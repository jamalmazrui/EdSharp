@echo off
rem sheet2md.cmd -- convert an Excel workbook to a Markdown table for EdSharp's
rem Import.  Pandoc has no Excel reader, so OfficeConvert's XlVert makes a CSV
rem first, and Pandoc turns the CSV into a Markdown table.
rem   %1 = EdSharp program directory   %2 = source workbook   %3 = target file
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
set "pandocExe=%prog%\Convert\Pandoc\pandoc.exe"
if not exist "%pandocExe%" set "pandocExe=pandoc"
set "mid=%dst%.csv"
if exist "%mid%" del /f /q "%mid%"
"%prog%\Convert\OfficeConvert\XlVert.exe" "%src%" "%mid%"
if exist "%mid%" "%pandocExe%" "%mid%" -f csv -t gfm -o "%dst%"
if exist "%mid%" del /f /q "%mid%"
endlocal
