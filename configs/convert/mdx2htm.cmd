@echo off
rem mdx2htm.cmd -- expand embedded inix tables, then let pandoc make the HTML.
rem   %1 = EdSharp program directory   %2 = source .mdx or .md   %3 = target file
rem Step 1: inixVert turns every fenced "inix" block into a real Markdown
rem         table (a grid table when cells are multi-line).
rem Step 2: pandoc converts the expanded Markdown to the target format.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
set "pandocExe=%prog%\Convert\Pandoc\pandoc.exe"
if not exist "%pandocExe%" set "pandocExe=pandoc"
set "mid=%dst%.md"
"%prog%\Convert\inixVert.exe" "%src%" "%mid%" /quiet
if exist "%mid%" "%pandocExe%" "%mid%" -f markdown -t html -s -o "%dst%"
if exist "%mid%" del /f /q "%mid%"
endlocal
