@echo off
rem any2md.cmd -- convert a supported document to Markdown for EdSharp's Import.
rem Two steps, per the conversion policy: 2htm makes HTML from formats Pandoc
rem cannot read (.doc, PDF, PowerPoint), then Pandoc turns that HTML into
rem GitHub-flavored Markdown.
rem   %1 = EdSharp program directory   %2 = source document   %3 = target file
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
set "pandocExe=%prog%\Convert\Pandoc\pandoc.exe"
if not exist "%pandocExe%" set "pandocExe=pandoc"
set "mid=%dst%.htm"
call "%prog%\Convert\any2htm.cmd" "%prog%" "%src%" "%mid%"
if exist "%mid%" "%pandocExe%" "%mid%" -f html -t gfm -o "%dst%"
if exist "%mid%" del /f /q "%mid%"
endlocal
