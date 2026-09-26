@echo off
rem pdf2md.cmd -- convert a PDF to RICH Markdown for EdSharp's Import, with
rem headings, lists, tables and emphasis preserved. No Microsoft Word.
rem   %1 = EdSharp program directory   %2 = source PDF   %3 = target .md
rem The work is done by pdfRich.py (PyMuPDF4LLM). Its console output is
rem captured beside the target as <target>.log so EdSharp can quote the
rem reason for any failure, and the real exit code is propagated.
setlocal
set "prog=%~1"
set "src=%~2"
set "dst=%~3"
if exist "%dst%" del /f /q "%dst%"
if not exist "%prog%\Convert\pdfRich.py" (
  echo pdfRich.py was not found at %prog%\Convert\pdfRich.py > "%dst%.log"
  exit /b 8
)
where python >nul 2>&1
if errorlevel 1 (
  echo Python was not found on this computer. > "%dst%.log"
  echo Rerun the EdSharp installer and tick the Python box on its last page, >> "%dst%.log"
  echo then run installPdfTools.cmd in the EdSharp program folder. >> "%dst%.log"
  exit /b 7
)
python "%prog%\Convert\pdfRich.py" "%src%" "%dst%"
exit /b %errorlevel%
