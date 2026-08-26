@echo off
rem installPdfTools.cmd -- install the free document tools EdSharp can use in
rem place of Microsoft Office: PyMuPDF4LLM, which turns a PDF's own structure
rem into Markdown with headings, lists and tables, and WordNet, the lexical
rem database behind the thesaurus. Free, no Microsoft Word, no account.
rem Roughly 55 megabytes together.
rem Probe first, install or upgrade, log milestones, pause on failure.
rem NOTHING PAUSES: a console waiting for a keypress interrupts the
rem installation. Failures are logged, and the summary shown at the very
rem end reports the outcome of every checkbox.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
echo [installPdfTools] started %date% %time% >> "%logFile%"
echo If Windows asks permission, a User Account Control prompt appears on a
echo separate screen; press Alt+Y to allow it.
echo.

call :findPython
if not defined pythonExe goto no_python
echo Using %pythonExe%
echo [installPdfTools] python: %pythonExe% >> "%logFile%"
"%pythonExe%" -c "import pymupdf4llm" >nul 2>&1
if errorlevel 1 goto install_reader
echo The PDF reader is already installed; checking for an update.
echo [installPdfTools] pip install --upgrade pymupdf4llm >> "%logFile%"
"%pythonExe%" -m pip install --upgrade pymupdf4llm
echo [installPdfTools] upgrade exit %errorlevel% >> "%logFile%"
goto done

:install_reader
echo Installing the PDF reader; this takes a minute and about 25 megabytes.
echo [installPdfTools] pip install pymupdf4llm >> "%logFile%"
"%pythonExe%" -m pip install pymupdf4llm
echo [installPdfTools] install exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto failed

:done
"%pythonExe%" -c "import pymupdf4llm" >nul 2>&1
if errorlevel 1 goto failed

rem The free thesaurus: WordNet, Princeton's lexical database, read through
rem the nltk package. Roughly 30 megabytes with its data, and it gives
rem synonyms grouped by meaning with a definition for each sense.
echo.
echo Installing the free thesaurus database.
echo [installPdfTools] pip install nltk >> "%logFile%"
"%pythonExe%" -m pip install --upgrade nltk
echo [installPdfTools] nltk exit %errorlevel% >> "%logFile%"
"%pythonExe%" -c "import nltk; nltk.download('wordnet'); nltk.download('omw-1.4')"
echo [installPdfTools] wordnet data exit %errorlevel% >> "%logFile%"
"%pythonExe%" -c "from nltk.corpus import wordnet; wordnet.synsets('test')" >nul 2>&1
if errorlevel 1 (echo The thesaurus did not install; the PDF reader is still ready.) else (echo The thesaurus is ready: press Shift+F7 on a word in EdSharp.)

echo.
echo Done. EdSharp can now convert PDF files to Markdown, HTML and Word
echo documents with their headings, lists and tables intact.
echo [installPdfTools] done >> "%logFile%"
exit /b 0

:no_python
echo The official Python from python.org was not found.
echo.
echo Windows may have a stub named python that only advertises the Microsoft
echo Store; that is not Python and cannot install anything. Run
echo installPython.cmd in this folder, or rerun the EdSharp installer and
echo tick the Python box on its last page, then run this script again.
echo [installPdfTools] FAILED: no python >> "%logFile%"
exit /b 7

:failed
echo The PDF reader did not install. The log is:
echo %logFile%
echo [installPdfTools] FAILED >> "%logFile%"
exit /b 3

:findPython
rem Sets pythonExe to a REAL python.org Python, or leaves it empty. The
rem Microsoft "app execution alias" at %LOCALAPPDATA%\Microsoft\WindowsApps
rem answers `where python` and then advertises the Store rather than running
rem anything, so that path is rejected; what remains must actually answer
rem --version. Freshly installed Python is also looked for by location,
rem because this console's PATH was inherited before the install happened.
set "pythonExe="
for /f "delims=" %%p in ('where python 2^>nul') do (
  echo %%p | find /i "\WindowsApps\" >nul
  if errorlevel 1 (
    if not defined pythonExe set "pythonExe=%%p"
  )
)
if not defined pythonExe (
  for %%d in ("%ProgramFiles%" "%LOCALAPPDATA%\Programs\Python") do (
    for /d %%s in ("%%~d\Python3*") do (
      if exist "%%~s\python.exe" if not defined pythonExe set "pythonExe=%%~s\python.exe"
    )
  )
)
if defined pythonExe (
  "%pythonExe%" --version >nul 2>&1
  if errorlevel 1 set "pythonExe="
)
exit /b 0
