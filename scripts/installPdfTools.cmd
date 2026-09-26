@echo off
setlocal enabledelayedexpansion
set "sScript=%~n0"
set "sCallerDir=%~dp0"
if not exist "%~dp0installCommon.cmd" (
  echo(
  echo installCommon.cmd is missing from %~dp0
  echo That file is part of this program. Reinstall, or copy it from the
  echo program's zip into this folder, and run this again.
  echo(
  if not defined noPause pause
  exit /b 1
)
call "%~dp0installCommon.cmd" setup "%~f0" %*
rem installPdfTools.cmd -- install the free document tools EdSharp can use in
rem place of Microsoft Office: PyMuPDF4LLM, which turns a PDF's own structure
rem into Markdown with headings, lists and tables, and WordNet, the lexical
rem database behind the thesaurus. Free, no Microsoft Word, no account.
rem Roughly 55 megabytes together.
rem Probe first, install or upgrade, log milestones, pause on failure.
rem NOTHING PAUSES: a console waiting for a keypress interrupts the
rem installation. Failures are logged, and the summary shown at the very
rem end reports the outcome of every checkbox.
echo.

call :findPython
if not defined pythonExe goto no_python
echo [installPdfTools] python: %pythonExe% >> "%log%"
"%pythonExe%" -c "import pymupdf4llm" >nul 2>&1
if errorlevel 1 goto install_reader
echo Updating the PDF reader
echo [installPdfTools] pip install --upgrade pymupdf4llm >> "%log%"
"%pythonExe%" -m pip install --upgrade pymupdf4llm >> "%log%" 2>&1
echo [installPdfTools] upgrade exit %errorlevel% >> "%log%"
goto done

:install_reader
echo Installing the PDF reader
echo [installPdfTools] pip install pymupdf4llm >> "%log%"
"%pythonExe%" -m pip install pymupdf4llm >> "%log%" 2>&1
echo [installPdfTools] install exit %errorlevel% >> "%log%"
if errorlevel 1 goto failed

:done
rem Prove it rather than assume it. pip returning 0 means pip ran, not
rem that the package can be imported: a wheel can install and still fail
rem to load. The import is attempted with the SAME interpreter that did
rem the installing, and whatever it says goes into the log, so a
rem disagreement with the summary can never again be a mystery.
"%pythonExe%" -c "import pymupdf4llm; print('pymupdf4llm ready')" >> "%log%" 2>&1
if errorlevel 1 goto failed
echo [installPdfTools] verified pymupdf4llm with %pythonExe% >> "%log%"
echo %pythonExe%> "%LOCALAPPDATA%\EdSharp\logs\EdSharp_python.txt"
echo PDF reader ready.

rem The free thesaurus: WordNet, Princeton's lexical database, read through
rem the nltk package. Roughly 30 megabytes with its data, and it gives
rem synonyms grouped by meaning with a definition for each sense.
echo.
echo Installing the thesaurus
echo [installPdfTools] pip install nltk >> "%log%"
"%pythonExe%" -m pip install --upgrade nltk >> "%log%" 2>&1
echo [installPdfTools] nltk exit %errorlevel% >> "%log%"
"%pythonExe%" -c "import nltk; nltk.download('wordnet'); nltk.download('omw-1.4')"
echo [installPdfTools] wordnet data exit %errorlevel% >> "%log%"
"%pythonExe%" -c "from nltk.corpus import wordnet; wordnet.synsets('test'); print('wordnet ready')" >> "%log%" 2>&1
if errorlevel 1 (
  echo Thesaurus not installed; the PDF reader is ready.
  echo [installPdfTools] wordnet verify FAILED >> "%log%"
) else (
  echo Thesaurus ready.
  echo [installPdfTools] verified wordnet with %pythonExe% >> "%log%"
)

echo.
echo Done.
echo [installPdfTools] done >> "%log%"
exit /b 0

:no_python
echo The official Python from python.org was not found.
echo.
echo Windows may have a stub named python that only advertises the Microsoft
echo Store; that is not Python and cannot install anything. Run
echo installPython.cmd in this folder, or rerun the EdSharp installer and
echo tick the Python box on its last page, then run this script again.
echo [installPdfTools] FAILED: no python >> "%log%"
exit /b 7

:failed
echo The PDF reader did not install. The log is:
echo %log%
echo [installPdfTools] FAILED >> "%log%"
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
  for %%d in ("%ProgramFiles%" "%LOCALAPPDATA%\Programs\Python" "%SystemDrive%\") do (
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
