@echo off
rem installPdfTools.cmd -- install the free document tools EdSharp can use in
rem place of Microsoft Office: PyMuPDF4LLM, which turns a PDF's own structure
rem into Markdown with headings, lists and tables, and WordNet, the lexical
rem database behind the thesaurus. Free, no Microsoft Word, no account.
rem Roughly 55 megabytes together.
rem Probe first, install or upgrade, log milestones, pause on failure.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
echo [installPdfTools] started %date% %time% >> "%logFile%"
echo If Windows asks permission, a User Account Control prompt appears on a
echo separate screen; press Alt+Y to allow it.
echo.

where python >nul 2>&1
if errorlevel 1 goto no_python
python -c "import pymupdf4llm" >nul 2>&1
if errorlevel 1 goto install_reader
echo The PDF reader is already installed; checking for an update.
echo [installPdfTools] pip install --upgrade pymupdf4llm >> "%logFile%"
python -m pip install --upgrade pymupdf4llm
echo [installPdfTools] upgrade exit %errorlevel% >> "%logFile%"
goto done

:install_reader
echo Installing the PDF reader; this takes a minute and about 25 megabytes.
echo [installPdfTools] pip install pymupdf4llm >> "%logFile%"
python -m pip install pymupdf4llm
echo [installPdfTools] install exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto failed

:done
python -c "import pymupdf4llm" >nul 2>&1
if errorlevel 1 goto failed

rem The free thesaurus: WordNet, Princeton's lexical database, read through
rem the nltk package. Roughly 30 megabytes with its data, and it gives
rem synonyms grouped by meaning with a definition for each sense.
echo.
echo Installing the free thesaurus database.
echo [installPdfTools] pip install nltk >> "%logFile%"
python -m pip install --upgrade nltk
echo [installPdfTools] nltk exit %errorlevel% >> "%logFile%"
python -c "import nltk; nltk.download('wordnet'); nltk.download('omw-1.4')"
echo [installPdfTools] wordnet data exit %errorlevel% >> "%logFile%"
python -c "from nltk.corpus import wordnet; wordnet.synsets('test')" >nul 2>&1
if errorlevel 1 (echo The thesaurus did not install; the PDF reader is still ready.) else (echo The thesaurus is ready: press Shift+F7 on a word in EdSharp.)

echo.
echo Done. EdSharp can now convert PDF files to Markdown, HTML and Word
echo documents with their headings, lists and tables intact.
echo [installPdfTools] done >> "%logFile%"
exit /b 0

:no_python
echo Python was not found. Rerun the EdSharp installer and tick the Python
echo box on its last page, then run this script again.
echo [installPdfTools] FAILED: no python >> "%logFile%"
pause
exit /b 7

:failed
echo The PDF reader did not install. The log is:
echo %logFile%
echo [installPdfTools] FAILED >> "%logFile%"
pause
exit /b 3
