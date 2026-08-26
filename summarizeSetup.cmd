@echo off
rem summarizeSetup.cmd -- report the disposition of every optional install.
rem
rem The Results box the installer shows appears BEFORE the finish-page
rem checkboxes run, so it cannot know how they fared. This runs last,
rem after them, and reports what is actually on the computer now: for
rem each optional tool, whether it is present and which version, or that
rem it is absent and how to add it later. The report is printed here,
rem written to the logs folder, appended to the consolidated setup log,
rem and shown in a message box so it is read whether or not this window
rem has focus.
setlocal enabledelayedexpansion
set "logDir=%LOCALAPPDATA%\EdSharp\logs"
if not exist "%logDir%" mkdir "%logDir%" >nul 2>&1
set "logFile=%logDir%\EdSharp_setup.log"
set "sumFile=%logDir%\EdSharp_setup_summary.txt"
if exist "%sumFile%" del /f /q "%sumFile%"
if exist "%LOCALAPPDATA%\Programs\Ollama" set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"

set "resultsFile=%logDir%\EdSharp_setup_results.txt"
call :head "EdSharp setup results  %date% %time%"
call :head ""
rem What the installer knew before the checkboxes ran -- the JAWS scripts,
rem the NVDA add-on, pandoc -- was handed over in a file so that ONE box
rem tells the whole story instead of two telling halves.
if exist "%resultsFile%" (
  for /f "usebackq delims=" %%l in ("%resultsFile%") do call :head "%%l"
  del /f /q "%resultsFile%" >nul 2>&1
  call :head ""
  call :head "Optional installs"
)
call :tool "Git" git "winget install Git.Git"
call :tool "GitHub command line" gh "winget install GitHub.cli"
call :tool "Node.js" node "run installNode.cmd in the EdSharp folder"
call :python
call :module "PDF reader (rich PDF conversion)" pymupdf4llm "run installPdfTools.cmd in the EdSharp folder"
call :wordnet
call :tool "Ollama (Chat with AI)" ollama "run installOllama.cmd in the EdSharp folder"
call :model
call :pandoc
call :head ""
call :head "This summary is saved as: %sumFile%"
call :head "Full detail is in: %logFile%"
call :head ""
call :head "To start EdSharp, press Alt+Control+E."

rem The only thing the person sees: one Results box, after everything.
rem This script runs hidden, so nothing flashes and nothing waits for a
rem keypress; the text also stays in the summary file and the log.
powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show((Get-Content -Raw '%sumFile%'), 'EdSharp Setup Results')" >nul 2>&1
exit /b 0

:head
rem An empty line must be echoed as a blank line; "echo" with nothing after
rem it prints "ECHO is off." instead, which appeared in the first summary.
if "%~1"=="" (echo.) else (echo %~1)
if "%~1"=="" (echo.>> "%sumFile%") else (echo %~1>> "%sumFile%")
if not "%~1"=="" echo [summary] %~1 >> "%logFile%"
exit /b 0

:tool
rem %1 = friendly name, %2 = executable, %3 = how to add it later
set "toolVersion="
where %2 >nul 2>&1
if errorlevel 1 (
  call :head "%~1: not installed. To add it later, %~3."
  exit /b 0
)
for /f "delims=" %%v in ('%2 --version 2^>^&1') do if not defined toolVersion set "toolVersion=%%v"
call :head "%~1: installed, !toolVersion!"
exit /b 0

:python
call :findPython
if not defined pythonExe (
  where python >nul 2>&1
  if errorlevel 1 (
    call :head "Python: not installed. To add it later, run installPython.cmd in the EdSharp folder."
  ) else (
    call :head "Python: not installed. Windows has only the Microsoft Store stub, which is not Python; run installPython.cmd in the EdSharp folder for the official python.org build."
  )
  exit /b 0
)
set "pyVersion="
for /f "delims=" %%v in ('"%pythonExe%" --version 2^>^&1') do if not defined pyVersion set "pyVersion=%%v"
call :head "Python: installed, !pyVersion! at %pythonExe%"
exit /b 0

:module
rem %1 = friendly name, %2 = python module, %3 = how to add it later
call :findPython
if not defined pythonExe (
  call :head "%~1: not installed, because Python is missing."
  exit /b 0
)
"%pythonExe%" -c "import %2" >nul 2>&1
if errorlevel 1 (
  call :head "%~1: not installed. To add it later, %~3."
) else (
  call :head "%~1: installed."
)
exit /b 0

:wordnet
call :findPython
if not defined pythonExe (
  call :head "Thesaurus database: not installed, because Python is missing."
  exit /b 0
)
"%pythonExe%" -c "from nltk.corpus import wordnet; wordnet.synsets('test')" >nul 2>&1
if errorlevel 1 (
  call :head "Thesaurus database: not installed. To add it later, run installPdfTools.cmd in the EdSharp folder."
) else (
  call :head "Thesaurus database: installed. Press Shift+F7 on a word."
)
exit /b 0

:model
where ollama >nul 2>&1
if errorlevel 1 exit /b 0
ollama list 2>nul | find /i "llama3.2" >nul 2>&1
if errorlevel 1 (
  call :head "Chat model llama3.2: not downloaded. Press F12 in EdSharp and answer Yes when it offers to fetch it."
) else (
  call :head "Chat model llama3.2: ready. Press F12 in EdSharp to chat."
)
exit /b 0

:pandoc
if exist "%~dp0Convert\Pandoc\pandoc.exe" (
  call :head "Pandoc: present. Document conversion will work."
) else (
  call :head "Pandoc: not present. To add it later, run installPandoc.cmd as an administrator."
)
exit /b 0

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
