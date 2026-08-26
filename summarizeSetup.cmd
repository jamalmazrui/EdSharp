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

call :head "EdSharp setup summary  %date% %time%"
call :head ""
call :tool "Git" git "winget install Git.Git"
call :tool "GitHub command line" gh "winget install GitHub.cli"
call :tool "Node.js" node "run installNode.cmd in the EdSharp folder"
call :tool "Python" python "run installPython.cmd in the EdSharp folder"
call :module "PDF reader (rich PDF conversion)" pymupdf4llm "run installPdfTools.cmd in the EdSharp folder"
call :wordnet
call :tool "Ollama (Chat with AI)" ollama "run installOllama.cmd in the EdSharp folder"
call :model
call :pandoc
call :head ""
call :head "Full detail is in: %logFile%"

echo.
echo This summary is also saved as:
echo %sumFile%
powershell -NoProfile -Command "Add-Type -AssemblyName System.Windows.Forms; [void][System.Windows.Forms.MessageBox]::Show((Get-Content -Raw '%sumFile%'), 'EdSharp Setup Summary')" >nul 2>&1
exit /b 0

:head
echo %~1
echo %~1>> "%sumFile%"
echo [summary] %~1 >> "%logFile%"
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

:module
rem %1 = friendly name, %2 = python module, %3 = how to add it later
where python >nul 2>&1
if errorlevel 1 (
  call :head "%~1: not installed, because Python is missing."
  exit /b 0
)
python -c "import %2" >nul 2>&1
if errorlevel 1 (
  call :head "%~1: not installed. To add it later, %~3."
) else (
  call :head "%~1: installed."
)
exit /b 0

:wordnet
where python >nul 2>&1
if errorlevel 1 (
  call :head "Thesaurus database: not installed, because Python is missing."
  exit /b 0
)
python -c "from nltk.corpus import wordnet; wordnet.synsets('test')" >nul 2>&1
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
