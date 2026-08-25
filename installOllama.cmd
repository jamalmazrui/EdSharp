@echo off
rem installOllama.cmd -- smart install of Ollama plus a chat model for
rem EdSharp's Chat with AI command. Probe first: an existing Ollama is
rem updated in place, never duplicated. Tool output stays in this
rem window so progress is readable; the log records milestones.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
set "modelName=llama3.2"
echo [installOllama] started %date% %time% >> "%logFile%"
echo If Windows asks permission, a User Account Control prompt appears on a
echo separate screen; press Alt+Y to allow it. The model download is about
echo 2 gigabytes and shows its progress here.
echo.

where ollama >nul 2>&1
if errorlevel 1 goto install_ollama
echo Ollama is already installed; checking for an update.
echo [installOllama] winget upgrade Ollama.Ollama >> "%logFile%"
winget upgrade --id Ollama.Ollama -e --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installOllama] winget upgrade exit %errorlevel% >> "%logFile%"
if errorlevel 1 (echo Ollama is already current.) else (echo Ollama updated.)
goto pull_model
:install_ollama
echo Installing Ollama with winget; this can take a few minutes.
echo [installOllama] winget install Ollama.Ollama >> "%logFile%"
winget install --id Ollama.Ollama -e --silent --disable-interactivity --accept-package-agreements --accept-source-agreements
echo [installOllama] winget install exit %errorlevel% >> "%logFile%"
set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto fail_ollama

:pull_model
echo Fetching the %modelName% chat model.
echo [installOllama] ollama pull %modelName% >> "%logFile%"
ollama pull %modelName%
echo [installOllama] ollama pull exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto fail_model
echo Done. EdSharp's Chat with AI command is ready.
echo [installOllama] done >> "%logFile%"
exit /b 0

:fail_ollama
echo Ollama was not found after the install step. The log is:
echo %logFile%
echo [installOllama] FAILED: ollama not found after install >> "%logFile%"
pause
exit /b 3

:fail_model
echo The model download did not finish. Run this script again later, or
echo run: ollama pull %modelName%
echo [installOllama] model pull failed >> "%logFile%"
pause
exit /b 4
