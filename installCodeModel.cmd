@echo off
rem installCodeModel.cmd -- install the AI model EdSharp uses for coding
rem questions about the file in the window. About 5 gigabytes.
rem Probe first, log milestones, never pause; the Results box reports the
rem outcome. The console says only what is happening.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
set "modelName=qwen2.5-coder:7b"
echo [installCodeModel] started %date% %time% >> "%logFile%"

if exist "%LOCALAPPDATA%\Programs\Ollama" set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto no_ollama

call :ollamaModels
echo %modelList% | find /i "%modelName%" >nul 2>&1
if not errorlevel 1 (
  echo The %modelName% model is already installed.
  echo [installCodeModel] already present >> "%logFile%"
  exit /b 0
)

echo Fetching the %modelName% model, about 5 GB
echo [installCodeModel] ollama pull %modelName% >> "%logFile%"
call :ollamaPullHidden %modelName%
echo [installCodeModel] pull exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto failed
echo Done.
echo [installCodeModel] done >> "%logFile%"
exit /b 0

:no_ollama
echo Ollama is not installed, so the coding model cannot be fetched.
echo Tick the Ollama box as well, or run installOllama.cmd first.
echo [installCodeModel] FAILED: no ollama >> "%logFile%"
exit /b 7

:failed
echo The model did not download. The log is:
echo %logFile%
echo [installCodeModel] FAILED >> "%logFile%"
exit /b 3
