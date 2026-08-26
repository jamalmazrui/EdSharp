@echo off
rem installTranslateModel.cmd -- install the larger AI model EdSharp uses for
rem translation when it is present. The small chat model translates
rem passably; this one translates well, at about 5 gigabytes.
rem Probe first, log milestones, never pause; the Results box reports the
rem outcome. The console says only what is happening.
setlocal
set "logFile=%LOCALAPPDATA%\EdSharp\logs\EdSharp_setup.log"
if not exist "%LOCALAPPDATA%\EdSharp\logs" mkdir "%LOCALAPPDATA%\EdSharp\logs" >nul 2>&1
set "modelName=qwen2.5:7b"
echo [installTranslateModel] started %date% %time% >> "%logFile%"

if exist "%LOCALAPPDATA%\Programs\Ollama" set "PATH=%LOCALAPPDATA%\Programs\Ollama;%PATH%"
where ollama >nul 2>&1
if errorlevel 1 goto no_ollama

ollama list 2>nul | find /i "%modelName%" >nul 2>&1
if not errorlevel 1 (
  echo The %modelName% model is already installed.
  echo [installTranslateModel] already present >> "%logFile%"
  exit /b 0
)

echo Fetching the %modelName% model, about 5 GB ...
echo [installTranslateModel] ollama pull %modelName% >> "%logFile%"
ollama pull %modelName% >> "%logFile%" 2>&1
echo [installTranslateModel] pull exit %errorlevel% >> "%logFile%"
if errorlevel 1 goto failed
echo Done.
echo [installTranslateModel] done >> "%logFile%"
exit /b 0

:no_ollama
echo Ollama is not installed, so the translation model cannot be fetched.
echo Tick the Ollama box as well, or run installOllama.cmd first.
echo [installTranslateModel] FAILED: no ollama >> "%logFile%"
exit /b 7

:failed
echo The model did not download. The log is:
echo %logFile%
echo [installTranslateModel] FAILED >> "%logFile%"
exit /b 3
