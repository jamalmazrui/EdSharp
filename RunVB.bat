@echo off
cls
if exist VB.dll del VB.dll
c:\windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe /nologo /t:library /r:Microsoft.VisualBasic.dll VB.vb
