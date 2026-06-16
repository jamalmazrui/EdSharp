@echo off
cls
if exist EdSharp.exe del EdSharp.exe
rem csc.exe /nologo /t:exe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll
c:\windows\Microsoft.NET\Framework\v2.0.50727\csc.exe /nologo /t:exe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs 
if exist EdSharp.Exe EdSharp.exe
