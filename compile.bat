@echo off
cls
if exist EdSharp.exe del EdSharp.exe
if exist EdSharp64.exe del EdSharp64.exe
rem csc.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll
rem Make EdSharp.exe Win32 and EdSharp64.exe Win64
rem c:\windows\Microsoft.NET\Framework\v2.0.50727\csc.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs 
rem try Roslyn
rem c:\windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /platform:x86 /out:EdSharp.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs >temp.txt
rem C:\Roslyn\csc.exe /platform:x86 /out:EdSharp.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs >temp.txt
rem try latest compiler from build tools as recommended by AI
rem c:\windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe /platform:x86 /out:EdSharp.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs >temp.txt
"c:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\MSBuild\Current\Bin\Roslyn\csc.exe" /nologo /platform:x86 /out:EdSharp.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs >temp.txt
if errorlevel 1 goto end
rem 32-bit is most reliable
rem c:\windows\Microsoft.NET\Framework\v4.0.30319\csc.exe /platform:x64 /out:EdSharp64.exe /nologo /t:winexe /r:Tektosyne.dll /r:IronCOM.dll /r:Microsoft.JScript.dll /r:eval.dll /r:VB.dll /r:Microsoft.VisualBasic.dll /r:Microsoft.VisualBasic.Compatibility.dll EdSharp.cs >nul  
rem if exist EdSharp.Exe EdSharp.exe
if exist EdSharp.Exe echo Done
:end
