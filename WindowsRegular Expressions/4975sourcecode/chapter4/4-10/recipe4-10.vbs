Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
Const octetRegex = "([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])"
re.Pattern = "^(" + octetRegex + "\.){3}" + octetRegex + "$" 
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(line) Then
		WScript.Echo "Found match: '" & line & "' at line " & lineNbr
	End If
Loop
s.Close
