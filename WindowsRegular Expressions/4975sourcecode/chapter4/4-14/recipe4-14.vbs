Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^(\d{4}-){3}|(\d{4} ){3}\d{4}|\d{15,16}|\d{4} \d{2} \d{4} \d{5}|\d{4}-\d{2}-\d{4}-\d{5}$"
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(line) Then
		WScript.Echo "Found match: '" & line & "' at line " & lineNbr
	End If
Loop
s.Close
