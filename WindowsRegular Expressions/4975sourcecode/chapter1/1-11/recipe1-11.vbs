Dim fso,s,re,line,lastLine,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "\b(\w+)\s*(\r\n)*\s*\1\b"
lineNbr = 0
lastLine = ""
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(lastLine & line) Then
		WScript.Echo "Found matches: '" & lastLine & "' and '" & line & _
			"' at lines " & (lineNbr - 1) & " and " & lineNbr
	End If
	lastLine = line
Loop
s.Close
