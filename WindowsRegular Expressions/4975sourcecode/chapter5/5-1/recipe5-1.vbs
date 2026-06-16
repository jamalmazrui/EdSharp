Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "<[a-z:_][-a-z0-9._:]+((\s+[^>]+)*|/s*)/?>"
re.IgnoreCase = True
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(line) Then
		WScript.Echo "Found match: '" & line & "' at line " & lineNbr
	End If
Loop
s.Close
