Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^(00[1-9]|0[1-9]\d|[1-5]\d{2}|6[0-5]\d|66[0-5]|66[7-9]|6[7-8]\d|690|7[0-2]\d|73[0-3]|750|76[4-9]|77[0-2])-(?!00)\d{2}-(?!0000)\d{4}$"
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(line) Then
		WScript.Echo "Found match: '" & line & "' at line " & lineNbr
	End If
Loop
s.Close
