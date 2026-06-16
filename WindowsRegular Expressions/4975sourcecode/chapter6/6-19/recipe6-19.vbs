Dim fso,s,re,line,lineNbr,m,matches
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "(({{)*{)(?!{)([^}]+)"
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	Set matches = re.Execute(line)
	For Each m In matches
		WScript.Echo "Found match: '" & m.Value & "' at line " & lineNbr
	Next
Loop
s.Close
