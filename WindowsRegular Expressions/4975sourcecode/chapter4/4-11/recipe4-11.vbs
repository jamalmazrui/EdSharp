Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
Private Const localRegex = "[\w\d!#$%&'*+-/=?^`{|}~]+(\.[\w\d!#$%&'*+-/=?^`{|}~]+)*"
Private Const hostNameRegex = "([a-z\d][-a-z\d]*[a-z\d]\.)*[a-z][-a-z\d]*[a-z]"
re.Pattern = "^" & localRegex & "@" & hostNameRegex & "$"
re.IgnoreCase = True
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(line) Then
		WScript.Echo "Found match: '" & line & "' at line " & lineNbr
	End If
Loop
s.Close
