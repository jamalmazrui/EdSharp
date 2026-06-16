Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^(?<key>[^#=]+)\s*=\s*(?<value>[^#]*)\s*"
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	If re.Test(line) = True Then
		newstr = re.Replace(line, "Found key: '$1'; value: '$2'")
		WScript.Echo newstr
	End If
Loop
s.Close
