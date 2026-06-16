Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^\s*Dim\s+(\w+)\s+As\s+([\w.]+)"
re.IgnoreCase = True
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	If re.Test(line) = True Then
		newstr = re.Replace(line, "Found variable: '$1'; type: '$2'")
		WScript.Echo newstr
	End If
Loop
s.Close
