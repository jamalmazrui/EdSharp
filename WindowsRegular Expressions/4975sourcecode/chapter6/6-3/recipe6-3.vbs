Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "\bMyMethod\s*\(\s*([^,]+|""[^""]*"")\s*,\s*([^,]+|""[^""]*"")(?=\s*[,)])"
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	newstr = re.Replace(line, "MyMethod($2, $1)")
	WScript.Echo "New string '" & newstr & "', original '" & line & "'"
Loop
s.Close
