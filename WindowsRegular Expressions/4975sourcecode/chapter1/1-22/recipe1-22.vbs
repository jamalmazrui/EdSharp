Dim fso,s,re,line,newstr,f
Set fso = CreateObject("Scripting.FileSystemObject")
Set f = fso.GetFile(WScript.Arguments.Item(0))
Set s = f.OpenAsTextStream(1, -1)
Set re = New RegExp
re.Pattern = "\u201c|\u201d"
re.Global = True
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	newstr = re.Replace(line, """")
	WScript.Echo "New string '" & newstr & "', original '" & line & "'"
Loop
s.Close
