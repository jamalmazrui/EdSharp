Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = """[^""]*"""
re.Global = true
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	newstr = re.Replace(line, """****""")
	WScript.Echo "New string '" & newstr & "', original '" & line & "'"
Loop
s.Close
