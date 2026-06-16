Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "\r\n"
re.Global = True
line = s.ReadAll()
newstr = re.Replace(line, ", ")
WScript.Echo "New string '" & newstr & "', original '" & line & "'"
s.Close
