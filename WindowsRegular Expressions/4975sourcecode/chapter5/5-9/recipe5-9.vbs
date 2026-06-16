Dim fso,s,re,line,matches,m
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "<[a-z:_][-a-z0-9._:]+((\s+[^>]+)*|/s*)/?>"
re.IgnoreCase = True
re.MultiLine = True
re.Global = True
line = s.ReadAll()
Set matches = re.Execute(line)
For Each m in matches
	WScript.Echo "Found match:  '" & m.Value & "'"
Next
s.Close
