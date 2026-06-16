Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+[AP]M)\s+((\<DIR\>)\s+)?([-\w\s.]+)$"
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	If re.Test(line) = True Then
		newstr = re.Replace(line, "Found directory: '$4' with timestamp: '$1'")
		WScript.Echo newstr
	End If
Loop
s.Close
