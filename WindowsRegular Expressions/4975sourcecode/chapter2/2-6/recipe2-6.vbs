Dim fso,s,re,line,newstr
Private Const hostRegex = "([a-z\d][-a-z\d]*[a-z\d]\.)*[a-z][-a-z\d]*[a-z]"
Private Const portRegex = "(:\d{1,})?"
Private Const pathRegex = "(/[^?<>#""\s]+)?"
Private Const queryRegex = "(\?[^<>#""\s]+)?"
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "((ht|f)tps?://" & hostRegex & portRegex & pathRegex & queryRegex & ")"
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	If re.Test(line) Then
		newstr = re.Replace(line, "<a href=""$1"">$1</a>")
		WScript.Echo "New string '" & newstr & "', original '" & line & "'"
	End If
Loop
s.Close
