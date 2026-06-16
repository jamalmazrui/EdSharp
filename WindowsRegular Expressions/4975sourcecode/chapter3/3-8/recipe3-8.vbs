Public Function EscapeField(s)
	Dim tmpStr,myRe
	Set myRe = New RegExp
	myRe.Pattern = """"
	myRe.Global = True
	tmpStr = myRe.Replace(Trim(s),"""""")
	myRe.Pattern = "([^\t]*,[^\t]*)"
	EscapeField = myRe.Replace(tmpStr, """$1""")
End Function

Dim fso,s,re,line
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^(.{10})(.{20})(.{15})"
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	WScript.Echo EscapeField(re.Replace(line,"$1")) & _
		"," & EscapeField(re.Replace(line,"$2")) & _
		"," & EscapeField(re.Replace(line,"$3"))
Loop
s.Close
