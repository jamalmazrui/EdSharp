Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
Private Const month29 = "(0[1-9]|[12][0-9])-Feb"
Private Const month30 = "(0[1-9]|[12][0-9]|30)-(Sep|Apr|Jun|Nov)"
Private Const month31 = "(0[1-9]|[12][0-9]|3[01])-(Jan|Mar|May|Jul|Aug|Oct|Dec)"
re.Pattern = "(" + month29 + "|" + month30 + "|" + month31 + ")-\d{4}"
lineNbr = 0
Do While Not s.AtEndOfStream
	line = s.ReadLine()
	lineNbr = lineNbr + 1
	If re.Test(line) Then
		WScript.Echo "Found match: '" & line & "' at line " & lineNbr
	End If
Loop
s.Close
