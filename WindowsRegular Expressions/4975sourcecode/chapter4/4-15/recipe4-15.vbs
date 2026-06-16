Dim fso,s,re,line,lineNbr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "^(0?2/(0?[1-9]|[12][0-9])|~CCC
(0?[469]|11)/(0?[1-9]|[12][0-9]|30)|~CCC
(0?[13578]|1[02])/(0?[1-9]|[12][0-9]|3[01]))/\d{4}$"
lineNbr = 0
Do While Not s.AtEndOfStream
    line = s.ReadLine()
    lineNbr = lineNbr + 1
    If re.Test(line) Then
        WScript.Echo "Found match: '" & line & "' at line " & lineNbr
    End If
Loop
s.Close
