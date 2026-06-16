Dim fso,s,re,line,newstr
Set fso = CreateObject("Scripting.FileSystemObject")
Set s = fso.OpenTextFile(WScript.Arguments.Item(0), 1, True)
Set re = New RegExp
re.Pattern = "([[<]assembly: assemblyversion\("")1\.0\.\*(?="")"
re.IgnoreCase = True
Do While Not s.AtEndOfStream
    line = s.ReadLine()
    newstr = re.Replace(line, "$12.1.123.1234")
    WScript.Echo "New string '" & newstr & "', original '" & line & "'"
Loop
s.Close
