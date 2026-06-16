Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

	Private Const hostRegex = "([a-z\d][-a-z\d]*[a-z\d]\.)*[a-z][-a-z\d]*[a-z]"
	Private Const portRegex = "(:\d{1,})?"
	Private Const pathRegex = "(/[^?<>#""\s]+)?"
	Private Const queryRegex = "(\?[^<>#""\s]+)?"

    Private Shared _Regex As Regex = New Regex("(?:(?<=^)|(?<=\s))((ht|f)tps?://" & hostRegex & portRegex & pathRegex & queryRegex & ")")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim newLine As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
            newLine = _Regex.Replace(line, "cow")
            Console.WriteLine("New string is: '{0}', original was: '{1}'", _
             newLine, _
             line)
            line = sr.ReadLine
        End While
        sr.Close()
    End Sub

    Public Shared Sub Main(ByVal args As String())
        Dim r As Recipe = New Recipe
        r.Run(args(0))
    End Sub

End Class
