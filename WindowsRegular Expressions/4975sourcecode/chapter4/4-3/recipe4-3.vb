Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

	Private Const month29 = "(0[1-9]|[12][0-9])-Feb"
	Private Const month30 = "(0[1-9]|[12][0-9]|30)-(Sep|Apr|Jun|Nov)"
	Private Const month31 = "(0[1-9]|[12][0-9]|3[01])-(Jan|Mar|May|Jul|Aug|Oct|Dec)"

    Private Shared _Regex As Regex = New Regex(String.Format("({0}|{1}|{2})", _
	                                               month29, _
	                                               month30, _
	                                               month31 _
	                                               ) + "-\d{4}")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim lineNbr As Integer = 0
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
            lineNbr = lineNbr + 1
            If _Regex.IsMatch(line) Then
                Console.WriteLine("Found match '{0}' at line {1}", _
                    line, _
                    lineNbr)
            End If
            line = sr.ReadLine
        End While
        sr.Close()
    End Sub

    Public Shared Sub Main(ByVal args As String())
        Dim r As Recipe = New Recipe
        r.Run(args(0))
    End Sub

End Class
