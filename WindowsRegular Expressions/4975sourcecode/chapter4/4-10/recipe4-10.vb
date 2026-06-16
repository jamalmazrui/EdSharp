Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

	Private Const octetRegex = "([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])"
    Private Shared _Regex As Regex = _
    	New Regex(String.Format("^{0}\.{0}\.{0}\.{0}$", octetRegex))

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
