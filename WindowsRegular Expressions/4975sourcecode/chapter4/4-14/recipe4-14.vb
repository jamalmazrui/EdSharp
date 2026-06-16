Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("^(\d{4}-){3}|(\d{4} ){3}\d{4}|\d{15,16}|\d{4} \d{2} \d{4} \d{5}|\d{4}-\d{2}-\d{4}-\d{5}$")

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
