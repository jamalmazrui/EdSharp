Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("^(00[1-9]|0[1-9]\d|[1-5]\d{2}|6[0-5]\d|66[0-5]|66[7-9]|6[7-8]\d|690|7[0-2]\d|73[0-3]|750|76[4-9]|77[0-2])-(?!00)\d{2}-(?!0000)\d{4}$")

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
