Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("\s+[-/](?<arg>[a-z0-9?]+)([:=](?<val>[^\s""]+|""[^""]*""))?", RegexOptions.IgnoreCase)

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim lineNbr As Integer = 0
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
            lineNbr = lineNbr + 1
            Dim m As Match
            For Each m In _Regex.Matches(line)
                Console.WriteLine("Found match '{0}' at line {1}", _
                    m.Result("${arg} with value ${val}"), _
                    lineNbr)
            Next m
            line = sr.ReadLine
        End While
        sr.Close()
    End Sub

    Public Shared Sub Main(ByVal args As String())
        Dim r As Recipe = New Recipe
        r.Run(args(0))
    End Sub

End Class
