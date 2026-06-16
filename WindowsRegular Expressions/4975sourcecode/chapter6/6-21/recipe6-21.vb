Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("^(?<key>[^#=]+)\s*=\s*(?<value>[^#]*)\s*")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim lineNbr As Integer = 0
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
			If _Regex.IsMatch(line) = True Then
				Console.WriteLine("Found key:  '{0}'; value: '{1}'", _
					_Regex.Match(line).Result("${key}"), _
					_Regex.Match(line).Result("${value}"))
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
