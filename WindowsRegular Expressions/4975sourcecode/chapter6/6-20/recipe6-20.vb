Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("^\s*Dim\s+(?<var>\w+)\s+As\s+(?<type>[\w.]+)", RegexOptions.IgnoreCase)

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim newLine As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
			If _Regex.IsMatch(line) = True Then
				Console.WriteLine("Found variable:  '{0}'; type: '{1}'", _
					_Regex.Match(line).Result("${var}"), _
					_Regex.Match(line).Result("${type}"))
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
