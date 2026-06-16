Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("^(?<ts>\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+[AP]M)\s+((?<type>\<DIR\>)\s+)?(?<name>[-\w\s.]+)$")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim newLine As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
			If _Regex.IsMatch(line) = True Then
				Console.WriteLine("Found directory:  '{0}' with timestamp: '{1}'", _
					_Regex.Match(line).Result("${name}"), _
					_Regex.Match(line).Result("${ts}"))
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
