Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("\b(\w+)(\s*$?\s*|\s+)\1\b", _
    	RegexOptions.IgnoreCase Or RegexOptions.Multiline)

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadToEnd()
        If Not line Is Nothing
            If _Regex.IsMatch(line) Then
            	For Each myMatch As Match In _Regex.Matches(line)
					Console.WriteLine("Found match '{0}'", myMatch.ToString())
				Next
            End If
        End If
        sr.Close()
    End Sub

    Public Shared Sub Main(ByVal args As String())
        Dim r As Recipe = New Recipe
        r.Run(args(0))
    End Sub

End Class
