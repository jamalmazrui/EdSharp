Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("(?<path>(\\(?<file>[^\\/:*?""<>|.][^\\/:*?""<>|]*))+)")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim newLine As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
        	If _Regex.IsMatch(line) Then
            	Console.WriteLine("Found file '{0}' in path '{1}'", _
            		_Regex.Match(line).Result("${file}"), _
            		_Regex.Match(line).Result("${path}"))
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
