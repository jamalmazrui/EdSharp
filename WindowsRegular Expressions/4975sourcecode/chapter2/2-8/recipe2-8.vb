Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("\\\\(?<unc>[^~!@#$^&()=+\[\]{}\\|;:',<>/?]+)")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim newLine As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadLine
        While Not line Is Nothing
        	If (_Regex.IsMatch(line)) Then
            	Console.WriteLine("Captured value '{0}'", _Regex.Match(line).Result("${unc}"))
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
