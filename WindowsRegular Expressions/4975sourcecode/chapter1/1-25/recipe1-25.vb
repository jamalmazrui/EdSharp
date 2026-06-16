Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = New Regex("\r\n")

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim newLine As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadToEnd
        If Not line Is Nothing
            newLine = _Regex.Replace(line, ", ")
            Console.WriteLine("New string is: '{0}', original was: '{1}'", _
             newLine, _
             line)
            line = sr.ReadLine
        End If
        sr.Close()
    End Sub

    Public Shared Sub Main(ByVal args As String())
        Dim r As Recipe = New Recipe
        r.Run(args(0))
    End Sub

End Class
