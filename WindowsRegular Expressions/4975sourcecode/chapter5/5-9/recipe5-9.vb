Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

    Private Shared _Regex As Regex = _
    	New Regex("<script language=""javascript"">([^<""]*|[^<]*<[^/][^<]*|(([^<""]|[^<]*<[^/][^<])*""[^""]*""([^<""]|[^<]*<[^/][^<])*)*)?</script>", _
    		RegexOptions.IgnoreCase Or RegexOptions.Multiline)

    Public Sub Run(ByVal fileName As String)
        Dim line As String
        Dim sr As StreamReader = File.OpenText(fileName)
        line = sr.ReadToEnd()
        If Not line Is Nothing
            For Each m As Match In _Regex.Matches(line)
            	Console.WriteLine("Found match '{0}'", m.ToString())
            Next
        End If
        sr.Close()
    End Sub

    Public Shared Sub Main(ByVal args As String())
        Dim r As Recipe = New Recipe
        r.Run(args(0))
    End Sub

End Class
