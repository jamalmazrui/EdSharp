Imports System
Imports System.IO
Imports System.Text.RegularExpressions
Public Class Recipe

	Private Const nameSpaceRe As String = _
		"(?<ns>[a-z_][a-z\d_]+(\.[a-z_][a-z\d_]+)*)"
	Private Const typeRe As String = "(?<t>[a-z_][a-z\d_]+(\.[a-z_][a-z\d_]+)*)"
	Private Const verRe As String = "(?<v>Version=(\d+\.){3}\d+)"
	Private Const cultureRe As String = "(?<c>Culture=[^,]+)"
	Private Const tokenRe As String = _
		"(,\s*(?<pkt>PublicKeyToken=([a-f\d]{16}|null)))?" 

	Private Shared _Regex As Regex = New Regex( _
	String.Format("^{0},\s*{1},\s*{2},\s*{3}{4}", _
			nameSpaceRe, _
			typeRe, _
			verRe, _
			cultureRe, _
			tokenRe), RegexOptions.IgnoreCase)

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
					m.Result("${1}"), _
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
