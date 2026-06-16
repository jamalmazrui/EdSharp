Imports System
Imports System.IO
Imports System.Text.RegularExpressions

Public Class Recipe

	Private Shared _Regex As Regex = New Regex("^(?<f1>.{10})(?<f2>.{20})(?<f3>.{15})")

	Public Sub Run(ByVal fileName As String)
		Dim line As String
		Dim sr As StreamReader = File.OpenText(fileName)
		line = sr.ReadLine
		While Not line Is Nothing
			If _Regex.IsMatch(line) Then
				Console.WriteLine("{0},{1},{2}", _
					EscapeField(_Regex.Match(line).Result("${f1}")), _
					EscapeField(_Regex.Match(line).Result("${f2}")), _
					EscapeField(_Regex.Match(line).Result("${f3}")))
			End If
			line = sr.ReadLine
		End While
		sr.Close()
	End Sub

	Private Function EscapeField(ByVal s As String) As String
		Dim tmpStr As String = Regex.Replace(s.Trim, """", """""")
		Return Regex.Replace(tmpStr, "([^\t]*,[^\t]*)", """$1""")
	End Function

	Public Shared Sub Main(ByVal args As String())
		Dim r As Recipe = New Recipe
		r.Run(args(0))
	End Sub
End Class 
