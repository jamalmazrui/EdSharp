using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex(@"^(?<f1>.{10})(?<f2>.{20})(?<f3>.{15})");

	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				if (_Regex.IsMatch(line))
				{
					Console.WriteLine("{0},{1},{2}",
						EscapeField(_Regex.Match(line).Result("${f1}")),
						EscapeField(_Regex.Match(line).Result("${f2}")),
						EscapeField(_Regex.Match(line).Result("${f3}")));
				}
			}
		}
	}
	
	private string EscapeField(string s)
	{
		string tmpStr = Regex.Replace(s.Trim(),"\"", "\"\"");
		return Regex.Replace(tmpStr, @"([^\t]*,[^\t]*)", "\"$1\"");
	}

	public static void Main( string[] args )
	{
		Recipe r = new Recipe();
		r.Run(args[0]);
	}
}
