using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static string hostRegex = @"([a-z\d][-a-z\d]*[a-z\d]\.)*[a-z][-a-z\d]*[a-z]";
	private static string portRegex = @"(:\d{1,})?";
	private static string pathRegex = @"(/[^\s?]+)?";
	private static string queryRegex = "(\\?[^<>#\"\\s]+)?";

	private static string fullRegex = @"(?:(?<=^)|(?<=\s))((ht|f)tps?://" + hostRegex + portRegex + pathRegex + queryRegex + ")" ;
	
	private static Regex _Regex = new Regex(fullRegex);
	
	public void Run(string fileName)
	{
		String line;
		String newLine;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				newLine = _Regex.Replace(line, "<a href=\"$1\">$1</a>");
				Console.WriteLine("New string is: '{0}', original was: '{1}'",
					newLine,
					line);
			}
		}
	}
	
	public static void Main( string[] args )
	{
		Recipe r = new Recipe();
		r.Run(args[0]);
	}
}

