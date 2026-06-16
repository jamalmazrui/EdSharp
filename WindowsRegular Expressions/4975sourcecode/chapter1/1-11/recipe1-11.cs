using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex( @"\b(\w+)(\s*$\s*|\s+)\1\b", 
		RegexOptions.Multiline | RegexOptions.IgnoreCase );
	
	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			if(null != (line = sr.ReadToEnd()))
			{
				if (_Regex.IsMatch(line))
				{
					foreach (Match myMatch in _Regex.Matches(line))
					{
						Console.WriteLine("Found match '{0}'", myMatch.ToString());
					}
				}
			}
		}
	}

	public static void Main( string[] args )
	{
		Recipe r = new Recipe();
		r.Run(args[0]);
	}
}
