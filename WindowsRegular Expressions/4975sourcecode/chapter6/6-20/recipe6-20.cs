using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex( @"^\s*Dim\s+(?<var>\w+)\s+As\s+(?<type>[\w.]+)", RegexOptions.IgnoreCase );
	
	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				if (_Regex.IsMatch(line))
				{
					Console.WriteLine("Found variable:  '{0}'; type: '{1}'", 
						_Regex.Match(line).Result("${var}"),
						_Regex.Match(line).Result("${type}"));
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

