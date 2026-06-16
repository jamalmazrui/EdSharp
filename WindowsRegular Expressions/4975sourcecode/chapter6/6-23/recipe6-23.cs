using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex( @"^(?<ts>\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2}\s+[AP]M)\s+((?<type>\<DIR\>)\s+)?(?<name>[-\w\s.]+)$" );
	
	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				if (_Regex.IsMatch(line))
				{
					Console.WriteLine("Found directory:  '{0}' with timestamp: '{1}'", 
						_Regex.Match(line).Result("${name}"),
						_Regex.Match(line).Result("${ts}"));
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

