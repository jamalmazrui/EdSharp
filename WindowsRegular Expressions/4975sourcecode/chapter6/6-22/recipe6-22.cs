using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex( @"^(?<path>[A-Za-z]:\\[^/:*?""<>|]+)\((?<line>\d+),(?<pos>\d+)\):" );
	
	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				if (_Regex.IsMatch(line))
				{
					Console.WriteLine("Error in file:  '{0}' at line '{1}', position '{2}'", 
						_Regex.Match(line).Result("${path}"),
						_Regex.Match(line).Result("${line}"),
						_Regex.Match(line).Result("${pos}"));
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
