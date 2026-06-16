using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex( @"^(?<key>[^#=]+)\s*=\s*(?<value>[^#]*)\s*" );
	
	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				if (_Regex.IsMatch(line))
				{
					Console.WriteLine("Found key:  '{0}'; value: '{1}'", 
						_Regex.Match(line).Result("${key}"),
						_Regex.Match(line).Result("${value}"));
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
