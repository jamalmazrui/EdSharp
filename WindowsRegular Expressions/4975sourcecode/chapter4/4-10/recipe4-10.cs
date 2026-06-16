using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static string octetRegex = "([1-9]?[0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])";
	private static Regex _Regex = 
		new Regex(String.Format(@"^{0}\.{0}\.{0}\.{0}$", octetRegex));
	
	public void Run(string fileName)
	{
		String line;
		int lineNbr = 0;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				lineNbr++;
				if (_Regex.IsMatch(line))
				{
					Console.WriteLine("Found match '{0}' at line {1}", 
						line, 
						lineNbr);
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
