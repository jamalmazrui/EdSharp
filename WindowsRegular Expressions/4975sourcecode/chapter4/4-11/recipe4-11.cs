using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	// Any ASCII character except CTRL chars, space, and ()<>@,;:\".[]
	private static string localRegex = @"[\w\d!#$%&'*+-/=?^`{|}~]+(\.[\w\d!#$%&'*+-/=?^`{|}~]+)*";
	private static string hostNameRegex = @"([a-z\d][-a-z\d]*[a-z\d]\.)*[a-z][-a-z\d]*[a-z]";
	private static Regex _Regex = new Regex("^" + localRegex + "@" + hostNameRegex + "$", RegexOptions.IgnoreCase);
	
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
