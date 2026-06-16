using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static string month29 = "0?2/(0?[1-9]|[12][0-9])";	
	private static string month30 = "(0?[469]|11)/(0?[1-9]|[12][0-9]|30)";
	private static string month31 = "(0?[13578]|1[02])/(0?[1-9]|[12][0-9]|3[01])";
	private static Regex _Regex = new Regex(String.Format("^({0}|{1}|{2})",
														  month29,
														  month30,
														  month31
														  ) + @"/\d{4}$");
		
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
