using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{

	private static string month29 = "(0[1-9]|[12][0-9])-Feb";	
	private static string month30 = "(0[1-9]|[12][0-9]|30)-(Sep|Apr|Jun|Nov)";
	private static string month31 = "(0[1-9]|[12][0-9]|3[01])-(Jan|Mar|May|Jul|Aug|Oct|Dec)";

    private static Regex _Regex = new Regex(String.Format("({0}|{1}|{2})",
	                                                      month29,
	                                                      month30,
	                                                      month31
	                                                      ) + @"-\d{4}");
	
	public void Run(string fileName)
	{
		Console.WriteLine("Full regex is:  '{0}'", _Regex.ToString());
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
