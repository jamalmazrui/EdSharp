using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static Regex _Regex = new Regex( @">(?![^><]+?\/?<)" );
	
	private string ReverseString(string s)
	{
		char[] array = s.ToCharArray();
		Array.Reverse(array);
		return new String(array);
	}
	
	public void Run(string fileName)
	{
		String line;
		String newLine;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				newLine = _Regex.Replace(ReverseString(line), ";tl&");
				Console.WriteLine("New string is: '{0}', original was: '{1}'",
				  	ReverseString(newLine),
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

