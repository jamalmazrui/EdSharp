using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{
	private static string nameSpaceRe = @"(?<ns>[a-z_][a-z\d_]+(\.[a-z_][a-z\d_]+)*)";
	private static string typeRe = @"(?<t>[a-z_][a-z\d_]+(\.[a-z_][a-z\d_]+)*)";
	private static string verRe = @"(?<v>Version=(\d+\.){3}\d+)";
	private static string cultureRe = "(?<c>Culture=[^,]+)";
	private static string tokenRe = @"(,\s*(?<pkt>PublicKeyToken=([a-f\d]{16}|null)))?";
	private static Regex _Regex = 
		new Regex(String.Format(@"^{0},\s*{1},\s*{2},\s*{3}{4}",
			nameSpaceRe,
			typeRe,
			verRe,
			cultureRe,
			tokenRe), RegexOptions.IgnoreCase);
	
	public void Run(string fileName)
	{
		String line;
		int lineNbr = 0;
		using (StreamReader sr = new StreamReader(fileName))
		{
			while(null != (line = sr.ReadLine()))
			{
				lineNbr++;
				foreach (Match m in _Regex.Matches(line))
				{
					Console.WriteLine("Found match '{0}' at line {1}", 
						m.Result("${1}"), 
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
