using System;
using System.IO;
using System.Text.RegularExpressions;

public class Recipe
{

	private static Regex _Regex = 
		new Regex(@"<script language=""javascript"">([^<""]*|[^<]*<[^/][^<]*|(([^<""]|[^<]*<[^/][^<])*""[^""]*""([^<""]|[^<]*<[^/][^<])*)*)?</script>", 
		          RegexOptions.IgnoreCase | RegexOptions.Multiline);
	
	public void Run(string fileName)
	{
		String line;
		using (StreamReader sr = new StreamReader(fileName))
		{
			line = sr.ReadToEnd();
			foreach (Match m in _Regex.Matches(line))
			{
				Console.WriteLine("Found match '{0}'", 
					m.ToString());
			}
		}
	}	

	public static void Main( string[] args )
	{
		Recipe r = new Recipe();
		r.Run(args[0]);
	}
}

