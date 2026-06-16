using System;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_16
    {
        private static string defualtUrl 
            = "http://www.apress.com:80/book/view/9781430225256";
        
        static void Main(string[] args)
        {
            Uri uri = null;

            // Extract information from a string URL passed as a 
            // command line argument or use the default URL.
            string strUri = defualtUrl;

            if (args.Length > 0 && !String.IsNullOrEmpty(args[0]))
            {
                strUri = args[0];
            }

            // Safely parse the url
            if (Uri.TryCreate(strUri, UriKind.RelativeOrAbsolute, out uri))
            {
                Console.WriteLine("Parsed URI: " + uri.OriginalString);
                Console.WriteLine("\tScheme: " + uri.Scheme);
                Console.WriteLine("\tHost: " + uri.Host);
                Console.WriteLine("\tPort: " + uri.Port);
                Console.WriteLine("\tPath and Query: " + uri.PathAndQuery);
            }
            else
            {
                Console.WriteLine("Unable to parse URI: " + strUri);
            }

            // Create a new URI.
            UriBuilder newUri = new UriBuilder();
            newUri.Scheme = "http";
            newUri.Host = "www.apress.com";
            newUri.Port = 80;
            newUri.Path = "book/view/9781430225256";

            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Created URI: " + newUri.Uri.AbsoluteUri);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
