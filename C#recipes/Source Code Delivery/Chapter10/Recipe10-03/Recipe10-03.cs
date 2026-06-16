using System;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_03 
    {
        private static void Main() 
        {
            // Specify the URI of the resource to parse.
            string srcUriString = "http://www.apress.com";
            Uri srcUri = new Uri(srcUriString);

            // Create a WebClient to perform the download.
            WebClient client = new WebClient();

            Console.WriteLine("Downloading {0}", srcUri);

            // Perform the download getting the resource as a string.
            string str = client.DownloadString(srcUri);

            // Use a regular expression to extract all HTML <img> 
            // elements and extract the path to any that reference
            // files with a gif, jpg, or jpeg extension.
            MatchCollection matches = Regex.Matches(str, 
                "<img.*?src\\s*=\\s*[\"'](?<url>.*?\\.(gif|jpg|jpeg)).*?>",
                RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // Try to download each referenced image file.
            foreach(Match match in matches) 
            {
                var urlGrp = match.Groups["url"];
                
                if (urlGrp != null && urlGrp.Success)
                {
                    Uri imgUri = null;

                    // Determine the source Uri.
                    if (urlGrp.Value.StartsWith("http"))
                    {
                        // Absolute.
                        imgUri = new Uri(urlGrp.Value);
                    }
                    else if (urlGrp.Value.StartsWith("/"))
                    {
                        // Relative
                        imgUri = new Uri(srcUri, urlGrp.Value);
                    }
                    else
                    {
                        // Skip it.
                        Console.WriteLine("Skipping image {0}", urlGrp.Value);
                    }

                    if (imgUri != null)
                    {
                        // Determine the local filename to use.
                        string fileName = 
                            urlGrp.Value.Substring(urlGrp.Value.LastIndexOf('/')+1);

                        try
                        {
                            // Download and store the file.
                            Console.WriteLine("Downloading {0} to {1}",
                                imgUri.AbsoluteUri, fileName);

                            client.DownloadFile(imgUri, fileName);
                        }
                        catch
                        {
                            Console.WriteLine("Failed to download {0}", 
                                imgUri.AbsoluteUri);
                        }
                    }
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
