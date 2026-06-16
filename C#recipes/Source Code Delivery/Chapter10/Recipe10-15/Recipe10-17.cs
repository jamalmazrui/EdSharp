using System;
using System.Net;
using System.ServiceModel.Syndication;
using System.Xml.Linq;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_15
    {
        static void Main(string[] args)
        {
            Uri feedUrl = null;

            if (args.Length == 0 || String.IsNullOrEmpty(args[0])
                || !Uri.TryCreate(args[0], UriKind.RelativeOrAbsolute, 
                    out feedUrl))
            {
                // Error and wait to continue.
                Console.WriteLine("Invalid feed URL. Press Enter.");
                Console.ReadLine(); 
                return;
            }

            // Create the web request based on the URL provided for the feed.
            WebRequest req = WebRequest.Create(feedUrl);

            // Get the data from the feed.
            WebResponse res = req.GetResponse();

            // Simple test for the type of feed: Atom 1.0 or RSS 2.0.
            SyndicationFeedFormatter formatter = null;
            XElement feed = XElement.Load(res.GetResponseStream());

            if (feed.Name.LocalName == "rss")
            {
                formatter = new Rss20FeedFormatter();
            }
            else if (feed.Name.LocalName == "feed")
            {
                formatter = new Atom10FeedFormatter();
            }
            else
            {
                // Error and wait to continue.
                Console.WriteLine("Unsupported feed type: "
                    + feed.Name.LocalName);
                Console.ReadLine();
                return;
            }

            // Read the feed data into the formatter.
            formatter.ReadFrom(feed.CreateReader());

            // Display feed level data:
            Console.WriteLine("Title: " + formatter.Feed.Title.Text);
            Console.WriteLine("Description: " 
                + formatter.Feed.Description.Text);
            Console.Write(Environment.NewLine);
            Console.WriteLine("Items: ");

            // Display the item data.
            foreach (var item in formatter.Feed.Items)
            {
                Console.WriteLine("\tTitle: " + item.Title.Text);
                Console.WriteLine("\tSummary: " + item.Summary.Text);
                Console.WriteLine("\tPublish Date: " + item.PublishDate);
                Console.Write(Environment.NewLine);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
