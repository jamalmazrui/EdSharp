using System;
using System.Xml;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	public class FindNodesByName
	{
		[STAThread]
		private static void Main()
		{
			// Load the document.
			XmlDocument doc = new XmlDocument();
            doc.Load(@"..\..\ProductCatalog.xml");

            // Retrieve all prices.
			XmlNodeList prices = doc.GetElementsByTagName("productPrice");
			
			decimal totalPrice = 0;
            foreach (XmlNode price in prices)
			{
				// Get the inner text of each matching element.
				totalPrice += Decimal.Parse(price.ChildNodes[0].Value);
			}
			
			Console.WriteLine("Total catalog value: " + totalPrice.ToString());
        	Console.ReadLine();
		}
	}
}
