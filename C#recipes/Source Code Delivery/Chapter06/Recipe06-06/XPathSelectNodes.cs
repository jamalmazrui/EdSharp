using System;
using System.Xml;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	public class XPathSelectNodes
	{
		[STAThread]
		private static void Main()
		{
			// Load the document.
			XmlDocument doc = new XmlDocument();
			doc.Load("orders.xml");

			// Retrieve the name of every item.
			// This could not be accomplished as easily with the
			// GetElementsByTagName() method, because Name elements are
			// used in Item elements and Client elements, and so
			// both types would be returned.
			XmlNodeList nodes = doc.SelectNodes("/Order/Items/Item/Name");
            
			foreach (XmlNode node in nodes)
			{
				Console.WriteLine(node.InnerText);
			}
        
			Console.ReadLine();
		}
	}
}
