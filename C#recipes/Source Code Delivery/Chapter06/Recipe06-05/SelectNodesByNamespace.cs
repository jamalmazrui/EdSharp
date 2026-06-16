using System;
using System.Xml;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	public class SelectNodesByNamespace
	{
		[STAThread]
		private static void Main()
		{
			// Load the document.
			XmlDocument doc = new XmlDocument();
			doc.Load("Order.xml");

			// Retrieve all order tags.
			XmlNodeList matches = doc.GetElementsByTagName("*", "http://mycompany/OrderML");

			// Display all the information.
			Console.WriteLine("Element \tAttributes");
			Console.WriteLine("******* \t**********");
			foreach (XmlNode node in matches)
			{
				Console.Write(node.Name + "\t");
				foreach (XmlAttribute attribute in node.Attributes)
				{
					Console.Write(attribute.Value + "  ");
				}
				Console.WriteLine();
			}
 
			Console.ReadLine();
		}
	}
}
