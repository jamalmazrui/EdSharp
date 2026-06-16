using System;
using System.Xml;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	/// <summary>
	/// Summary description for Class1.
	/// </summary>
	public class GenerateXml
	{
		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		private static void Main(string[] args)
		{
			XmlDocument doc = new XmlDocument();
			XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
			doc.AppendChild(docNode);

			// Create and insert a new element.
			XmlNode productsNode = doc.CreateElement("products");
			doc.AppendChild(productsNode);

			// Create a nested element (with an attribute).
			XmlNode productNode = doc.CreateElement("product");
			XmlAttribute productAttribute = doc.CreateAttribute("id");
			productAttribute.Value = "1001";
			productNode.Attributes.Append(productAttribute);
	        productsNode.AppendChild(productNode);

            // Create and add the sub-elements for this product node
			// (with contained text data).
			XmlNode nameNode = doc.CreateElement("productName");
			nameNode.AppendChild(doc.CreateTextNode("Gourmet Coffee"));
			productNode.AppendChild(nameNode);
			XmlNode priceNode = doc.CreateElement("productPrice");
			priceNode.AppendChild(doc.CreateTextNode("0.99"));
			productNode.AppendChild(priceNode);

			// Create and add another product node.
			productNode = doc.CreateElement("product");
			productAttribute = doc.CreateAttribute("id");
			productAttribute.Value = "1002";
			productNode.Attributes.Append(productAttribute);
			productsNode.AppendChild(productNode);
			nameNode = doc.CreateElement("productName");
			nameNode.AppendChild(doc.CreateTextNode("Blue China Tea Pot"));
			productNode.AppendChild(nameNode);
			priceNode = doc.CreateElement("productPrice");
			priceNode.AppendChild(doc.CreateTextNode("102.99"));
			productNode.AppendChild(priceNode);

        	// Save the document (to the Console window rather than a file).
			doc.Save(Console.Out);
			Console.ReadLine();
			
		}
	}
}
