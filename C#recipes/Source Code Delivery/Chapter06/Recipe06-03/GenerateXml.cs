using System;
using System.Xml;

namespace Recipe06_03
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
			// Create the basic document.
			XmlDocument doc = new XmlDocument();
			XmlNode docNode = doc.CreateXmlDeclaration("1.0", "UTF-8", null);
			doc.AppendChild(docNode);
			XmlNode productsNode = doc.CreateElement("products");
			doc.AppendChild(productsNode);

			// Add two products.
			XmlNode productNode = XmlHelper.AddElement("product", null, productsNode);
			XmlHelper.AddAttribute("id", "1001", productNode);
			XmlHelper.AddElement("productName", "Gourmet Coffee", productNode);
			XmlHelper.AddElement("productPrice", "0.99", productNode);
			
			productNode = XmlHelper.AddElement("product", null, productsNode);
			XmlHelper.AddAttribute("id", "1002", productNode);
			XmlHelper.AddElement("productName", "Blue China Tea Pot", productNode);
			XmlHelper.AddElement("productPrice", "102.99", productNode);
			
			// Save the document (to the Console window rather than a file).
			doc.Save(Console.Out);
			Console.ReadLine();
		}
	}
	public class XmlHelper
	{
		public static XmlNode AddElement(string tagName, string textContent, XmlNode parent)
		{
			XmlNode node = parent.OwnerDocument.CreateElement(tagName);
			parent.AppendChild(node);

			if (textContent != null)
			{
				XmlNode content = parent.OwnerDocument.CreateTextNode(textContent);
				node.AppendChild(content);
			}

			return node;
		}

		public static XmlNode AddAttribute(string attributeName, string textContent, XmlNode parent)
		{
			XmlAttribute attribute = parent.OwnerDocument.CreateAttribute(attributeName);
			attribute.Value = textContent;
			parent.Attributes.Append(attribute);

			return attribute;
		}
	}
}
