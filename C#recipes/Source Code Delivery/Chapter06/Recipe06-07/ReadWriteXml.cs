using System;
using System.Xml;
using System.IO;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter06
{
	public class ReadWriteXml
	{
		private static void Main()
		{
			// Create the file and writer.
			FileStream fs = new FileStream("products.xml", FileMode.Create);

            // If you want to configure additional details (like indenting,
            // encoding, and new line handling) use the overload of the Create
            // method that accepts an XmlWriterSettings object instead.
            XmlWriter w = XmlWriter.Create(fs); 

			// Start the document.
			w.WriteStartDocument();
			w.WriteStartElement("products");

			// Write a product.
			w.WriteStartElement("product");
			w.WriteAttributeString("id", "1001");	
			w.WriteElementString("productName", "Gourmet Coffee");
			w.WriteElementString("productPrice", "0.99");
			w.WriteEndElement();

			// Write another product.
			w.WriteStartElement("product");
			w.WriteAttributeString("id", "1002");	
			w.WriteElementString("productName", "Blue China Tea Pot");
			w.WriteElementString("productPrice", "102.99");
			w.WriteEndElement();

			// End the document.
			w.WriteEndElement();
			w.WriteEndDocument();
			w.Flush();
			fs.Close();
				
			Console.WriteLine("Document created. Press Enter to read the document.");
			Console.ReadLine();

			fs = new FileStream("products.xml", FileMode.Open);
            
            // If you want to configure additional details (like comments,
            // whitespace handling, or validation) use the overload of the Create
            // method that accepts an XmlReaderSettings object instead.
            XmlReader r = XmlReader.Create(fs);
			
			// Read all nodes.
			while (r.Read())
			{
				if (r.NodeType == XmlNodeType.Element)
				{
					Console.WriteLine();
					Console.WriteLine("<" + r.Name + ">");
					if (r.HasAttributes)
					{
						for (int i = 0; i < r.AttributeCount; i++)
						{
							Console.WriteLine("\tATTRIBUTE: " + r.GetAttribute(i));
                        }
					}
				}
				else if (r.NodeType == XmlNodeType.Text)
				{                    
					Console.WriteLine("\tVALUE: " + r.Value);
				}
			}
			Console.ReadLine();
		}
	}
}
