using System;
using System.Xml;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_08
    {
        public static void ConnectedExample()
        {
            // Create a new SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Create and configure a new command that includes the
                // FOR XML AUTO clause.
                using (SqlCommand com = con.CreateCommand())
                {
                    com.CommandType = CommandType.Text;
                    com.CommandText = "SELECT CustomerID, CompanyName" +
                        " FROM Customers FOR XML AUTO";

                    // Open the database connection.
                    con.Open();

                    // Execute the command and retrieve an XmlReader to access
                    // the results.
                    using (XmlReader reader = com.ExecuteXmlReader())
                    {
                        while (reader.Read())
                        {
                            Console.Write("Element: " + reader.Name);
                            if (reader.HasAttributes)
                            {
                                for (int i = 0; i < reader.AttributeCount; i++)
                                {
                                    reader.MoveToAttribute(i);
                                    Console.Write("  {0}: {1}",
                                        reader.Name, reader.Value);
                                }

                                // Move the XmlReader back to the element node.
                                reader.MoveToElement();
                                Console.WriteLine(Environment.NewLine);
                            }
                        }
                    }
                }
            }
        }

        public static void DisconnectedExample()
        {
            XmlDocument doc = new XmlDocument();

            // Create a new SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Create and configure a new command that includes the
                // FOR XML AUTO clause.
                SqlCommand com = con.CreateCommand();
                com.CommandType = CommandType.Text;
                com.CommandText =
                    "SELECT CustomerID, CompanyName FROM Customers FOR XML AUTO";

                // Open the database connection.
                con.Open();

                // Load the XML data into the XmlDocument. Must first create a 
                // root element into which to place each result row element.
                XmlReader reader = com.ExecuteXmlReader();
                doc.LoadXml("<results></results>");

                // Create an XmlNode from the next XML element read from the 
                // reader.
                XmlNode newNode = doc.ReadNode(reader);

                while (newNode != null)
                {
                    doc.DocumentElement.AppendChild(newNode);
                    newNode = doc.ReadNode(reader);
                }
            }

            // Process the disconnected XmlDocument.
            Console.WriteLine(doc.OuterXml);
        }

        public static void Main(string[] args)
        {
            ConnectedExample();
            Console.WriteLine(Environment.NewLine);

            DisconnectedExample();
            Console.WriteLine(Environment.NewLine);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
