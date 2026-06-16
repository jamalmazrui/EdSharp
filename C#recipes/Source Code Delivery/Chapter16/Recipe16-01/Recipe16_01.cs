using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Data;

namespace Apress.VisualCSharpRecipes.Chapter16
{
    class Recipe16_01
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Using an array source");
            // create the source
            string[] array = createArray();
            // perform the query
            IEnumerable<string> arrayEnum = from e in array select e;
            // write out the elements
            foreach (string str in arrayEnum)
            {
                Console.WriteLine("Element {0}", str);
            }

            Console.WriteLine("\nUsing a collection source");
            // create the source
            ICollection<string> collection = createCollection();
            // perform the query
            IEnumerable<string> collEnum = from e in collection select e;
            // write out the elements
            foreach (string str in collEnum)
            {
                Console.WriteLine("Element {0}", str);
            }

            Console.WriteLine("\nUsing an xml source");
            // create the source
            XElement xmltree = createXML();
            // perform the query
            IEnumerable<XElement> xmlEnum = from e in xmltree.Elements() select e;
            // write out the elements
            foreach (string str in xmlEnum)
            {
                Console.WriteLine("Element {0}", str);
            }

            Console.WriteLine("\nUsing a data table source");
            // create the source
            DataTable table = createDataTable();

            // perform the query
            IEnumerable<string> dtEnum = from e in table.AsEnumerable() select e.Field<string>(0);
            // write out the elements
            foreach (string str in dtEnum)
            {
                Console.WriteLine("Element {0}", str);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }

        static string[] createArray()
        {
            return new string[] { "apple", "orange", "grape", "fig", "plum", "banana", "cherry" };
        }

        static IList<string> createCollection()
        {
            return new List<string>() { "apple", "orange", "grape", "fig", "plum", "banana", "cherry" };
        }

        static XElement createXML()
        {
            return new XElement("fruit",
                new XElement("name", "apple"),
                new XElement("name", "orange"),
                new XElement("name", "grape"),
                new XElement("name", "fig"),
                new XElement("name", "plum"),
                new XElement("name", "banana"),
                new XElement("name", "cherry")
            );
        }

        static DataTable createDataTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("name", typeof(string));
            string[] fruit = { "apple", "orange", "grape", "fig", "plum", "banana", "cherry" };
            foreach (string name in fruit)
            {
                table.Rows.Add(name);
            }
            return table;
        }
    }
}
