using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Recipe06_16
{
    class Recipe06_16
    {
        static void Main(string[] args)
        {
            // load the xml tree from the file
            XElement rootElem = XElement.Load(@"..\..\ProductCatalog.xml");

            // write out the xml
            Console.WriteLine(rootElem);

            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

            // select all of the product elements
            IEnumerable<XElement> prodElements 
                = from elem in rootElem.Element("products").Elements()
                    where (elem.Name == "product")
                    select elem;

            // run through the elements and change the ID attribute
            foreach(XElement elem in prodElements) 
            {
                // get the current product ID
                int current_id = Int32.Parse((string)elem.Attribute("id"));
                // perform the replace operation on the attribute
                elem.ReplaceAttributes(new XAttribute("id", current_id + 500));
            }

            Console.WriteLine(rootElem);
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

            // remove all elements that contain the word "tea" in the description
            IEnumerable<XElement> teaElements = from elem in rootElem.Element("products").Elements()
                    where (((string)elem.Element("description")).Contains("tea"))
                    select elem;

            foreach (XElement elem in teaElements)
            {
                elem.Remove();
            }

            Console.WriteLine(rootElem);
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

            // define and add a new element
            XElement newElement = new XElement("product",
                new XAttribute("id", 3000),
                new XElement("productName", "Chrome French Press"),
                new XElement("description", 
                    "A simple and elegant way of making great coffee"),
                new XElement("productPrice", 25.00),
                new XElement("inStock", true));

            rootElem.Element("products").Add(newElement);

            Console.WriteLine(rootElem);
        }
    }
}
