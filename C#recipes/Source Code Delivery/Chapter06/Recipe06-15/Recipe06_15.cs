using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;


namespace Recipe06_15
{
    class Recipe06_15
    {
        static void Main(string[] args)
        {
            // load the xml tree from the sample file
            XElement rootElement = XElement.Load(@"..\..\store.xml");

            // select the name of elements who have a category ID of 16
            IEnumerable<string> catEnum = from elem in rootElement.Elements()
                                          where (elem.Name == "Products" &&  
                                            ((string)elem.Element("CategoryID")) == "16")
                                          select ((string)elem.Element("ModelName"));

            foreach (string stringVal in catEnum)
            {
                Console.WriteLine("Category 16 item: {0}", stringVal);
            }

            Console.WriteLine("Press enter to proceed");
            Console.ReadLine();

            // perform the select again using instance methods
            IEnumerable<string> catEnum2 = rootElement.Elements().Where(e => e.Name == "Products" 
                && (string)e.Element("CategoryID") == "16").Select(e => (string)e.Element("ModelName"));

            foreach (string stringVal in catEnum2)
            {
                Console.WriteLine("Category 16 item: {0}", stringVal);
            }
        }
    }
}
