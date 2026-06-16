using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;

namespace Recipe06_13
{
    class Recipe06_13
    {
        static void Main(string[] args)
        {
            // define the path to the sample file
            string filename = @"..\..\ProductCatalog.xml";

            // load the XML using the file name
            Console.WriteLine("Loading using file name");
            XElement root = XElement.Load(filename);
            // write out the xml
            Console.WriteLine(root);
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

            // load via a stream to the file
            Console.WriteLine("Loading using a stream");
            FileStream filestream = File.OpenRead(filename);
            root = XElement.Load(filestream);
            // write out the xml
            Console.WriteLine(root);
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

            // load via a textreader
            Console.WriteLine("Loading using a TextReader");
            TextReader reader = new StreamReader(filename);
            root = XElement.Load(reader);
            // write out the xml
            Console.WriteLine(root);
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

            // load via an xmlreader
            Console.WriteLine("Loading using an XmlReader");
            XmlReader xmlreader = new XmlTextReader(new StreamReader(filename));
            root = XElement.Load(xmlreader);
            // write out the xml
            Console.WriteLine(root);
            Console.WriteLine("Press enter to continue");
            Console.ReadLine();

        }
    }
}
