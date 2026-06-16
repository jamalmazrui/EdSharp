using System;
using System.IO;
using System.Collections;
using System.Runtime.Serialization.Formatters.Soap;
using System.Runtime.Serialization.Formatters.Binary;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_13
    {
        // Serialize an ArrayList object to a binary file.
        private static void BinarySerialize(ArrayList list)
        {
            using (FileStream str = File.Create("people.bin"))
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(str, list);
            }
        }

        // Deserialize an ArrayList object from a binary file.
        private static ArrayList BinaryDeserialize()
        {
            ArrayList people = null;

            using (FileStream str = File.OpenRead("people.bin"))
            {
                BinaryFormatter bf = new BinaryFormatter();
                people = (ArrayList)bf.Deserialize(str);
            }
            return people;
        }

        // Serialize an ArrayList object to a SOAP file.
        private static void SoapSerialize(ArrayList list)
        {
            using (FileStream str = File.Create("people.soap"))
            {
                SoapFormatter sf = new SoapFormatter();
                sf.Serialize(str, list);
            }
        }

        // Deserialize an ArrayList object from a SOAP file.
        private static ArrayList SoapDeserialize()
        {
            ArrayList people = null;

            using (FileStream str = File.OpenRead("people.soap"))
            {
                SoapFormatter sf = new SoapFormatter();
                people = (ArrayList)sf.Deserialize(str);
            }
            return people;
        }

        public static void Main()
        {
            // Create and configure the ArrayList to serialize
            ArrayList people = new ArrayList();
            people.Add("Graeme");
            people.Add("Lin");
            people.Add("Andy");

            // Serialize the list to a file in both binary and SOAP form.
            BinarySerialize(people);
            SoapSerialize(people);

            // Rebuild the lists of people from the binary and SOAP
            // serializations and display them to the console.
            ArrayList binaryPeople = BinaryDeserialize();
            ArrayList soapPeople = SoapDeserialize();

            Console.WriteLine("Binary people:");
            foreach (string s in binaryPeople)
            {
                Console.WriteLine("\t" + s);
            }

            Console.WriteLine("\nSOAP people:");
            foreach (string s in soapPeople)
            {
                Console.WriteLine("\t" + s);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
