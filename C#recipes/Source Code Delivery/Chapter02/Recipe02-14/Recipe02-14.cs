using System;
using System.Collections.Generic;
using System.Runtime.Serialization.Json;
using System.IO;
using System.Text;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_14
    {
        public static void Main()
        {
    
            // create a list of strings
            List<string> myList = new List<string>()
            {
                "apple", "orange", "banana", "cherry"
            };

            // create memory stream - we will use this
            // to get the JSON serialization as a string
            MemoryStream memoryStream = new MemoryStream();

            // create the JSON serializer
            DataContractJsonSerializer jsonSerializer 
                = new DataContractJsonSerializer(myList.GetType());

            // serialize the list
            jsonSerializer.WriteObject(memoryStream, myList);

            // get the JSON string from the memory stream
            string jsonString = Encoding.Default.GetString(memoryStream.ToArray());

            // write the string to the console
            Console.WriteLine(jsonString);

            // create a new stream so we can read the JSON data
            memoryStream = new MemoryStream(Encoding.Default.GetBytes(jsonString));

            // deserialize the list
            myList = jsonSerializer.ReadObject(memoryStream) as List<string>;

            // enumerate the strings in the list
            foreach (string strValue in myList)
            {
                Console.WriteLine(strValue);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
