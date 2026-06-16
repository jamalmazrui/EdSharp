using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Recipe05_25
{
    class Recipe05_25
    {
        static void Main(string[] args)
        {
            // read all of the entries from the file
            IEnumerable<string> alldata = File.ReadAllLines("all_entries.log");
            foreach (string entry in alldata)
            {
                Console.WriteLine("Entry: {0}", entry);
            }
        
            // read the entries and select only some of them
            IEnumerable<string> somedata = File.ReadLines("all_entries.log").Where(e => e.StartsWith("Error"));
            foreach (string entry in somedata) 
            {
                Console.WriteLine("Error entry: {0}", entry);
            }

            // read selected lines and write only the first character
            IEnumerable<char> chardata = File.ReadLines("all_entries.log").Where(e => !e.StartsWith("Error")).Select(e => e[0]);
            foreach (char entry in chardata) 
            { 
                Console.WriteLine("Character entry: {0}", entry);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
