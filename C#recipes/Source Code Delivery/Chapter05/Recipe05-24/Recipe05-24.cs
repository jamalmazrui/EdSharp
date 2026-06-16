using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace Recipe05_24
{
    class Recipe05_24
    {
        static void Main(string[] args)
        {
            // create a list and populate it
            List<string> myList = new List<string>();
            myList.Add("Log entry 1");
            myList.Add("Log entry 2");
            myList.Add("Log entry 3");
            myList.Add("Error entry 1");
            myList.Add("Error entry 2");
            myList.Add("Error entry 3");

            // write all of the entries to a file
            File.WriteAllLines("all_entries.log", myList);

            // only write out the errors
            File.WriteAllLines("only_errors.log", 
                myList.Where(e => e.StartsWith("Error")));
        }
    }
}
