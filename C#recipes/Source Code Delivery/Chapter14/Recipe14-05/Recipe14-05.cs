using System;
using Microsoft.Win32;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_05
    {
        public static void SearchSubKeys(RegistryKey root, String searchKey)
        {
            try
            {
                // get the subkeys contained in the current key
                string[] subkeys = root.GetSubKeyNames();

                // Loop through all subkeys contained in the current key.
                foreach (string keyname in subkeys)
                {
                    try
                    {
                        using (RegistryKey key = root.OpenSubKey(keyname))
                        {
                            if (keyname == searchKey) PrintKeyValues(key);
                            SearchSubKeys(key, searchKey);
                        }
                    }
                    catch (System.Security.SecurityException)
                    {
                        // Ignore SecurityException for the purpose of the example.
                        // Some subkeys of HKEY_CURRENT_USER are secured and will
                        // throw a SecurityException when opened.
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                // ignore UnauthorizedAccessException for the purpose of the example
                // - this exception is thrown if the user does not have the 
                // rights to read part of the registry
            }

        }

        public static void PrintKeyValues(RegistryKey key)
        {
            // Display the name of the matching subkey and the number of
            // values it contains.
            Console.WriteLine("Registry key found : {0} contains {1} values",
                key.Name, key.ValueCount);

            // Loop through the values and display. 
            foreach (string valuename in key.GetValueNames())
            {
                if (key.GetValue(valuename) is String)
                {
                    Console.WriteLine("  Value : {0} = {1}",
                        valuename, key.GetValue(valuename));
                }
            }
        }

        public static void Main(String[] args)
        {
            if (args.Length > 0)
            {
                // Open the CurrentUser base key.
                using (RegistryKey root = Registry.CurrentUser)
                {
                    // Search recursively through the registry for any keys
                    // with the specified name.
                    SearchSubKeys(root, args[0]);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
