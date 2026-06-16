using System;
using System.Runtime.InteropServices;
using System.Text;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    class Recipe12_01
    {
        // Declare the unmanaged functions.
        [DllImport("kernel32.dll", EntryPoint = "GetPrivateProfileString")]
        private static extern int GetPrivateProfileString(string lpAppName,
          string lpKeyName, string lpDefault, StringBuilder lpReturnedString,
          int nSize, string lpFileName);

        [DllImport("kernel32.dll", EntryPoint = "WritePrivateProfileString")]
        private static extern bool WritePrivateProfileString(string lpAppName,
          string lpKeyName, string lpString, string lpFileName);

        static void Main(string[] args)
        {
            // Must use full path or Windows will try to write the INI file
            // to the Windows folder causing issues on Vista and Windows 7.
            string iniFileName = Path.Combine(Directory.GetCurrentDirectory(),
                "Recipe12-01.ini");

            string message = "Value of LastAccess in [SampleSection] is: {0}";

            // Write a new value to the INI file.
            WriteIniValue("SampleSection", "LastAccess",
                DateTime.Now.ToString(), iniFileName);
            
            // Obtain the value contained in the INI file.
            string val = GetIniValue("SampleSection", "LastAccess", iniFileName);
            Console.WriteLine(message, val ?? "???");

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Press Enter to continue the example.");
            Console.ReadLine();

            // Update the INI file.
            WriteIniValue("SampleSection", "LastAccess",
                DateTime.Now.ToString(), iniFileName);

            // Obtain the new value.
            val = GetIniValue("SampleSection", "LastAccess", iniFileName);
            Console.WriteLine(message, val ?? "???");

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }

        public static string GetIniValue(string section, string key, 
            string filename)
        {
            int chars = 256;
            StringBuilder buffer = new StringBuilder(chars);
            string sDefault = "";
            if (GetPrivateProfileString(section, key, sDefault,
              buffer, chars, filename) != 0)
            {
                return buffer.ToString();
            }
            else
            {
                // Look at the last Win32 error.
                int err = Marshal.GetLastWin32Error();
                return null;
            }
        }

        public static bool WriteIniValue(string section, string key, 
            string value, string filename)
        {
            return WritePrivateProfileString(section, key, value, filename);
        }
    }
}
