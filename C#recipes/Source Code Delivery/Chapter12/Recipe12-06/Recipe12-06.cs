using System;
using System.IO;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    class Recipe12_06
    {
        static void Main()
        {
            string fileName = 
                Path.Combine(Directory.GetCurrentDirectory(), 
                "Ranges.xlsx");

            // Create an instance of Excel.
            Console.WriteLine("Creating Excel instance...");
            Console.WriteLine(Environment.NewLine);
            Excel.Application excel = new Excel.Application();

            // Open the required file in Excel.
            Console.WriteLine("Opening file: {0}", fileName);
            Console.WriteLine(Environment.NewLine);

            // Open the specified file in Excel using .NET 4.0 optional 
            // and named argument capabilities.
            Excel.Workbook workbook =
                excel.Workbooks.Open(fileName, ReadOnly: true);

            /* Pre-.NET 4.0 syntax required to open Excel file:
            Excel.Workbook workbook = 
                excel.Workbooks.Open(fileName, Type.Missing, 
                false, Type.Missing, Type.Missing, Type.Missing, 
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, 
                Type.Missing); */

            // Display the list of named ranges from the file.
            Console.WriteLine("Named ranges:");
            foreach (Excel.Name name in workbook.Names)
            {
                Console.WriteLine("  {0} ({1})",name.Name,name.Value);
            }
            Console.WriteLine(Environment.NewLine);

            // Close the workbook.
            workbook.Close();

            /* Pre-.NET 4.0 syntax required to close Excel file:
            workbook.Close(Type.Missing, Type.Missing, Type.Missing); */

            // Terminate Excel instance.
            Console.WriteLine("Closing Excel instance...");
            excel.Quit();
            Marshal.ReleaseComObject(excel);
            excel = null;

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
