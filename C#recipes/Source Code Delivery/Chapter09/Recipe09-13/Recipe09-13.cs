using System;
using System.Data;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_13
    {
        static void Main(string[] args)
        {

            // create the data set
            DataSet dataset = new DataSet();

            // create the table and add it to the dataset
            DataTable table = new DataTable("Regions");
            dataset.Tables.Add(table);

            // create the colums for the table
            table.Columns.Add("RegionID", typeof(int));
            table.Columns.Add("RegionDescription", typeof(string));

            // populate the table
            string[] regions = { "Eastern", "Western", "Northern", "Southern" };
            for (int i = 0; i < regions.Length; i++)
            {
                DataRow row = table.NewRow();
                row["RegionID"] = i + 1;
                row["RegionDescription"] = regions[i];
                table.Rows.Add(row);
            }

            // print details of the schema
            Console.WriteLine("\nSchema for table");
            foreach (DataColumn col in table.Columns)
            {
                Console.WriteLine("Column: {0} Type: {1}", col.ColumnName, col.DataType);
            }

            // enumerate the data we have received
            Console.WriteLine("\nData in table");
            foreach (DataRow row in table.Rows)
            {
                Console.WriteLine("Data {0} {1}", row[0], row[1]);
            }

            // create a new row
            DataRow newrow = table.NewRow();
            newrow["RegionID"] = 5;
            newrow["RegionDescription"] = "Central";
            table.Rows.Add(newrow);

            // modify an existing row
            table.Rows[0]["RegionDescription"] = "North Eastern";

            // enumerate the cached data again
            // enumerate the data we have received
            Console.WriteLine("\nData in (modified) table");
            foreach (DataRow row in table.Rows)
            {
                Console.WriteLine("Data {0} {1}", row[0], row[1]);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
