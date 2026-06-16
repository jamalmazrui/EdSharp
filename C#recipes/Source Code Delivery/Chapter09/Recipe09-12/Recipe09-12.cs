using System;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_12
    {
        static void Main(string[] args)
        {

            // Create a new SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Open the database connection
                con.Open();

                // create the query string
                string query = "SELECT * from Region";

                // create the data set
                DataSet dataset = new DataSet();
                
                // create the sql data adapter
                SqlDataAdapter adapter = new SqlDataAdapter(query, con);
                // create the command builder so we can do modifications
                SqlCommandBuilder commbuilder = new SqlCommandBuilder(adapter);

                // populate the data set from the database
                adapter.Fill(dataset);                

                // print details of the schema
                Console.WriteLine("\nSchema for table");
                DataTable table = dataset.Tables[0];
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

                // write the data back to the database
                adapter.Update(dataset);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
