using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_16
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

                // obtain the data table
                DataTable table = dataset.Tables[0];

                // perform the first LINQ query
                IEnumerable<DataRow> result1 = from e in table.AsEnumerable()
                                              where e.Field<int>(0) < 3
                                              select e;

                // enumerate the results of the first linq query
                Console.WriteLine("Results from first LINQ query");
                foreach (DataRow row in result1)
                {
                    Console.WriteLine("ID: {0} Name: {1}", 
                        row.Field<int>(0), row.Field<string>(1));
                }

                // perform the first LINQ query
                IEnumerable<DataRow> result2 = from e in table.AsEnumerable()
                                              let name = e.Field<string>(1)
                                              where name.StartsWith("North")
                                                || name.StartsWith("East")
                                              select e;

                // enumerate the results of the first linq query
                Console.WriteLine("\nResults from second LINQ query");
                foreach (DataRow row in result2)
                {
                    Console.WriteLine("ID: {0} Name: {1}",
                        row.Field<int>(0), row.Field<string>(1));
                }

                IEnumerable<DataRow> union = result1.Union(result2);
                // enumerate the results
                Console.WriteLine("\nResults from union");
                foreach (DataRow row in union)
                {
                    Console.WriteLine("ID: {0} Name: {1}",
                        row.Field<int>(0), row.Field<string>(1));
                }


                IEnumerable<DataRow> intersect = result1.Intersect(result2);
                // enumerate the results
                Console.WriteLine("\nResults from intersect");
                foreach (DataRow row in intersect)
                {
                    Console.WriteLine("ID: {0} Name: {1}",
                        row.Field<int>(0), row.Field<string>(1));
                }

                IEnumerable<DataRow> except = result1.Except(result2);
                // enumerate the results
                Console.WriteLine("\nResults from except");
                foreach (DataRow row in except)
                {
                    Console.WriteLine("ID: {0} Name: {1}",
                        row.Field<int>(0), row.Field<string>(1));
                }

            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
