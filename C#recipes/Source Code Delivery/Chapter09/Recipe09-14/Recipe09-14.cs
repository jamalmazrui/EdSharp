using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_14
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

                // perform the LINQ query
                IEnumerable<string> result = from e in table.AsEnumerable()
                                             where e.Field<int>(0) < 3
                                             select e.Field<string>(1);

                // enumerate the results of the linq query
                foreach (string str in result)
                {
                    Console.WriteLine("Result: {0}", str);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
