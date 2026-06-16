using System;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_02
    {
        public static void Main()
        {
            // Obtain a pooled connection.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString =
                    @"Data Source = .\sqlexpress;" +// local SQL Server instance
                    "Database = Northwind;" +       // the sample Northwind DB
                    "Integrated Security = SSPI;" + // integrated Windows security
                    "Min Pool Size = 5;" +          // configure minimum pool size
                    "Max Pool Size = 15;" +         // configure maximum pool size
                    "Connection Reset = True;" +    // reset connections each use
                    "Connection Lifetime = 600";    // set max connection lifetime

                // Open the database connection.
                con.Open();

                // Access the database...

                // At the end of the using block, the Dispose calls Close, which
                // returns the connection to the pool for reuse.
            }

            // Obtain a nonpooled connection.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString =
                    @"Data Source = .\sqlexpress;" +//local SQL Server instance
                    "Database = Northwind;" +       //the sample Northwind DB
                    "Integrated Security = SSPI;" + //integrated Windows security
                    "Pooling = False";              //specify nonpooled connection

                // Open the database connection.
                con.Open();

                // Access the database...

                // At the end of the using block, the Dispose calls Close, which
                // closes the nonpooled connection.
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
