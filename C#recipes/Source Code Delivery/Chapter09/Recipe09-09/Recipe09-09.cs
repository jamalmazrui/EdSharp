using System;
using System.Data;
using System.Threading;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_09
    {
        // A method to handle asynchronous completion using callbacks.
        public static void CallbackHandler(IAsyncResult result)
        {
            // Obtain a reference to the SqlCommand used to initiate the 
            // asynchronous operation.
            using (SqlCommand cmd = result.AsyncState as SqlCommand)
            {
                // Obtain the result of the stored procedure.
                using (SqlDataReader reader = cmd.EndExecuteReader(result))
                {
                    // Display the results of the stored procedure to the console.
                    lock (Console.Out)
                    {
                        Console.WriteLine(
                            "Price of the Ten Most Expensive Products:");

                        while (reader.Read())
                        {
                            // Display the product details.
                            Console.WriteLine("  {0} = {1}",
                                reader["TenMostExpensiveProducts"],
                                reader["UnitPrice"]);
                        }
                    }
                }
            }
        }

        public static void Main()
        {
            // Create a new SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                // You must specify Asynchronous Processing=true to support
                // asynchronous operations over the connection.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI;" +
                    "Asynchronous Processing=true";

                // Create and configure a new command to run a stored procedure.
                // Do not wrap it in a using statement because the asynchronous 
                // completion handler will dispose of the SqlCommand object.
                SqlCommand cmd = con.CreateCommand();
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.CommandText = "Ten Most Expensive Products";

                // Open the database connection and execute the command
                // asynchronously. Pass the reference to the SqlCommand
                // used to initiate the asynchronous operation.
                con.Open();
                cmd.BeginExecuteReader(CallbackHandler, cmd);

                // Continue with other processing.
                for (int count = 0; count < 10; count++)
                {
                    lock (Console.Out)
                    {
                        Console.WriteLine("{0} : Continue processing...",
                            DateTime.Now.ToString("HH:mm:ss.ffff"));
                    }
                    Thread.Sleep(500);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}