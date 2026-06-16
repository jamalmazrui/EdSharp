using System;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_05
    {
        public static void ExecuteNonQueryExample(IDbConnection con)
        {
            // Create and configure a new command.
            IDbCommand com = con.CreateCommand();
            com.CommandType = CommandType.Text;
            com.CommandText = "UPDATE Employees SET Title = 'Sales Director'" +
                " WHERE EmployeeId = '5'";

            // Execute the command and process the result.
            int result = com.ExecuteNonQuery();

            if (result == 1)
            {
                Console.WriteLine("Employee title updated.");
            }
            else
            {
                Console.WriteLine("Employee title not updated.");
            }
        }

        public static void ExecuteReaderExample(IDbConnection con)
        {
            // Create and configure a new command.
            IDbCommand com = con.CreateCommand();
            com.CommandType = CommandType.StoredProcedure;
            com.CommandText = "Ten Most Expensive Products";

            // Execute the command and process the results
            using (IDataReader reader = com.ExecuteReader())
            {
                Console.WriteLine("Price of the Ten Most Expensive Products.");

                while (reader.Read())
                {
                    // Display the product details.
                    Console.WriteLine("  {0} = {1}",
                        reader["TenMostExpensiveProducts"],
                        reader["UnitPrice"]);
                }
            }
        }

        public static void ExecuteScalarExample(IDbConnection con)
        {
            // Create and configure a new command.
            IDbCommand com = con.CreateCommand();
            com.CommandType = CommandType.Text;
            com.CommandText = "SELECT COUNT(*) FROM Employees";

            // Execute the command and cast the result.
            int result = (int)com.ExecuteScalar();

            Console.WriteLine("Employee count = " + result);
        }

        public static void Main()
        {
            // Create a new SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Open the database connection and execute the example
                // commands through the connection.
                con.Open();

                ExecuteNonQueryExample(con);
                Console.WriteLine(Environment.NewLine);

                ExecuteReaderExample(con);
                Console.WriteLine(Environment.NewLine);

                ExecuteScalarExample(con);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
