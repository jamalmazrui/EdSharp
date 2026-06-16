using System;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_06
    {
        public static void ParameterizedCommandExample(SqlConnection con,
            string employeeID, string title)
        {
            // Create and configure a new command containing 2 named parameters.
            using (SqlCommand com = con.CreateCommand())
            {
                com.CommandType = CommandType.Text;
                com.CommandText = "UPDATE Employees SET Title = @title" +
                    " WHERE EmployeeId = @id";

                // Create a SqlParameter object for the title parameter.
                SqlParameter p1 = com.CreateParameter();
                p1.ParameterName = "@title";
                p1.SqlDbType = SqlDbType.VarChar;
                p1.Value = title;
                com.Parameters.Add(p1);

                // Use a shorthand syntax to add the id parameter.
                com.Parameters.Add("@id", SqlDbType.Int).Value = employeeID;

                // Execute the command and process the result.
                int result = com.ExecuteNonQuery();

                if (result == 1)
                {
                    Console.WriteLine("Employee {0} title updated to {1}.",
                        employeeID, title);
                }
                else
                {
                    Console.WriteLine("Employee {0} title not updated.",
                        employeeID);
                }
            }
        }

        public static void StoredProcedureExample(SqlConnection con,
            string category, string year)
        {
            // Create and configure a new command.
            using (SqlCommand com = con.CreateCommand())
            {
                com.CommandType = CommandType.StoredProcedure;
                com.CommandText = "SalesByCategory";

                // Create a SqlParameter object for the category parameter.
                com.Parameters.Add("@CategoryName", SqlDbType.NVarChar).Value = 
                    category;

                // Create a SqlParameter object for the year parameter.
                com.Parameters.Add("@OrdYear", SqlDbType.NVarChar).Value = year;

                // Execute the command and process the results.
                using (IDataReader reader = com.ExecuteReader())
                {
                    Console.WriteLine("Sales By Category ({0}).", year);

                    while (reader.Read())
                    {
                        // Display the product details.
                        Console.WriteLine("  {0} = {1}",
                            reader["ProductName"],
                            reader["TotalPurchase"]);
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
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Open the database connection and execute the example
                // commands through the connection.
                con.Open();

                ParameterizedCommandExample(con, "5", "Cleaner");
                Console.WriteLine(Environment.NewLine);

                StoredProcedureExample(con, "Seafood", "1999");
                Console.WriteLine(Environment.NewLine);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
