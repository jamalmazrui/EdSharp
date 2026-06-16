using System;
using System.Data;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_07
    {
        public static void Main()
        {
            // Create a new SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Create and configure a new command.
                using (SqlCommand com = con.CreateCommand())
                {
                    com.CommandType = CommandType.Text;
                    com.CommandText = "SELECT BirthDate,FirstName,LastName FROM "+
                        "Employees ORDER BY BirthDate;SELECT * FROM Employees";

                    // Open the database connection and execute the example
                    // commands through the connection.
                    con.Open();

                    // Execute the command and obtain a SqlReader.
                    using (SqlDataReader reader = com.ExecuteReader())
                    {
                        // Process the first set of results and display the 
                        // content of the result set.
                        Console.WriteLine("Employee Birthdays (By Age).");

                        while (reader.Read())
                        {
                            Console.WriteLine("  {0,18:D} - {1} {2}",
                                reader.GetDateTime(0),  // Retrieve typed data
                                reader["FirstName"],    // Use string index
                                reader[2]);             // Use ordinal index
                        }
                        Console.WriteLine(Environment.NewLine);

                        // Process the second set of results and display details
                        // about the columns and data types in the result set.
                        reader.NextResult();
                        Console.WriteLine("Employee Table Metadata.");
                        for (int field = 0; field < reader.FieldCount; field++)
                        {
                            Console.WriteLine("  Column Name:{0}  Type:{1}",
                                reader.GetName(field),
                                reader.GetDataTypeName(field));
                        }
                    }
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
