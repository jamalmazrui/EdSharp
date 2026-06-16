using System;
using System.Data;
using System.Data.Common;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_10
    {
        public static void Main(string[] args)
        {
            // Obtain the list of ADO.NET data providers registered in the 
            // machine and application configuration file.
            using (DataTable providers = DbProviderFactories.GetFactoryClasses())
            {
                // Enumerate the set of data providers and display details.
                Console.WriteLine("Available ADO.NET Data Providers:");
                foreach (DataRow prov in providers.Rows)
                {
                    Console.WriteLine(" Name:{0}", prov["Name"]);
                    Console.WriteLine("   Description:{0}", 
                        prov["Description"]);
                    Console.WriteLine("   Invariant Name:{0}", 
                        prov["InvariantName"]);
                }
            }

            // Obtain the DbProviderFactory for SQL Server, the provider to use 
            // could be selected by the user or read from a configuration file. 
            // In this case, we simply pass the invariant name.
            DbProviderFactory factory =
                DbProviderFactories.GetFactory("System.Data.SqlClient");

            // Use the DbProviderFactory to create the initial IDbConnection, and
            // then the data provider interface factory methods for other objects.
            using (IDbConnection con = factory.CreateConnection())
            {
                // Normally, read the connection string from secure storage. 
                // See recipe 9-3. In this case, use a default value.
                con.ConnectionString = @"Data Source = .\sqlexpress;" +
                    "Database = Northwind; Integrated Security=SSPI";

                // Create and configure a new command.
                using (IDbCommand com = con.CreateCommand())
                {
                    com.CommandType = CommandType.StoredProcedure;
                    com.CommandText = "Ten Most Expensive Products";

                    // Open the connection.
                    con.Open();

                    // Execute the command and process the results.
                    using (IDataReader reader = com.ExecuteReader())
                    {
                        Console.WriteLine(Environment.NewLine);
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
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
