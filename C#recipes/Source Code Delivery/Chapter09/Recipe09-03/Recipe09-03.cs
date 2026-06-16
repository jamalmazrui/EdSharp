using System;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_03
    {
        public static void Main(string[] args)
        {
            string conString = @"Data Source=.\sqlexpress;" +
                "Database=Northwind;Integrated Security=SSPI;" +
                "Min Pool Size=5;Max Pool Size=15;Connection Reset=True;" +
                "Connection Lifetime=600;";

            // Parse the SQL Server connection string and display the component
            // configuration parameters.
            SqlConnectionStringBuilder sb1 = 
                new SqlConnectionStringBuilder(conString);

            Console.WriteLine("Parsed SQL Connection String Parameters:");
            Console.WriteLine("  Database Source = " + sb1.DataSource);
            Console.WriteLine("  Database = " + sb1.InitialCatalog);
            Console.WriteLine("  Use Integrated Security = " 
                + sb1.IntegratedSecurity);
            Console.WriteLine("  Min Pool Size = " + sb1.MinPoolSize);
            Console.WriteLine("  Max Pool Size = " + sb1.MaxPoolSize);
            Console.WriteLine("  Lifetime = " + sb1.LoadBalanceTimeout);

            // Build a connection string from component parameters and display it.
            SqlConnectionStringBuilder sb2 =
                new SqlConnectionStringBuilder(conString);

            sb2.DataSource = @".\sqlexpress";
            sb2.InitialCatalog = "Northwind";
            sb2.IntegratedSecurity = true;
            sb2.MinPoolSize = 5;
            sb2.MinPoolSize = 15;
            sb2.LoadBalanceTimeout = 600;

            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Constructed connection string:");
            Console.WriteLine("  " + sb2.ConnectionString);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}