using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_01
    {
        public static void SqlConnectionExample()
        {
            // Create an empty SqlConnection object.
            using (SqlConnection con = new SqlConnection())
            {
                // Configure the SqlConnection object's connection string.
                con.ConnectionString =
                    @"Data Source=.\sqlexpress;" + // local SQL Server instance
                    "Database=Northwind;" +        // the sample Northwind DB
                    "Integrated Security=SSPI";    // integrated Windows security

                // Open the database connection.
                con.Open();

                // Display information about the connection.
                if (con.State == ConnectionState.Open)
                {
                    Console.WriteLine("SqlConnection Information:");
                    Console.WriteLine("  Connection State = " + con.State);
                    Console.WriteLine("  Connection String = " +
                        con.ConnectionString);
                    Console.WriteLine("  Database Source = " + con.DataSource);
                    Console.WriteLine("  Database = " + con.Database);
                    Console.WriteLine("  Server Version = " + con.ServerVersion);
                    Console.WriteLine("  Workstation Id = " + con.WorkstationId);
                    Console.WriteLine("  Timeout = " + con.ConnectionTimeout);
                    Console.WriteLine("  Packet Size = " + con.PacketSize);
                }
                else
                {
                    Console.WriteLine("SqlConnection failed to open.");
                    Console.WriteLine("  Connection State = " + con.State);
                }
                // At the end of the using block Dispose() calls Close().
            }
        }

        public static void OleDbConnectionExample()
        {

            // Create an empty OleDbConnection object.
            using (OleDbConnection con = new OleDbConnection())
            {
                // Configure the OleDbConnection object's connection string.
                con.ConnectionString =
                    "Provider=SQLOLEDB;" +         // OLE DB Provider for SQL Server
                    @"Data Source=.\sqlexpress;" + // local SQL Server instance
                    "Initial Catalog=Northwind;" + // the sample Northwind DB
                    "Integrated Security=SSPI";      // integrated Windows security

                // Open the database connection.
                con.Open();

                // Display information about the connection.
                if (con.State == ConnectionState.Open)
                {
                    Console.WriteLine("OleDbConnection Information:");
                    Console.WriteLine("  Connection State = " + con.State);
                    Console.WriteLine("  Connection String = " +
                        con.ConnectionString);
                    Console.WriteLine("  Database Source = " + con.DataSource);
                    Console.WriteLine("  Database = " + con.Database);
                    Console.WriteLine("  Server Version = " + con.ServerVersion);
                    Console.WriteLine("  Timeout = " + con.ConnectionTimeout);
                }
                else
                {
                    Console.WriteLine("OleDbConnection failed to open.");
                    Console.WriteLine("  Connection State = " + con.State);
                }
                // At the end of the using block Dispose() calls Close().
            }
        }

        public static void Main()
        {
            // Open connection using SqlConnection.
            SqlConnectionExample();
            Console.WriteLine(Environment.NewLine);

            // Open connection using OleDbConnection.
            OleDbConnectionExample();

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
