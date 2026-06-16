using System;
using System.Data;
using System.Data.Sql;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_11
    {
        public static void Main(string[] args)
        {
            // Obtain the DataTable of SQL Server instances.
            using (DataTable SqlSources =
                SqlDataSourceEnumerator.Instance.GetDataSources())
            {
                // Enumerate the set of SQL Servers and display details.
                Console.WriteLine("Discover SQL Server Instances:");
                foreach (DataRow source in SqlSources.Rows)
                {
                    Console.WriteLine(" Server Name:{0}", source["ServerName"]);
                    Console.WriteLine("   Instance Name:{0}", 
                        source["InstanceName"]);
                    Console.WriteLine("   Is Clustered:{0}", 
                        source["IsClustered"]);
                    Console.WriteLine("   Version:{0}", source["Version"]);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
