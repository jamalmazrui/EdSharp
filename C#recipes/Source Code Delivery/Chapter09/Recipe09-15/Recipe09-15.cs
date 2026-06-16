using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Data.Linq;
using System.Data.Linq.Mapping;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    [Table(Name = "Region")]
    public partial class Region
    {
        [Column]
        public int RegionID;
        [Column]
        public string RegionDescription;
    }
    class Recipe09_15
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

                // create the data context
                DataContext context = new DataContext(con);

                // get the table we are interested in
                Table<Region> regionstable = context.GetTable<Region>();

                IEnumerable<Region> result = from e in regionstable 
                                             where e.RegionID < 3 
                                             select e;
                foreach (Region res in result)
                {
                    Console.WriteLine("RegionID {0} Descr: {1}", res.RegionID, res.RegionDescription);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
