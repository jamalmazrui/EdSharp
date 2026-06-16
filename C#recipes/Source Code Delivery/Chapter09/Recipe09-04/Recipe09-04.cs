using System;
using System.Configuration;
using System.Data.SqlClient;

namespace Apress.VisualCSharpRecipes.Chapter09
{
    class Recipe09_04
    {
        private static void WriteEncryptedConnectionStringSection(
            string name, string constring, string provider)
        {
            // Get the configuration file for the current application. Specify
            // the ConfigurationUserLevel.None argument so that we get the
            // configuration settings that apply to all users.
            Configuration config = ConfigurationManager.OpenExeConfiguration(
                ConfigurationUserLevel.None);

            // Get the connectionStrings section from the configuration file.
            ConnectionStringsSection section = config.ConnectionStrings;

            // If the connectionString section does not exist, create it.
            if (section == null) 
            {
                Console.WriteLine("Section needs creating");
                section = new ConnectionStringsSection();
                config.Sections.Add("connectionSettings", section);
            }

            // If it is not already encrypted, configure the connectionStrings 
            // section to be encrypted using the standard RSA Proected 
            // Configuration Provider.
            if (!section.SectionInformation.IsProtected)
            {
                // Remove this statement to write the connection string in clear 
                // text for the purpose of testing.
                section.SectionInformation.ProtectSection(
                    "RsaProtectedConfigurationProvider");
            }

            // Create a new connection string element and add it to the 
            // connection string configuration section.
            ConnectionStringSettings cs = 
                new ConnectionStringSettings(name, constring, provider);
            section.ConnectionStrings.Add(cs);
 
            // Force the connection string section to be saved.
            section.SectionInformation.ForceSave = true;

            // Save the updated configuration file.
            config.Save(ConfigurationSaveMode.Full);
        }

        public static void Main(string[] args)
        {
            // The connection string information to be written to the 
            // configuration file.
            string conName = "ConnectionString1";
            string conString = @"Data Source=.\sqlexpress;" +
                "Database=Northwind;Integrated Security=SSPI;" +
                "Min Pool Size=5;Max Pool Size=15;Connection Reset=True;" +
                "Connection Lifetime=600;";
            string providerName = "System.Data.SqlClient";

            // Write the new connection string to the application's 
            // configuration file.
            WriteEncryptedConnectionStringSection(conName, conString, providerName);

            // Read the encrypted connection string settings from the 
            // application's configuration file.
            ConnectionStringSettings cs2 =
                ConfigurationManager.ConnectionStrings["ConnectionString1"];
            
            // Use the connection string to create a new SQL Server connection.
            using (SqlConnection con = new SqlConnection(cs2.ConnectionString))
            {
                // Issue database commands/queries...

            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}

