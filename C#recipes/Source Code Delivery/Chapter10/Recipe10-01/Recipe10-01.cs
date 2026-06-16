using System;
using System.Net.NetworkInformation;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_01
    {
        static void Main()
        {
            // Only proceed if there is a network available.
            if (NetworkInterface.GetIsNetworkAvailable())
            {
                // Get the set of all NetworkInterface objects for the local 
                // machine.
                NetworkInterface[] interfaces =
                    NetworkInterface.GetAllNetworkInterfaces();

                // Iterate through the interfaces and display information.
                foreach (NetworkInterface ni in interfaces)
                {
                    // Report basic interface information.
                    Console.WriteLine("Interface Name: {0}", ni.Name);
                    Console.WriteLine("    Description: {0}", ni.Description);
                    Console.WriteLine("    ID: {0}", ni.Id);
                    Console.WriteLine("    Type: {0}", ni.NetworkInterfaceType);
                    Console.WriteLine("    Speed: {0}", ni.Speed);
                    Console.WriteLine("    Status: {0}", ni.OperationalStatus);

                    // Report physical address.
                    Console.WriteLine("    Physical Address: {0}",
                        ni.GetPhysicalAddress().ToString());

                    // Report network statistics for the interface.
                    Console.WriteLine("    Bytes Sent: {0}",
                        ni.GetIPv4Statistics().BytesSent);
                    Console.WriteLine("    Bytes Received: {0}",
                        ni.GetIPv4Statistics().BytesReceived);

                    // Report IP configuration.
                    Console.WriteLine("    IP Addresses:");
                    foreach (UnicastIPAddressInformation addr
                        in ni.GetIPProperties().UnicastAddresses)
                    {
                        Console.WriteLine("        - {0} (lease expires {1})", 
                            addr.Address, DateTime.Now + 
                            new TimeSpan(0, 0, (int)addr.DhcpLeaseLifetime));
                    }

                    Console.WriteLine(Environment.NewLine);
                }
            }
            else
            {
                Console.WriteLine("No network available.");
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
