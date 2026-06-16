using System;
using System.Net.NetworkInformation;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_02
    {
        // Declare a method to handle NetworkAvailabilityChanged events.
        private static void NetworkAvailabilityChanged(
            object sender, NetworkAvailabilityEventArgs e)
        {
            // Report whether the network is now available or unavailable.
            if (e.IsAvailable)
            {
                Console.WriteLine("Network Available");
            }
            else
            {
                Console.WriteLine("Network Unavailable");
            }
        }

        // Declare a method to handle NetworkAdressChanged events.
        private static void NetworkAddressChanged(object sender, EventArgs e)
        {
            Console.WriteLine("Current IP Addresses:");

            // Iterate through the interfaces and display information.
            foreach (NetworkInterface ni in 
                NetworkInterface.GetAllNetworkInterfaces())
            {
                foreach (UnicastIPAddressInformation addr
                    in ni.GetIPProperties().UnicastAddresses)
                {
                    Console.WriteLine("    - {0} (lease expires {1})", 
                        addr.Address, DateTime.Now + 
                        new TimeSpan(0, 0, (int)addr.DhcpLeaseLifetime));
                }
            }
        }

        static void Main(string[] args)
        {
            // Add the handlers to the NetworkChange events.
            NetworkChange.NetworkAvailabilityChanged += 
                NetworkAvailabilityChanged;
            NetworkChange.NetworkAddressChanged += 
                NetworkAddressChanged;

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Press Enter to stop waiting for network events");
            Console.ReadLine();
        }
    }
}
