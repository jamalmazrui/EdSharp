using System;
using System.Net;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_08
    {
        public static void Main(string[] args) 
        {
            foreach (string comp in args) 
            {
                try
                {
                    // Retrieve the DNS entry for the specified computer.
                    IPAddress[] addresses = Dns.GetHostEntry(comp).AddressList;

                    // The DNS entry may contain more than one IP address. Iterate
                    // through them and display each one along with the type of
                    // address (AddressFamily).
                    foreach (IPAddress address in addresses)
                    {
                        Console.WriteLine("{0} = {1} ({2})", 
                            comp, address, address.AddressFamily);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("{0} = Error ({1})", comp, ex.Message);
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
