using System;
using System.Net.NetworkInformation;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_09
    {
        public static void Main(string[] args)
        {
            // Create an instance of the Ping class.
            using (Ping ping = new Ping())
            {
                Console.WriteLine("Pinging:");

                foreach (string comp in args)
                {
                    try
                    {
                        Console.Write("    {0}...", comp);

                        // Ping the specified computer with a tim-eout of 100ms.
                        PingReply reply = ping.Send(comp, 100);

                        if (reply.Status == IPStatus.Success)
                        {
                          Console.WriteLine("Success - IP Address:{0} Time:{1}ms",
                              reply.Address, reply.RoundtripTime);
                        }
                        else
                        {
                            Console.WriteLine(reply.Status);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error ({0})", 
                            ex.InnerException.Message);
                    }
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
