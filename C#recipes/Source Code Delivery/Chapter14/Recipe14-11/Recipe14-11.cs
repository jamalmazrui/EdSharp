using System;
using Microsoft.WindowsAPICodePack.Net;

namespace Recipe14_11
{
    class Recipe14_11
    {
        static void Main(string[] args)
        {
            // check the internet connection state
            bool isInternetConnected = 
                NetworkListManager.IsConnectedToInternet;

            Console.WriteLine("Machine connected to Internet: {0}", 
                isInternetConnected);
            
            if (isInternetConnected)
            {
                // get the list of all network connections
                NetworkCollection netCollection = 
                    NetworkListManager.GetNetworks(NetworkConnectivityLevels.Connected);

                // work through the set of connections and write out the 
                // name of those which are connected to the internet
                foreach (Network network in netCollection)
                {
                    if (network.IsConnectedToInternet)
                    {
                        Console.WriteLine("Connection {0} is connected to the internet",
                            network.Name);
                    }
                }
            }

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
