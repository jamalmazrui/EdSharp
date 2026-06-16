using System;
using System.Net;
using System.Net.Sockets;
using System.IO;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public class Recipe10_11Client
    {
        private static void Main()
        {
            using (TcpClient client = new TcpClient())
            {
                Console.WriteLine("Attempting to connect to the server ",
                    "on port 8000.");

                // Connect to the server.
                client.Connect(IPAddress.Parse("127.0.0.1"), 8000);

                // Retrieve the network stream from the TcpClient.
                using (NetworkStream networkStream = client.GetStream())
                {
                    // Create a BinaryWriter for writing to the stream.
                    using (BinaryWriter writer = new BinaryWriter(networkStream))
                    {
                        // Start a dialogue.
                        writer.Write(Recipe10_11Shared.RequestConnect);

                        // Create a BinaryReader for reading from the stream.
                        using (BinaryReader reader =
                            new BinaryReader(networkStream))
                        {
                            if (reader.ReadString() ==
                                Recipe10_11Shared.AcknowledgeOK)
                            {
                                Console.WriteLine("Connection established." +
                                    "Press Enter to download data.");

                                Console.ReadLine();

                                // Send message requesting data to server.
                                writer.Write(Recipe10_11Shared.RequestData);

                                // The server should respond with the size of
                                // the data it will send. Assume it does.
                                int fileSize = int.Parse(reader.ReadString());

                                // Only get part of the data then carry out a 
                                // premature disconnect.
                                for (int i = 0; i < fileSize / 3; i++)
                                {
                                    Console.Write(networkStream.ReadByte());
                                }

                                Console.WriteLine(Environment.NewLine);
                                Console.WriteLine("Press Enter to disconnect.");
                                Console.ReadLine();
                                Console.WriteLine("Disconnecting...");
                                writer.Write(Recipe10_11Shared.Disconnect);
                            }
                            else
                            {
                                Console.WriteLine("Connection not established.");
                            }
                        }
                    }
                }
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Connection closed. Press Enter");
            Console.ReadLine();
        }
    }
}