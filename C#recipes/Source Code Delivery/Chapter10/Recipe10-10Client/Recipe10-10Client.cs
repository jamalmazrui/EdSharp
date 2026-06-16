using System;
using System.IO;
using System.Net;
using System.Net.Sockets;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public class Recipe10_10Client
    {
        public static void Main() 
        {
            TcpClient client = new TcpClient();

            try
            {
                Console.WriteLine("Attempting to connect to the server ",
                    "on port 8000.");
                client.Connect(IPAddress.Parse("127.0.0.1"), 8000);
                Console.WriteLine("Connection established.");

                // Retrieve the network stream.
                NetworkStream stream = client.GetStream();

                // Create a BinaryWriter for writing to the stream.
                using (BinaryWriter w = new BinaryWriter(stream))
                {
                    // Create a BinaryReader for reading from the stream.
                    using (BinaryReader r = new BinaryReader(stream))
                    {
                        // Start a dialogue.
                        w.Write(Recipe10_10Shared.RequestConnect);

                        if (r.ReadString() == Recipe10_10Shared.AcknowledgeOK)
                        {
                            Console.WriteLine("Connected.");
                            Console.WriteLine("Press Enter to disconnect.");
                            Console.ReadLine();
                            Console.WriteLine("Disconnecting...");
                            w.Write(Recipe10_10Shared.Disconnect);
                        }
                        else
                        {
                            Console.WriteLine("Connection not completed.");
                        }
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err.ToString());
            }
            finally
            {
                // Close the connection socket.
                client.Close();
                Console.WriteLine("Port closed.");
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}