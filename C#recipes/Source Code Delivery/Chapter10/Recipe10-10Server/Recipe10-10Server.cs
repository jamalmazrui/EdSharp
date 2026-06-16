using System;
using System.IO;
using System.Net;
using System.Net.Sockets;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public class Recipe10_10Server
    {
        public static void Main() 
        {
            // Create a new listener on port 8000.
            TcpListener listener = 
                new TcpListener(IPAddress.Parse("127.0.0.1"), 8000);
                
            Console.WriteLine("About to initialize port.");
            listener.Start();
            Console.WriteLine("Listening for a connection...");
                
            try 
            {
                // Wait for a connection request, and return a TcpClient 
                // initialized for communication. 
                using (TcpClient client = listener.AcceptTcpClient())
                {
                    Console.WriteLine("Connection accepted.");
                        
                    // Retrieve the network stream.
                    NetworkStream stream = client.GetStream();

                    // Create a BinaryWriter for writing to the stream.
                    using (BinaryWriter w = new BinaryWriter(stream))
                    {
                        // Create a BinaryReader for reading from the stream.
                        using (BinaryReader r = new BinaryReader(stream))
                        {
                            if (r.ReadString() == 
                                Recipe10_10Shared.RequestConnect)
                            {
                                w.Write(Recipe10_10Shared.AcknowledgeOK);
                                Console.WriteLine("Connection completed.");

                                while (r.ReadString() != 
                                    Recipe10_10Shared.Disconnect) { }

                                Console.WriteLine(Environment.NewLine);
                                Console.WriteLine("Disconnect request received.");
                            }
                            else
                            {
                                Console.WriteLine("Can't complete connection.");
                            }
                        }
                    }
                }

                Console.WriteLine("Connection closed.");
            } 
            catch (Exception ex) 
            {
                Console.WriteLine(ex.ToString());
            }
            finally
            {
                // Close the underlying socket (stop listening for new requests).
                listener.Stop();
                Console.WriteLine("Listener stopped.");
            }
            
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
       }
    }
}