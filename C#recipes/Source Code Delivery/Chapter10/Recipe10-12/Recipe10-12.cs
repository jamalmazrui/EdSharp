using System;
using System.Text;
using System.Net;
using System.Net.Sockets;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_12
    {
        private static int localPort;

        private static void Main() 
        {
            // Define endpoint where messages are sent.
            Console.Write("Connect to IP: ");
            string IP = Console.ReadLine();
            Console.Write("Connect to port: ");
            int port = Int32.Parse(Console.ReadLine());

            IPEndPoint remoteEndPoint = 
                new IPEndPoint(IPAddress.Parse(IP), port);

            // Define local endpoint (where messages are received).
            Console.Write("Local port for listening: ");
            localPort = Int32.Parse(Console.ReadLine());

            // Create a new thread for receiving incoming messages.
            Thread receiveThread = new Thread(ReceiveData);
            receiveThread.IsBackground = true;
            receiveThread.Start();

            UdpClient client = new UdpClient();

            Console.WriteLine("Type message and press Enter to send:");

            try
            {
                string text;

                do
                {
                    text = Console.ReadLine();

                    // Send the text to the remote client.
                    if (text.Length != 0)
                    {
                        // Encode the data to binary using UTF8 encoding.
                        byte[] data = Encoding.UTF8.GetBytes(text);

                        // Send the text to the remote client.
                        client.Send(data, data.Length, remoteEndPoint);
                    }
                } while (text.Length != 0);
            }
            catch (Exception err)
            {
                Console.WriteLine(err.ToString());
            }
            finally
            {
                client.Close();
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }

        private static void ReceiveData() 
        {
            UdpClient client = new UdpClient(localPort);

            while (true) 
            {
                try 
                {
                    // Receive bytes.
                    IPEndPoint anyIP = new IPEndPoint(IPAddress.Any, 0);
                    byte[] data = client.Receive(ref anyIP);

                    // Convert bytes to text using UTF8 encoding.
                    string text = Encoding.UTF8.GetString(data);

                    // Display the retrieved text.
                    Console.WriteLine(">> " + text);
                } 
                catch (Exception err) 
                {
                    Console.WriteLine(err.ToString());
                }
            }
        }
    }
}

