using System;
using System.IO;
using System.Net;
using System.Threading;
using System.Net.Sockets;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    public class Recipe10_11Server
    {
        // A flag used to indicate whether the server is shutting down.
        private static bool terminate;
        public static bool Terminate { get { return terminate; } }

        // A variable to track the identity of each client connection.
        private static int ClientNumber = 0;

        // A single TcpListener will accept all incoming client connections.
        private static TcpListener listener;

        public static void Main()
        {
            // Create a 100Kb test file for use in the example. This file will be
            // sent to clients that connect.
            using (FileStream fs = new FileStream("test.bin", FileMode.Create))
            {
                fs.SetLength(100000);
            }

            try
            {
                // Create a TcpListener that will accept incoming client 
                // connections on port 8000 of the local machine.
                listener = new TcpListener(IPAddress.Parse("127.0.0.1"), 8000);

                Console.WriteLine("Starting TcpListener...");

                // Start the TcpListener accepting connections.
                terminate = false;
                listener.Start();

                // Begin asynchronously listening for client connections. When a 
                // new connection is established, call the ConnectionHandler 
                // method to process the new connection. 
                listener.BeginAcceptTcpClient(ConnectionHandler, null);

                // Keep the server active until the user presses Enter.
                Console.WriteLine("Server awaiting connections. " +
                    "Press Enter to stop server.");
                Console.ReadLine();
            }
            finally
            {
                // Shut down the TcpListener. This will cause any outstanding
                // asynchronous requests to stop and throw an exception in 
                // the ConnectionHandler when EndAcceptTcpClient is called.
                // More robust termination synchronization may be desired here, 
                // but for the purpose of this example ClientHandler threads are 
                // all background threads and will terminate automatically when 
                // the main thread terminates. This is suitable for our needs.
                Console.WriteLine("Server stopping...");
                terminate = true;
                if (listener != null) listener.Stop();
            }

            // Wait to continue. 
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Server stopped. Press Enter");
            Console.ReadLine();
        }

        // A method to handle the callback when a connection is established
        // from a client. This is a simple way to implement a dispatcher
        // but lacks the control and scalability required when implementing 
        // full-blown asynchronous server applications.
        private static void ConnectionHandler(IAsyncResult result)
        {
            TcpClient client = null;

            // Always end the asynchronous operation to avoid leaks.
            try
            {
                // Get the TcpClient that represents the new client connection.
                client = listener.EndAcceptTcpClient(result);
            }
            catch (ObjectDisposedException)
            {
                // Server is shutting down and the outstanding asynchronous 
                // request calls the completion method with this exception. 
                // The exception is thrown when EndAcceptTcpClient is called.
                // Do nothing and return.
                return;
            }

            Console.WriteLine("Dispatcher: New connection accepted.");

            // Begin asynchronously listening for the next client 
            // connection. 
            listener.BeginAcceptTcpClient(ConnectionHandler, null);

            if (client != null)
            {
                // Determine the identifier for the new client connection.
                Interlocked.Increment(ref ClientNumber);
                string clientName = "Client " + ClientNumber.ToString();

                Console.WriteLine("Dispatcher: Creating client handler ({0})."
                    , clientName);

                // Create a new ClientHandler to handle this connection.
                new ClientHandler(client, clientName);
            }
        }
    }

    // A class that encapsulates the logic to handle a client connection.
    public class ClientHandler
    {
        // The TcpClient that represents the connection to the client.
        private TcpClient client;

        // An ID that uniquely identifies this ClientHandler.
        private string ID;

        // The amount of data that will be written in one block (2 KB).
        private int bufferSize = 2048;

        // The buffer that holds the data to write.
        private byte[] buffer;

        // Used to read data from the local file.
        private FileStream fileStream;

        // A signal to stop sending data to the client.
        private bool stopDataTransfer;

        internal ClientHandler(TcpClient client, string ID)
        {
            this.buffer = new byte[bufferSize];
            this.client = client;
            this.ID = ID;

            // Create a new background thread to handle the client connection 
            // so that we do not consume a thread-pool thread for a long time
            // and also so that it will be terminated when the main thread ends.
            Thread thread = new Thread(ProcessConnection);
            thread.IsBackground = true;
            thread.Start();
        }

        private void ProcessConnection()
        {
            using (client)
            {
                // Create a BinaryReader to receive messages from the client. At
                // the end of the using block, it will close both the BinaryReader
                // and the underlying NetworkStream.
                using (BinaryReader reader = new BinaryReader(client.GetStream()))
                {
                    if (reader.ReadString() == Recipe10_11Shared.RequestConnect)
                    {
                        // Create a BinaryWriter to send messages to the client. 
                        // At the end of the using block, it will close both the 
                        // BinaryWriter and the underlying NetworkStream.
                        using (BinaryWriter writer =
                            new BinaryWriter(client.GetStream()))
                        {
                            writer.Write(Recipe10_11Shared.AcknowledgeOK);
                            Console.WriteLine(ID + ": Connection established.");

                            string message = "";

                            while (message != Recipe10_11Shared.Disconnect)
                            {
                                try
                                {
                                    // Read the message from the client.
                                    message = reader.ReadString();
                                }
                                catch
                                {
                                    // For the purpose of the example, any
                                    // exception should be taken as a 
                                    // client disconnect.
                                    message = Recipe10_11Shared.Disconnect;
                                }

                                if (message == Recipe10_11Shared.RequestData)
                                {
                                    Console.WriteLine(ID + ": Requested data. ",
                                        "Sending...");

                                    // The filename could be supplied by the 
                                    // client, but in this example a test file 
                                    // is hard-coded.
                                    fileStream = new FileStream("test.bin",
                                        FileMode.Open, FileAccess.Read);

                                    // Send the file size--this is how the client 
                                    // knows how much to read.
                                    writer.Write(fileStream.Length.ToString());

                                    // Start an asynchronous send operation.
                                    stopDataTransfer = false;
                                    StreamData(null);
                                }
                                else if (message == Recipe10_11Shared.Disconnect)
                                {
                                    Console.WriteLine(ID +
                                        ": Client disconnecting...");
                                    stopDataTransfer = true;
                                }
                                else
                                {
                                    Console.WriteLine(ID + ": Unknown command.");
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine(ID +
                            ": Could not establish connection.");
                    }
                }
            }

            Console.WriteLine(ID + ": Client connection closed.");
        }

        private void StreamData(IAsyncResult asyncResult)
        {
            // Always complete outstanding asynchronous operations to avoid leaks.
            if (asyncResult != null)
            {
                try
                {
                    client.GetStream().EndWrite(asyncResult);
                }
                catch
                {
                    // For the purpose of the example, any exception obtaining
                    // or writing to the network should just terminate the 
                    // download.
                    fileStream.Close();
                    return;
                }
            }

            if (!stopDataTransfer && !Recipe10_11Server.Terminate)
            {
                // Read the next block from the file.
                int bytesRead = fileStream.Read(buffer, 0, buffer.Length);

                // If no bytes are read, the stream is at the end of the file.
                if (bytesRead > 0)
                {
                    Console.WriteLine(ID + ": Streaming next block.");

                    // Write the next block to the network stream.
                    client.GetStream().BeginWrite(buffer, 0, buffer.Length,
                        StreamData, null);
                }
                else
                {
                    // End the operation.
                    Console.WriteLine(ID + ": File streaming complete.");
                    fileStream.Close();
                }
            }
            else
            {
                // Client disconnected.
                Console.WriteLine(ID + ": Client disconnected.");
                fileStream.Close();
            }
        }
    }
}