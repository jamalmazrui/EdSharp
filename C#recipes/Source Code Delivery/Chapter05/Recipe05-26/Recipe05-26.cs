using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.IO.Pipes;

namespace Recipe05_26
{
    class Recipe05_26
    {
        static void Main(string[] args)
        {
            if (args.Length > 0 && args[0] == "client")
            {
                pipeClient();
            }
            else
            {
                pipeServer();
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }

        static void pipeServer()
        {
            // create the server pipe
            NamedPipeServerStream pipestream = new NamedPipeServerStream("recipe_05_26_pipe");
            // wait for a client to connect
            Console.WriteLine("Waiting for a client connection");
            pipestream.WaitForConnection();

            Console.WriteLine("Received a client connection");
            // wrap a stream writer and stream reader around the pipe
            StreamReader reader = new StreamReader(pipestream);
            StreamWriter writer = new StreamWriter(pipestream);

            // write some data to the pipe
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine("Writing message ", i);
                writer.WriteLine("Message {0}", i);
                writer.Flush();
            }

            // read  data from the pipe
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine("Received: {0}", reader.ReadLine()); ;
            }
            // close the pipe
            pipestream.Close();
        }

        static void pipeClient()
        {
            // create the client pipe
            NamedPipeClientStream pipestream = new NamedPipeClientStream("recipe_05_26_pipe");
            // connect to the pipe server
            pipestream.Connect();

            // wrap a reader around the stream
            StreamReader reader = new StreamReader(pipestream);
            StreamWriter writer = new StreamWriter(pipestream);

            // read the data
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine("Received: {0}", reader.ReadLine());
            }

            // write data to the pipe
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine("Writing response ", i);
                writer.WriteLine("Response {0}", i);
                writer.Flush();
            }
            // close the pipe
            pipestream.Close();
        }
    }
}
