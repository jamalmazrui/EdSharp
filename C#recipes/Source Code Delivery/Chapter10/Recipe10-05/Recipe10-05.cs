using System;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter10
{
    class Recipe10_05
    {
        // Configure the maximum number of request that can be 
        // handled concurrently.
        private static int maxRequestHandlers = 5;

        // An integer used to assign each HTTP request handler a unique 
        // identifier.
        private static int requestHandlerID = 0;

        // The HttpListener is the class that provides all the capabilities 
        // to receive and process HTTP requests.
        private static HttpListener listener;

        // A method to asynchronously process individual requests and send 
        // responses.
        private static void RequestHandler(IAsyncResult result)
        {
            Console.WriteLine("{0}: Activated.", result.AsyncState);

            try
            {
                // Obtain the HttpListenerContext for the new request.
                HttpListenerContext context = listener.EndGetContext(result);

                Console.WriteLine("{0}: Processing HTTP Request from {1} ({2}).",
                    result.AsyncState,
                    context.Request.UserHostName,
                    context.Request.RemoteEndPoint);

                // Build the response using a StreamWriter feeding the 
                // Response.OutputStream. 
                StreamWriter sw =
                   new StreamWriter(context.Response.OutputStream, Encoding.UTF8);

                sw.WriteLine("<html>");
                sw.WriteLine("<head>");
                sw.WriteLine("<title>Visual C# Recipes</title>");
                sw.WriteLine("</head>");
                sw.WriteLine("<body>");
                sw.WriteLine("Recipe 10-5: " + result.AsyncState);
                sw.WriteLine("</body>");
                sw.WriteLine("</html>");
                sw.Flush();

                // Configure the Response.
                context.Response.ContentType = "text/html";
                context.Response.ContentEncoding = Encoding.UTF8;

                // Close the Response to send it to the client.
                context.Response.Close();

                Console.WriteLine("{0}: Sent HTTP response.", result.AsyncState);
            }
            catch (ObjectDisposedException)
            {
                Console.WriteLine("{0}: HttpListener disposed--shutting down.",
                    result.AsyncState);
            }
            finally
            {
                // Start another handler if unless the HttpListener is closing.
                if (listener.IsListening)
                {
                    Console.WriteLine("{0}: Creating new request handler.",
                        result.AsyncState);

                    listener.BeginGetContext(RequestHandler, "RequestHandler_" +
                        Interlocked.Increment(ref requestHandlerID));
                }
            }
        }

        public static void Main(string[] args)
        {
            // Quit gracefully if this feature is not supported.
            if (!HttpListener.IsSupported)
            {
                Console.WriteLine(
                    "You must be running this example on Windows XP SP2, ",
                    "Windows Server 2003, or higher to create ",
                    "an HttpListener.");
                return;
            }

            // Create the HttpListener.
            using (listener = new HttpListener())
            {
                // Configure the URI prefixes that will map to the HttpListener.
                listener.Prefixes.Add(
                    "http://localhost:19080/VisualCSharpRecipes/");
                listener.Prefixes.Add(
                    "http://localhost:20000/Recipe10-05/");

                // Start the HttpListener before listening for incoming requests.
                Console.WriteLine("Starting HTTP Server");
                listener.Start();
                Console.WriteLine("HTTP Server started");
                Console.WriteLine(Environment.NewLine);

                // Create a number of asynchronous request handlers up to 
                // the configurable maximum. Give each a unique identifier.
                for (int count = 0; count < maxRequestHandlers; count++)
                {
                    listener.BeginGetContext(RequestHandler, "RequestHandler_" +
                        Interlocked.Increment(ref requestHandlerID));
                }

                // Wait for the user to stop the HttpListener.
                Console.WriteLine("Press Enter to stop the HTTP Server");
                Console.ReadLine();

                // Stop accepting new requests.
                listener.Stop();

                // Terminate the HttpListener without processing current requests.
                listener.Abort();
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
