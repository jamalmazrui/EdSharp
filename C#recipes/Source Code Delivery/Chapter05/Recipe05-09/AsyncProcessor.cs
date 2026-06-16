using System;
using System.IO;
using System.Threading;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    public class AsyncProcessor
    {
        private Stream inputStream;

        // The amount that will be read in one block (2 KB).
        private int bufferSize = 2048;

        public int BufferSize
        {
            get { return bufferSize; }
            set { bufferSize = value; }
        }

        // The buffer that will hold the retrieved data.
        private byte[] buffer;

        public AsyncProcessor(string fileName)
        {
            buffer = new byte[bufferSize];

            // Open the file, specifying true for asynchronous support.
            inputStream = new FileStream(fileName, FileMode.Open,
              FileAccess.Read, FileShare.Read, bufferSize, true);
        }

        public void StartProcess()
        {

            // Start the asynchronous read, which will fill the buffer.
            inputStream.BeginRead(buffer, 0, buffer.Length,
              OnCompletedRead, null);
        }

        private void OnCompletedRead(IAsyncResult asyncResult)
        {
            // One block has been read asynchronously.
            // Retrieve the data.
            int bytesRead = inputStream.EndRead(asyncResult);

            // If no bytes are read, the stream is at the end of the file.
            if (bytesRead > 0)
            {
                // Pause to simulate processing this block of data.
                Console.WriteLine("\t[ASYNC READER]: Read one block.");
                Thread.Sleep(TimeSpan.FromMilliseconds(20));

                // Begin to read the next block asynchronously.
                inputStream.BeginRead(
                buffer, 0, buffer.Length, OnCompletedRead,
                  null);
            }
            else
            {
                // End the operation.
                Console.WriteLine("\t[ASYNC READER]: Complete.");
                inputStream.Close();
            }
        }
    }
}