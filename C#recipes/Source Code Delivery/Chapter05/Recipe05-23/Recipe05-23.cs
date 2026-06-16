using System;
using System.IO;
using System.IO.Compression;

namespace Recipe05_23
{
    class Recipe05_23
    {
        static void Main(string[] args)
        {
            // create the compression stream
            GZipStream zipout = new GZipStream(
                File.OpenWrite("compressed_data.gzip"),
                CompressionMode.Compress);
            // wrap the gzip stream in a stream writer
            StreamWriter writer = new StreamWriter(zipout);

            // write the data to the file
            writer.WriteLine("the quick brown fox");
            // close the streams
            writer.Close();

            // open the same file so we can read the 
            // data and decompress it
            GZipStream zipin = new GZipStream(
                File.OpenRead("compressed_data.gzip"),
                CompressionMode.Decompress);
            // wrap the gzip stream in a stream reader
            StreamReader reader = new StreamReader(zipin);

            // read a line from the stream and print it out
            Console.WriteLine(reader.ReadLine());

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
