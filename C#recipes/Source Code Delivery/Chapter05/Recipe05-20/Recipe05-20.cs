using System;
using System.IO.Ports;

namespace Apress.VisualCSharpRecipes.Chapter05
{
    static class Recipe05_20
    {
        static void Main(string[] args)
        {
            using (SerialPort port = new SerialPort("COM1"))
            {
                // Set the properties.
                port.BaudRate = 9600;
                port.Parity = Parity.None;
                port.ReadTimeout = 10;
                port.StopBits = StopBits.One;

                // Write a message into the port.
                port.Open();
                port.Write("Hello world!");

                Console.WriteLine("Wrote to the port.");
            }
            
            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
