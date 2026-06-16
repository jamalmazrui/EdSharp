using System;
using System.Text;
using System.Runtime.InteropServices;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    class Recipe12_04
    {
        // The signature for the callback method.
        public delegate bool CallBack(IntPtr hwnd, int lParam);

        // The unmanaged function that will trigger the callback
        // as it enumerates the open windows.
        [DllImport("user32.dll")]
        public static extern int EnumWindows(CallBack callback, int param);

        [DllImport("user32.dll")]
        public static extern int GetWindowText(IntPtr hWnd, 
            StringBuilder lpString, int nMaxCount);

        static void Main(string[] args)
        {
            // Request that the operating system enumerate all windows,
            // and trigger your callback with the handle of each one. 
            EnumWindows(new CallBack(DisplayWindowInfo), 0);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }

        // The method that will receive the callback. The second 
        // parameter is not used, but is needed to match the 
        // callback's signature.
        public static bool DisplayWindowInfo(IntPtr hWnd, int lParam)
        {
            int chars = 100;
            StringBuilder buf = new StringBuilder(chars);
            if (GetWindowText(hWnd, buf, chars) != 0)
            {
                Console.WriteLine(buf);
            }
            return true;
        }
    }
}
