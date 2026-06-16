using System;
using System.Runtime.InteropServices;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    class Recipe12_05
    {
        // Declare the unmanaged functions.
        [DllImport("kernel32.dll")]
        private static extern int FormatMessage(int dwFlags, int lpSource,
          int dwMessageId, int dwLanguageId, ref String lpBuffer, int nSize,
          int Arguments);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int MessageBox(IntPtr hWnd, string pText,
          string pCaption, int uType);

        static void Main(string[] args)
        {
            // Invoke the MessageBox function passing an invalid
            // window handle and thus force an error.
            IntPtr badWindowHandle = (IntPtr)453;
            MessageBox(badWindowHandle, "Message", "Caption", 0);

            // Obtain the error information.
            int errorCode = Marshal.GetLastWin32Error();
            Console.WriteLine(errorCode);
            Console.WriteLine(GetErrorMessage(errorCode));

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }

        // GetErrorMessage formats and returns an error message
        // corresponding to the input errorCode.
        public static string GetErrorMessage(int errorCode)
        {
            int FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x00000100;
            int FORMAT_MESSAGE_IGNORE_INSERTS = 0x00000200;
            int FORMAT_MESSAGE_FROM_SYSTEM = 0x00001000;

            int messageSize = 255;
            string lpMsgBuf = "";
            int dwFlags = FORMAT_MESSAGE_ALLOCATE_BUFFER |
              FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS;

            int retVal = FormatMessage(dwFlags, 0, errorCode, 0,
              ref lpMsgBuf, messageSize, 0);

            if (0 == retVal)
            {
                return null;
            }
            else
            {
                return lpMsgBuf;
            }
        }
    }
}
