using System;
using System.Runtime.InteropServices;

namespace Apress.VisualCSharpRecipes.Chapter12
{
    class Recipe12_03
    {
        // Declare the external function.
        [DllImport("kernel32.dll")]
        public static extern bool GetVersionEx([In, Out] OSVersionInfo osvi);

        static void Main(string[] args)
        {
            OSVersionInfo osvi = new OSVersionInfo();
            osvi.dwOSVersionInfoSize = Marshal.SizeOf(osvi);

            // Obtain the OS version info.
            GetVersionEx(osvi);

            // Display the version information.
            Console.WriteLine("Class size: " + osvi.dwOSVersionInfoSize);
            Console.WriteLine("Major Version: " + osvi.dwMajorVersion);
            Console.WriteLine("Minor Version: " + osvi.dwMinorVersion);
            Console.WriteLine("Build Number: " + osvi.dwBuildNumber);
            Console.WriteLine("Platform Id: " + osvi.dwPlatformId);
            Console.WriteLine("CSD Version: " + osvi.szCSDVersion);
            Console.WriteLine("Platform: " + Environment.OSVersion.Platform);
            Console.WriteLine("Version: " + Environment.OSVersion.Version);

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter.");
            Console.ReadLine();
        }
    }

    // Define the structure and specify the layout type as sequential.
    [StructLayout(LayoutKind.Sequential)]
    public class OSVersionInfo
    {
        public int dwOSVersionInfoSize;
        public int dwMajorVersion;
        public int dwMinorVersion;
        public int dwBuildNumber;
        public int dwPlatformId;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
        public String szCSDVersion;
    }
}
