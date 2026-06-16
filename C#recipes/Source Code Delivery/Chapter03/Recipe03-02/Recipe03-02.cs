using System;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    // Declare a class that is passed by value. 
    [Serializable]
    public class Recipe03_02MBV 
    {
        public string HomeAppDomain 
        {
            get 
            {
                return AppDomain.CurrentDomain.FriendlyName;
            }
        }
    }

    // Declare a class that is passed by reference.
    public class Recipe03_02MBR: MarshalByRefObject 
    {
        public string HomeAppDomain
        {
            get
            {
                return AppDomain.CurrentDomain.FriendlyName;
            }
        }
    }

    public class Recipe03_02
    {
        public static void Main(string[] args)
        {
            // Create a new Application Domain.
            AppDomain newDomain =
                AppDomain.CreateDomain("My New AppDomain");

            // Instantiate an MBV object in the new Application Domain.
            Recipe03_02MBV mbvObject = 
                (Recipe03_02MBV)newDomain.CreateInstanceFromAndUnwrap(
                    "Recipe03-02.exe", 
                    "Apress.VisualCSharpRecipes.Chapter03.Recipe03_02MBV");

            // Instantiate an MBR object in the new Application Domain.
            Recipe03_02MBR mbrObject = 
                (Recipe03_02MBR)newDomain.CreateInstanceFromAndUnwrap(
                    "Recipe03-02.exe", 
                    "Apress.VisualCSharpRecipes.Chapter03.Recipe03_02MBR");

            // Display the name if the Application Domain in which each of
            // the objects is located.
            Console.WriteLine("Main AppDomain = {0}",
                AppDomain.CurrentDomain.FriendlyName);
            Console.WriteLine("AppDomain of MBV object = {0}", 
                mbvObject.HomeAppDomain);
            Console.WriteLine("AppDomain of MBR object = {0}",
                mbrObject.HomeAppDomain);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
