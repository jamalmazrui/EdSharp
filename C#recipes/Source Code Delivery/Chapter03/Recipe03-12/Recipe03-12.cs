using System;
using System.Text;
using System.Reflection;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    class Recipe03_12
    {
        public static StringBuilder CreateStringBuilder()
        {
            // Obtain the Type for the StringBuilder class.
            Type type = typeof(StringBuilder);

            // Create a Type[] containing Type instances for each
            // of the constructor arguments - a string and an int.
            Type[] argTypes = new Type[] { typeof(System.String), typeof(System.Int32) };

            // Obtain the ConstructorInfo object.
            ConstructorInfo cInfo = type.GetConstructor(argTypes);

            // Create an object[] containing the constructor arguments.
            object[] argVals = new object[] { "Some string", 30 };

            // Create the object and cast it to StringBuilder.
            StringBuilder sb = (StringBuilder)cInfo.Invoke(argVals);

            return sb;
        }
    }
}

namespace Apress.VisualCSharpRecipes.Chapter03
{
    // A common interface that all plugins must implement.
    public interface IPlugin
    {
        string Description { get; set; }
        void Start();
        void Stop();
    }

    // An abstract base class from which all plug-ins must derive.
    public abstract class AbstractPlugin : IPlugin
    {
        // Hold a description for the plug-in instance.
        private string description = "";

        // Sealed property to get the plug-in description.
        public string Description
        {
            get { return description; }
            set { description = value; }
        }

        // Declare the members of the IPlugin interface as abstract.
        public abstract void Start();
        public abstract void Stop();
    }

    // A simple IPlugin implementation to demonstrate the PluginFactory class.
    public class SimplePlugin : AbstractPlugin
    {
        // Implement Start method.
        public override void Start()
        {
            Console.WriteLine(Description + ": Starting...");
        }

        // Implement Stop method.
        public override void Stop()
        {
            Console.WriteLine(Description + ": Stopping...");
        }
    }

    // A factory to instantiate instances of IPlugin.
    public sealed class PluginFactory
    {
        public static IPlugin CreatePlugin(string assembly,
            string pluginName, string description)
        {
            // Obtain the Type for the specified plug-in.
            Type type = Type.GetType(pluginName + ", " + assembly);

            // Obtain the ConstructorInfo object.
            ConstructorInfo cInfo = type.GetConstructor(Type.EmptyTypes);

            // Create the object and cast it to StringBuilder.
            IPlugin plugin = cInfo.Invoke(null) as IPlugin;

            // Configure the new IPlugin.
            plugin.Description = description;

            return plugin;
        }

        public static void Main(string[] args)
        {
            // Instantiate a new IPlugin using the PluginFactory.
            IPlugin plugin = PluginFactory.CreatePlugin(
                "Recipe03-12",  // Private assembly name
                "Apress.VisualCSharpRecipes.Chapter03.SimplePlugin", // Plugin class name
                "A Simple Plugin"       // Plugin instance description
            );

            // Start and stop the new plugin.
            plugin.Start();
            plugin.Stop();

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}

