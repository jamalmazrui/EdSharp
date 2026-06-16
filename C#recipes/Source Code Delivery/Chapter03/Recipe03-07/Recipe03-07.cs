using System;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;

namespace Apress.VisualCSharpRecipes.Chapter03
{
    // A common interface that all plug-ins must implement.
    public interface IPlugin
    {
        void Start();
        void Stop();
    }

    // A simple IPlugin implementation to demonstrate the PluginManager 
    // controller class.
    public class SimplePlugin : IPlugin
    {
        public void Start()
        {
            Console.WriteLine(AppDomain.CurrentDomain.FriendlyName +
                ": SimplePlugin starting...");
        }

        public void Stop()
        {
            Console.WriteLine(AppDomain.CurrentDomain.FriendlyName +
                ": SimplePlugin stopping...");
        }
    }

    // The controller class, which manages the loading and manipulation
    // of plug-ins in its application domain.
    public class PluginManager : MarshalByRefObject
    {
        // A Dictionary to hold keyed references to IPlugin instances.
        private Dictionary<string, IPlugin> plugins = 
            new Dictionary<string, IPlugin> ();

        // Default constructor.
        public PluginManager() { }

        // Constructor that loads a set of specified plug-ins on creation.
        public PluginManager(NameValueCollection pluginList)
        {
            // Load each of the specified plug-ins.
            foreach (string plugin in pluginList.Keys)
            {
                this.LoadPlugin(pluginList[plugin], plugin);
            }
        }

        // Load the specified assembly and instantiate the specified
        // IPlugin implementation from that assembly.
        public bool LoadPlugin(string assemblyName, string pluginName)
        {
            try
            {
                // Load the named private assembly.
                Assembly assembly = Assembly.Load(assemblyName);

                // Create the IPlugin instance, ignore case.
                IPlugin plugin = assembly.CreateInstance(pluginName, true) 
                    as IPlugin;

                if (plugin != null)
                {
                    // Add new IPlugin to ListDictionary
                    plugins[pluginName] = plugin;

                    return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                // Return false on all exceptions for the purpose of
                // this example. Do not suppress exceptions like this
                // in production code.
                return false;
            }
        }

        public void StartPlugin(string plugin)
        {
            try
            {
                // Extract the IPlugin from the Dictionary and call Start.
                plugins[plugin].Start();
            }
            catch
            {
                // Log or handle exceptions appropriately
            }
        }

        public void StopPlugin(string plugin)
        {
            try
            {
                // Extract the IPLugin from the Dictionary and call Stop.
                plugins[plugin].Stop();
            }
            catch
            {
                // Log or handle exceptions appropriately
            }
        }

        public ArrayList GetPluginList()
        {
            // Return an enumerable list of plugin names. Take the keys
            // and place them in an ArrayList, which supports marshal-by-value.
            return new ArrayList(plugins.Keys);
        }
    }

    class Recipe03_07
    {
        public static void Main()
        {
            // Create a new application domain.
            AppDomain domain1 = AppDomain.CreateDomain("NewAppDomain1");

            // Create a PluginManager in the new application domain using 
            // the default constructor.
            PluginManager manager1 = 
                (PluginManager)domain1.CreateInstanceAndUnwrap("Recipe03-07", 
                "Apress.VisualCSharpRecipes.Chapter03.PluginManager");

            // Load a new plugin into NewAppDomain1.
            manager1.LoadPlugin("Recipe03-07", 
                "Apress.VisualCSharpRecipes.Chapter03.SimplePlugin");

            // Start and stop the plug-in in NewAppDomain1.
            manager1.StartPlugin(
                "Apress.VisualCSharpRecipes.Chapter03.SimplePlugin");
            manager1.StopPlugin(
                "Apress.VisualCSharpRecipes.Chapter03.SimplePlugin");

            // Create a new application domain.
            AppDomain domain2 = AppDomain.CreateDomain("NewAppDomain2");

            // Create a ListDictionary containing a list of plug-ins to create.
            NameValueCollection pluginList = new NameValueCollection();
            pluginList["Apress.VisualCSharpRecipes.Chapter03.SimplePlugin"] = 
                "Recipe03-07";

            // Create a PluginManager in the new application domain and 
            // specify the default list of plug-ins to create.
            PluginManager manager2 = (PluginManager)domain1.CreateInstanceAndUnwrap(
                "Recipe03-07", "Apress.VisualCSharpRecipes.Chapter03.PluginManager", 
                true, 0, null, new object[] { pluginList }, null, null);

            // Display the list of plug-ins loaded into NewAppDomain2.
            Console.WriteLine("\nPlugins in NewAppDomain2:");
            foreach (string s in manager2.GetPluginList())
            {
                Console.WriteLine(" - " + s);
            }

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
