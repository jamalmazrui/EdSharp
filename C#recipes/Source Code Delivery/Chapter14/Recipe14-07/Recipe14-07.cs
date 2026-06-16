using System.Configuration.Install;
using System.ServiceProcess;
using System.ComponentModel;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    [RunInstaller(true)]
    public class Recipe14_07 : Installer
    {
        public Recipe14_07()
        {
            // Instantiate and configure a ServiceProcessInstaller.
            ServiceProcessInstaller ServiceExampleProcess =
                new ServiceProcessInstaller();
            ServiceExampleProcess.Account = ServiceAccount.LocalSystem;

            // Instantiate and configure a ServiceInstaller.
            ServiceInstaller ServiceExampleInstaller =
                new ServiceInstaller();
            ServiceExampleInstaller.DisplayName = 
                "Visual C# Recipes Service Example";
            ServiceExampleInstaller.ServiceName = "Recipe 14_06 Service";
            ServiceExampleInstaller.StartType = ServiceStartMode.Automatic;

            // Add both the ServiceProcessInstaller and ServiceInstaller to 
            // the Installers collection, which is inherited from the 
            // Installer base class.
            Installers.Add(ServiceExampleInstaller);
            Installers.Add(ServiceExampleProcess);
        }
    }
}
