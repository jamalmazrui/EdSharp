using System.Security;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_03
    {
        // A method to turn on execution permission checking 
        // and persist the change.
        public void ExecutionCheckOn()
        {
            // Turn on execution permission checks.
            SecurityManager.CheckExecutionRights = true;

            // Persist the configuration change.
            SecurityManager.SavePolicy();
        }

        // A method to turn off execution permission checking 
        // and persist the change.
        public void ExecutionCheckOff()
        {
            // Turn off execution permission checks.
            SecurityManager.CheckExecutionRights = false;

            // Persist the configuration change.
            SecurityManager.SavePolicy();
        }
    }
}
