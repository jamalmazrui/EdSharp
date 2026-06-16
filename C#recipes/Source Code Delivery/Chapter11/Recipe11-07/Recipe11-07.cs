using System.Security;
using System.Security.Permissions;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_07
    {
        // Define a variable to indicate whether the assembly has write 
        // access to the C:\Data folder.
        private bool canWrite = false;

        public Recipe11_07()
        {
            // Create and configure a FileIOPermission object that represents 
            // write access to the C:\Data folder.
            FileIOPermission fileIOPerm = 
                new FileIOPermission(FileIOPermissionAccess.Write, @"C:\Data");
                            
            // Test if the current assembly has the specified permission.
            canWrite = SecurityManager.IsGranted(fileIOPerm);
        }
    }
}
