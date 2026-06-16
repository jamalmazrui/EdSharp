using System.Security.Permissions;

[assembly:FileIOPermission(SecurityAction.RequestRefuse, Write = @"C:\")]

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_05_RefuseRequest
    {
        // Class implementation...
    }
}