using System.Security.Permissions;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    [PublisherIdentityPermission(SecurityAction.InheritanceDemand, 
        CertFile = "pubcert.cer")]
    public class Recipe11_08
    {
        [PermissionSet(SecurityAction.InheritanceDemand, Name="FullTrust")]
        public void SomeProtectedMethod () 
        {
            // Method implementation...
        }
    }
}