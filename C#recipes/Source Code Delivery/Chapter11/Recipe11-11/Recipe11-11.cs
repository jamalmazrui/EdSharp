using System;
using System.Security.Permissions;

namespace Apress.VisualCSharpRecipes.Chapter11
{
    class Recipe11_11
    {
        public static void Method1() 
        { 
            // An imperative role-based security demand for the current principal
            // to represent an identity with the name Anya, the roles of the
            // principal are irrelevant.
            PrincipalPermission perm = 
                new PrincipalPermission(@"MACHINE\Anya", null);
                
            // Make the demand.
            perm.Demand();
        }

        public static void Method2() 
        { 
            // An imperative role-based security demand for the current principal
            // to be a member of the roles Managers OR Developers. If the 
            // principal is a member of either role, access is granted. Using the 
            // PrincipalPermission you can express only an OR type relationship.
            // This is because the PrincipalPolicy.Intersect method always 
            // returns an empty permission unless the two inputs are the same.
            // However, you can use code logic to implement more complex 
            // conditions. In this case, the name of the identity is irrelevant.
            PrincipalPermission perm1 = 
                new PrincipalPermission(null, @"MACHINE\Managers");

            PrincipalPermission perm2 = 
                new PrincipalPermission(null, @"MACHINE\Developers");

            // Make the demand.
            perm1.Union(perm2).Demand();
        }

        public static void Method3() 
        { 
            // An imperative role-based security demand for the current principal
            // to represent an identity with the name Anya AND be a member of the
            // Managers role. 
            PrincipalPermission perm = 
                new PrincipalPermission(@"MACHINE\Anya", @"MACHINE\Managers");
                
            // Make the demand
            perm.Demand();
        }

        // A declarative role-based security demand for the current principal
        // to represent an identity with the name Anya, the roles of the
        // principal are irrelevant.
        [PrincipalPermission(SecurityAction.Demand, Name = @"MACHINE\Anya")]
        public static void Method4() 
        {
            // Method implementation...
        }

        // A declarative role-based security demand for the current principal
        // to be a member of the roles Managers OR Developers. If the 
        // principal is a member of either role, access is granted. You 
        // can express only an OR type relationship, not an AND relationship. 
        // The name of the identity is irrelevant.
        [PrincipalPermission(SecurityAction.Demand, Role = @"MACHINE\Managers")]
        [PrincipalPermission(SecurityAction.Demand, Role = @"MACHINE\Developers")]    
        public static void Method5()
        {
            // Method implementation...
        }

        // A declarative role-based security demand for the current principal
        // to represent an identity with the name Anya AND be a member of the
        // Managers role. 
        [PrincipalPermission(SecurityAction.Demand, Name = @"MACHINE\Anya", 
            Role = @"MACHINE\Managers")]
        public static void Method6()
        {
            // Method implementation...
        }
    }
}
