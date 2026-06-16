using System;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_06
    {
        public static void Main() 
        {
            // Create the array to hold the Regex info objects.
            RegexCompilationInfo[] regexInfo = new RegexCompilationInfo[2];

            // Create the RegexCompilationInfo for PinRegex.
            regexInfo[0] = new RegexCompilationInfo(@"^\d{4}$", 
                RegexOptions.Compiled, "PinRegex", "", true);

            // Create the RegexCompilationInfo for CreditCardRegex.
            regexInfo[1] = new RegexCompilationInfo( 
                @"^\d{4}-?\d{4}-?\d{4}-?\d{4}$", 
                RegexOptions.Compiled, "CreditCardRegex", "", true);

            // Create the AssemblyName to define the target assembly.
            AssemblyName assembly = new AssemblyName();
            assembly.Name = "MyRegEx";

            // Create the compiled regular expression
            Regex.CompileToAssembly(regexInfo, assembly);
        }
    }
}