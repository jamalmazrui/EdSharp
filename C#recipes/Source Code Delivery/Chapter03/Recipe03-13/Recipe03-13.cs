using System;

// Declare Allen as the assembly author. Assembly attributes
// must be declared after "using" statements but before any other.
// Author name is a positional parameter.
// Company name is a named parameter.
[assembly: Apress.VisualCSharpRecipes.Chapter03.Author("Allen",
    Company = "Apress")]

namespace Apress.VisualCSharpRecipes.Chapter03
{
    // Declare a class authored by Allen. 
    [Author("Allen", Company = "Apress")]
    public class SomeClass
    {
        // Class implementation.
    }

    // Declare a class authored by Lena.
    [Author("Lena")]
    public class SomeOtherClass
    {
        // Class implementation.
    }
}

