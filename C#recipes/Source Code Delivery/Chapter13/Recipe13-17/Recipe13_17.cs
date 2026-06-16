using System;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    class Recipe13_17
    {
        static void Main(string[] args)
        {
            // create an implementation of the interface that is typed
            // (and contains) the derived type
            IMyInterface<DerivedType> variant = new ImplClass<DerivedType>(new DerivedType());
            // call a method that accepts an instance of the interface typed
            // with the base type - if the interface has been defined with the out keyword
            // this will work, otherwise the compiler will report an error
            processData(variant);
        }

        static void processData(IMyInterface<BaseType> data)
        {
            data.getValue().printMessage();
        }
    }

    class BaseType
    {
        public virtual void printMessage()
        {
            Console.WriteLine("BaseType Message");
        }
    }

    class DerivedType : BaseType
    {
        public override void printMessage()
        {
            Console.WriteLine("DerivedType Message");
        }
    }

    public interface IMyInterface<out T>
    {
        T getValue();
    }

    public class ImplClass<T> : IMyInterface<T>
    {
        private T value;
        public ImplClass(T val)
        {
            value = val;
        }

        public T getValue()
        {
            return value;
        }
    }

  
}
