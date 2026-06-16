using System;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    // Implement the IDisposable interface.
    public class DisposeExample : IDisposable
    {
        // Private data member to signal if the object has already been 
        // disposed.
        bool isDisposed = false;

        // Private data member that holds the handle to an unmanaged resource.
        private IntPtr resourceHandle;

        // Constructor.
        public DisposeExample()
        {
            // Constructor code obtains reference to unmanaged resource.
            resourceHandle = default(IntPtr);
        }

        // Destructor / Finalizer. Because Dispose calls GC.SuppressFinalize,
        // this method is called by the garbage collection process only if
        // the consumer of the object does not call Dispose as it should.
        ~DisposeExample()
        {
            // Call the Dispose method as opposed to duplicating the code to 
            // clean up any unmanaged resources. Use the protected Dispose
            // overload and pass a value of "false" to indicate that Dispose is
            // being called during the garbage collection process, not by
            // consumer code.
            Dispose(false);
        }

        // Public implementation of the IDisposable.Dispose method, called
        // by the consumer of the object in order to free unmanaged resources
        // deterministically. 
        public void Dispose()
        {
            // Call the protected Dispose overload and pass a value of "true"
            // to indicate that Dispose is being called by consumer code, not
            // by the garbage collector.
            Dispose(true);

            // Because the Dispose method performs all necessary cleanup,
            // ensure the garbage collector does not call the class destructor.
            GC.SuppressFinalize(this);
        }

        // Protected overload of the Dispose method. The disposing argument 
        // signals whether the method is called by consumer code (true), or by 
        // the garbage collector (false). Note that this method is not part of
        // the IDisposable interface because it has a different signature to the
        // parameterless Dispose method.
        protected virtual void Dispose(bool disposing)
        {
            // Don't try to Dispose of the object twice.
            if (!isDisposed)
            {
                // Determine if consumer code or the garbage collector is 
                // calling. Avoid referencing other managed objects during
                // finalization.
                if (disposing)
                {
                    // Method called by consumer code. Call the Dispose method
                    // of any managed data members that implement the 
                    // IDisposable interface. 
                    // ...
                }

                // Whether called by consumer code or the garbage collector,
                // free all unmanaged resources and set the value of managed 
                // data members to null.
                // Close(resourceHandle);

                // In the case of an inherited type, call base.Dispose(disposing).
            }

            // Signal that this object has been disposed.
            isDisposed = true;
        }

        // Before executing any functionality, ensure that Dispose has not 
        // already been executed on the object.
        public void SomeMethod()
        {
            // Throw an exception if the object has already been disposed.
            if (isDisposed)
            {
                throw new ObjectDisposedException("DisposeExample");
            }

            // Execute method functionality.
            // ...
        }
    }

    // A class to demonstrate the use of DisposeExample
    public class Recipe13_06
    {
        public static void Main()
        {
            // The using statement ensures the Dispose method is called
            // even if an exception occurs.
            using (DisposeExample d = new DisposeExample())
            {
                // Do something with d
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}