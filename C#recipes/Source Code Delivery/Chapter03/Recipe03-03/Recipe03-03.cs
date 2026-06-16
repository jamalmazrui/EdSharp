using System;
using System.Data;
using System.Runtime.Remoting;

namespace Apress.VisualCSharpRecipes.Chapter03 {
    class Recipe03_03 {
        // A method to wrap a Dataset.
        public static ObjectHandle WrapDataSet(DataSet ds) {
            // Wrap the DataSet.
            ObjectHandle objHandle = new ObjectHandle(ds);

            // Return the wrapped DataSet.
            return objHandle;
        }

        // A method to unwrap a DataSet.
        public static DataSet UnwrapDataSet(ObjectHandle handle) {
            // Unwrap the DataSet.
            DataSet ds = (System.Data.DataSet)handle.Unwrap();

            // Return the wrapped DataSet.
            return ds;
        }

        public static void Main() {
            DataSet ds = new DataSet();
            Console.WriteLine(ds.ToString());

            ObjectHandle oh = WrapDataSet(ds);
            DataSet ds2 = UnwrapDataSet(oh);
            Console.WriteLine(ds2.ToString());

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
