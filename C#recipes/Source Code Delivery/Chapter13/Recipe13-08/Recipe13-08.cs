using System;
using System.Runtime.Serialization;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    // Mark CustomException as Serializable. 
    [Serializable]
    public sealed class CustomException : Exception
    {
        // Custom data members for CustomException.
        private string stringInfo;
        private bool booleanInfo;

        // Three standard constructors and simply call the base class
        // constructor (System.Exception).
        public CustomException() : base() { }

        public CustomException(string message) : base(message) { }

        public CustomException(string message, Exception inner)
            : base(message, inner) { }

        // The deserialization constructor required by the ISerialization 
        // interface. Because CustomException is sealed, this constructor 
        // is private. If CustomException were not sealed, this constructor 
        // should be declared as protected so that derived classes can call
        // it during deserialization.
        private CustomException(SerializationInfo info,
            StreamingContext context) : base(info, context)
        {
            // Deserialize each custom data member.
            stringInfo = info.GetString("StringInfo");
            booleanInfo = info.GetBoolean("BooleanInfo");
        }

        // Additional constructors to allow code to set the custom data 
        // members.
        public CustomException(string message, string stringInfo,
            bool booleanInfo) : this(message)
        {
            this.stringInfo = stringInfo;
            this.booleanInfo = booleanInfo;
        }

        public CustomException(string message, Exception inner,
            string stringInfo, bool booleanInfo): this(message, inner)
        {
            this.stringInfo = stringInfo;
            this.booleanInfo = booleanInfo;
        }

        // Read-only properties that provide access to the custom data members.
        public string StringInfo
        {
            get { return stringInfo; }
        }

        public bool BooleanInfo
        {
            get { return booleanInfo; }
        }

        // The GetObjectData method (declared in the ISerializable interface) 
        // is used during serialization of CustomException. Because 
        // CustomException declares custom data members, it must override the  
        // base class implementation of GetObjectData.
        public override void GetObjectData(SerializationInfo info,
            StreamingContext context)
        {
            // Serialize the custom data members.
            info.AddValue("StringInfo", stringInfo);
            info.AddValue("BooleanInfo", booleanInfo);

            // Call the base class to serialize its members.
            base.GetObjectData(info, context);
        }

        // Override the base class Message property to include the custom data 
        // members.
        public override string Message
        {
            get
            {
                string message = base.Message;
                if (stringInfo != null)
                {
                    message += Environment.NewLine +
                        stringInfo + " = " + booleanInfo;
                }
                return message;
            }
        }
    }

    // A class to demonstrate the use of CustomException.
    public class Recipe13_08
    {
        public static void Main()
        {
            try
            {
                // Create and throw a CustomException object.
                throw new CustomException("Some error",
                    "SomeCustomMessage", true);
            } 
            catch (CustomException ex)
            {
                Console.WriteLine(ex.Message);
            }

            // Wait to continue.
            Console.WriteLine(Environment.NewLine);
            Console.WriteLine("Main method complete. Press Enter");
            Console.ReadLine();
        }
    }
}