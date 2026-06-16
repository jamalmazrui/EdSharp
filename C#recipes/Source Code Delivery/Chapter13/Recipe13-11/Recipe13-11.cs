using System;

namespace Apress.VisualCSharpRecipes.Chapter13
{
    // An event argument class that contains information about a temperature 
    // change event. An instance of this class is passed with every event.
    public class TemperatureChangedEventArgs : EventArgs
    {
        // Private data members contain the old and new temperature readings.
        private readonly int oldTemperature, newTemperature;

        // Constructor that takes the old and new temperature values.
        public TemperatureChangedEventArgs(int oldTemp, int newTemp)
        {
            oldTemperature = oldTemp;
            newTemperature = newTemp;
        }

        // Read-only properties provide access to the temperature values.
        public int OldTemperature { get { return oldTemperature; } }
        public int NewTemperature { get { return newTemperature; } }
    }

    // A delegate that specifies the signature that all temperature event 
    // handler methods must implement.
    public delegate void TemperatureChangedEventHandler(Object sender,
        TemperatureChangedEventArgs args);

    // A Thermostat observer that displays information about the change in
    // temperature when a temperature change event occurs.
    public class TemperatureChangeObserver
    {
        // A constructor that takes a reference to the Thermostat object that
        // the TemperatureChangeObserver object should observe.
        public TemperatureChangeObserver(Thermostat t)
        {
            // Create a new TemperatureChangedEventHandler delegate instance and 
            // register it with the specified Thermostat.
            t.TemperatureChanged += this.TemperatureChange;
        }

        // The method to handle temperature change events.
        public void TemperatureChange(Object sender,
            TemperatureChangedEventArgs temp)
        {
            Console.WriteLine ("ChangeObserver: Old={0}, New={1}, Change={2}",
                temp.OldTemperature, temp.NewTemperature,
                temp.NewTemperature - temp.OldTemperature);
        }
    }

    // A Thermostat observer that displays information about the average
    // temperature when a temperature change event occurs.
    public class TemperatureAverageObserver
    {
        // Sum contains the running total of temperature readings.
        // Count contains the number of temperature events received.
        private int sum = 0, count = 0;

        // A constructor that takes a reference to the Thermostat object that
        // the TemperatureAverageObserver object should observe.
        public TemperatureAverageObserver(Thermostat t)
        {
            // Create a new TemperatureChangedEventHandler delegate instance and 
            // register it with the specified Thermostat.
            t.TemperatureChanged += this.TemperatureChange;
        }

        // The method to handle temperature change events.
        public void TemperatureChange(Object sender,
            TemperatureChangedEventArgs temp)
        {
            count++;
            sum += temp.NewTemperature;

            Console.WriteLine
                ("AverageObserver: Average={0:F}", (double)sum / (double)count);
        }
    }

    // A class that represents a thermostat, which is the source of temperature
    // change events. In the Observer pattern, a Thermostat object is the 
    // Subject that Observers listen to for change notifications.
    public class Thermostat
    {
        // Private field to hold current temperature.
        private int temperature = 0;

        // The event used to maintain a list of observer delegates and raise
        // a temperature change event when a temperature change occurs.
        public event TemperatureChangedEventHandler TemperatureChanged;

        // A protected method used to raise the TemperatureChanged event. 
        // Because events can be triggered only from within the containing 
        // type, using a protected method to raise the event allows derived 
        // classes to provide customized behavior and still be able to raise
        // the base class event.
        virtual protected void OnTemperatureChanged
            (TemperatureChangedEventArgs args)
        {
            // Notify all observers. A test for null indicates whether any
            // observers are registered.
            if (TemperatureChanged != null)
            {
                TemperatureChanged(this, args);
            }
        }

        // Public property to get and set the current temperature. The "set" 
        // side of the property is responsible for raising the temperature
        // change event to notify all observers of a change in temperature.
        public int Temperature
        {
            get { return temperature; }

            set
            {
                // Create a new event argument object containing the old and 
                // new temperatures.
                TemperatureChangedEventArgs args =
                    new TemperatureChangedEventArgs(temperature, value);

                // Update the current temperature.
                temperature = value;

                // Raise the temperature change event.
                OnTemperatureChanged(args);
            }
        }
    }

    // A class to demonstrate the use of the Observer pattern.
    public class Recipe13_11
    {
        public static void Main()
        {
            // Create a Thermostat instance.
            Thermostat t = new Thermostat();

            // Create the Thermostat observers.
            new TemperatureChangeObserver(t);
            new TemperatureAverageObserver(t);

            // Loop, getting temperature readings from the user.
            // Any noninteger value will terminate the loop.
            do
            {
                Console.WriteLine(Environment.NewLine);
                Console.Write("Enter current temperature: ");

                try
                {
                    // Convert the user's input to an integer and use it to set
                    // the current temperature of the Thermostat.
                    t.Temperature = Int32.Parse(Console.ReadLine());
                }
                catch (Exception)
                {
                    // Use the exception condition to trigger termination.
                    Console.WriteLine("Terminating Observer Pattern Example.");

                    // Wait to continue.
                    Console.WriteLine(Environment.NewLine);
                    Console.WriteLine("Main method complete. Press Enter");
                    Console.ReadLine();
                    return;
                }
            } while (true);
        }
    }
}