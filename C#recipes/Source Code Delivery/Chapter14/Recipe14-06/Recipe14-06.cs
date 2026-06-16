using System;
using System.Timers;
using System.ServiceProcess;

namespace Apress.VisualCSharpRecipes.Chapter14
{
    class Recipe14_06 : ServiceBase
    {
        // A Timer that controls how frequently the example writes to the
        // event log.
        private System.Timers.Timer timer;

        public Recipe14_06()
        {
            // Set the ServiceBase.ServiceName property.
            ServiceName = "Recipe 14_06 Service";

            // Configure the level of control available on the service.
            CanStop = true;
            CanPauseAndContinue = true;
            CanHandleSessionChangeEvent = true;

            // Configure the service to log important events to the 
            // Application event log automatically.
            AutoLog = true;
        }

        // The method executed when the timer expires and writes an
        // entry to the Application event log.
        private void WriteLogEntry(object sender, ElapsedEventArgs e)
        {
            // Use the EventLog object automatically configured by the 
            // ServiceBase class to write to the event log. 
            EventLog.WriteEntry("Recipe14_06 Service active : " + e.SignalTime);
        }

        protected override void OnStart(string[] args)
        {
            // Obtain the interval between log entry writes from the first
            // argument. Use 5000 milliseconds by default and enforce a 1000
            // millisecond minimum.
            double interval;

            try
            {
                interval = Double.Parse(args[0]);
                interval = Math.Max(1000, interval);
            }
            catch
            {
                interval = 5000;
            }

            EventLog.WriteEntry(String.Format("Recipe14_06 Service starting. " +
                "Writing log entries every {0} milliseconds...", interval));

            // Create, configure, and start a System.Timers.Timer to 
            // periodically call the WriteLogEntry method. The Start
            // and Stop methods of the System.Timers.Timer class
            // make starting, pausing, resuming, and stopping the 
            // service straightforward.
            timer = new Timer();
            timer.Interval = interval;
            timer.AutoReset = true;
            timer.Elapsed += new ElapsedEventHandler(WriteLogEntry);
            timer.Start();
        }

        protected override void OnStop()
        {
            EventLog.WriteEntry("Recipe14_06 Service stopping...");
            timer.Stop();

            // Free system resources used by the Timer object.
            timer.Dispose();
            timer = null;
        }

        protected override void OnPause()
        {
            if (timer != null)
            {
                EventLog.WriteEntry("Recipe14_06 Service pausing...");
                timer.Stop();
            }
        }

        protected override void OnContinue()
        {
            if (timer != null)
            {
                EventLog.WriteEntry("Recipe14_06 Service resuming...");
                timer.Start();
            }
        }

        protected override void OnSessionChange(SessionChangeDescription change)
        {
            EventLog.WriteEntry("Recipe14_06 Session change..." +
                change.Reason);
        }

        public static void Main()
        {
            // Create an instance of the Recipe14_06 class that will write 
            // an entry to the Application event log. Pass the object to the
            // static ServiceBase.Run method.
            ServiceBase.Run(new Recipe14_06());
        }
    }
}
