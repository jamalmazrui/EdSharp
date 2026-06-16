using System;
using System.Collections.Generic;
using System.Linq;

namespace Apress.VisualCSharpRecipes.Chapter01
{
    public class WeatherReport
    {
        public int DayOfWeek
        {
            get;
            set;
        }

        public int DailyTemp
        {
            get;
            set;
        }

        public int AveTempSoFar
        {
            get;
            set;
        }
    }

    public class WeatherForecast
    {
        private int[] temps = { 54, 63, 61, 55, 61, 63, 58 };
        IList<string> daysOfWeek = new List<string>() 
            {"Monday", "Tuesday", "Wednesday", 
            "Thursday", "Friday", "Saturday", "Sunday"};

        public WeatherReport this[string dow]
        {
            get
            {
                // get the day of the week index
                int dayindex = daysOfWeek.IndexOf(dow);
                return new WeatherReport()
                {
                    DayOfWeek = dayindex,
                    DailyTemp = temps[dayindex],
                    AveTempSoFar = calculateTempSoFar(dayindex)
                };
            }
            set
            {
                temps[daysOfWeek.IndexOf(dow)] = value.DailyTemp;
            }
        }

        private int calculateTempSoFar(int dayofweek)
        {
            int[] subset = new int[dayofweek + 1];
            Array.Copy(temps, 0, subset, 0, dayofweek + 1);
            return (int)subset.Average();
        }
    }

    public class Recipe01_24
    {
        static void Main(string[] args)
        {
            // create a new weather forecast
            WeatherForecast forecast = new WeatherForecast();

            // use the indexer to obtain forecast values and write them out
            string[] days = {"Monday", "Thursday", "Tuesday", "Saturday"};
            foreach (string day in days)
            {
                WeatherReport report = forecast[day];
                Console.WriteLine("Day: {0} DayIndex {1}, Temp: {2} Ave {3}", day, 
                    report.DayOfWeek, report.DailyTemp, report.AveTempSoFar);
            }

            // change one of the temperatures
            forecast["Tuesday"] = new WeatherReport()
            {
                DailyTemp = 34
            };

            // repeat the loop
            Console.WriteLine("\nModified results...");
            foreach (string day in days)
            {
                WeatherReport report = forecast[day];
                Console.WriteLine("Day: {0} DayIndex {1}, Temp: {2} Ave {3}", day,
                    report.DayOfWeek, report.DailyTemp, report.AveTempSoFar);
            }                

            Console.WriteLine("\nMain method complete. Press Enter.");
            Console.ReadLine();
        }
    }
}
