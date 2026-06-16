using System;

namespace Apress.VisualCSharpRecipes.Chapter02
{
    class Recipe02_07
    {
        public static void Main(string[] args)
        {
            string ds1 = "Sep 2005";
            string ds2 = "Monday 5 September 2005 14:15:33";
            string ds3 = "5,9,5";
            string ds4 = "5/9/2005 14:15:33";
            string ds5 = "2:15 PM";

            // 1st September 2005 00:00:00
            DateTime dt1 = DateTime.Parse(ds1);

            // 5th September 2005 14:15:33
            DateTime dt2 = DateTime.Parse(ds2);

            // 5th September 2005 00:00:00
            DateTime dt3 = DateTime.Parse(ds3);

            // 5th September 2005 14:15:33
            DateTime dt4 = DateTime.Parse(ds4);

            // Current Date 14:15:00
            DateTime dt5 = DateTime.Parse(ds5);

            // Display the converted DateTime objects.
            Console.WriteLine("String: {0} DateTime: {1}", ds1, dt1);
            Console.WriteLine("String: {0} DateTime: {1}", ds2, dt2);
            Console.WriteLine("String: {0} DateTime: {1}", ds3, dt3);
            Console.WriteLine("String: {0} DateTime: {1}", ds4, dt4);
            Console.WriteLine("String: {0} DateTime: {1}", ds5, dt5);

            // Parse only strings containing LongTimePattern.
            DateTime dt6 = DateTime.ParseExact("2:13:30 PM", "h:mm:ss tt", null);

            // Parse only strings containing RFC1123Pattern.
            DateTime dt7 = DateTime.ParseExact(
                "Mon, 05 Sep 2005 14:13:30 GMT", "ddd, dd MMM yyyy HH':'mm':'ss 'GMT'", null);

            // Parse only strings containing MonthDayPattern.
            DateTime dt8 = DateTime.ParseExact("September 05", "MMMM dd", null);

            // Display the converted DateTime objects.
            Console.WriteLine(dt6);
            Console.WriteLine(dt7);
            Console.WriteLine(dt8);

            // Wait to continue.
            Console.WriteLine("\nMain method complete. Press Enter");
            Console.ReadLine();
        }
    }
}
