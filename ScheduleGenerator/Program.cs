using System;

namespace ScheduleGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();

            // Test
            foreach (var day in monthsWeekDays)
            {
                Console.WriteLine(day);
            }
        }
    }
}
