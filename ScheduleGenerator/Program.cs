using System;

namespace ScheduleGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();
            foreach (var day in monthsWeekDays)
            
            // Test
            {
                Console.WriteLine(day);
            }
        }
    }
}
