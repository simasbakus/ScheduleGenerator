using System;

namespace ScheduleGenerator
{
    public class MonthDays
    {
        public string[] getNextMonthDays()
        {
            //gets the next month's weekday calendar

            var nextMonth = DateTime.Now.AddMonths(1);

            int daysInMonth = DateTime.DaysInMonth(nextMonth.Year, nextMonth.Month);

            string[] monthsWeekDays = new string[daysInMonth];

            for (int i = 0; i < daysInMonth; i++)
            {
                monthsWeekDays[i] = new DateTime(nextMonth.Year, nextMonth.Month, i + 1).DayOfWeek.ToString();
            }

            return monthsWeekDays;
        }
    }
}
