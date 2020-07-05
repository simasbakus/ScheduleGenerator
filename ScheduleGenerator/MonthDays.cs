using System;

namespace ScheduleGenerator
{
    public class MonthDays
    {
        public System.DateTime nextMonth = DateTime.Now.AddMonths(1);
        public string[] getNextMonthDays()
        {
            //gets the next month's weekday calendar

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
