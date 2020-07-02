using System;
using System.Collections.Generic;
using System.Text;

namespace ScheduleGenerator
{
    public class MonthDays
    {
        public string[] getNextMonthDays()
        {
            //gets the next month's weekday calendar

            int daysInMonth;
            int year;
            var nextMonth = DateTime.Now.AddMonths(1).Month;

            if (nextMonth > 1)
            {
                year = DateTime.Now.Year;
            } else
            {
                year = DateTime.Now.AddYears(1).Year;
            }

            daysInMonth = DateTime.DaysInMonth(year, nextMonth);

            string[] monthsWeekDays = new string[daysInMonth];

            for (int i = 0; i < daysInMonth; i++)
            {
                monthsWeekDays[i] = new DateTime(year, nextMonth, i + 1).DayOfWeek.ToString();
            }

            return monthsWeekDays;

        }
    }
}
