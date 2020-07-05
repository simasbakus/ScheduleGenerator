using Nager.Date;
using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Text;

namespace ScheduleGenerator
{
    class Employee
    {
        public string Name { get; set; }
        public string Position { get; set; }

        public WorkingHours WorkingHours { get; set; } = new WorkingHours();

        public string[] getMonthSchedule()
        {
            MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();
            string[] schedule = new string[monthsWeekDays.Length];

            int i = 0;
            foreach (var day in monthsWeekDays)
            {
                if (DateSystem.IsPublicHoliday(new DateTime(month.nextMonth.Year, month.nextMonth.Month, i + 1), CountryCode.LT))
                {
                    schedule[i] = "P";
                }
                else
                {
                    switch (day)
                    {
                        case "Monday":
                            schedule[i] = WorkingHours.Monday;
                            break;

                        case "Tuesday":
                            schedule[i] = WorkingHours.Tuesday;
                            break;

                        case "Wednesday":
                            schedule[i] = WorkingHours.Wednesday;
                            break;

                        case "Thursday":
                            schedule[i] = WorkingHours.Thursday;
                            break;

                        case "Friday":
                            schedule[i] = WorkingHours.Friday;
                            break;

                        default:
                            schedule[i] = "P";
                            break;
                    }
                }
                i++;
            }
            return schedule;
        }
    }
}
