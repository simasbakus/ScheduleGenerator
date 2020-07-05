using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace ScheduleGenerator
{
    class Program
    {
        static void Main()
        {
            /*MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();

            // Tests
            foreach (var day in monthsWeekDays)
            {
                Console.WriteLine(day);
            }*/

            EmployeesList employeesList = new EmployeesList();
            employeesList.test();
        }
    }
}
