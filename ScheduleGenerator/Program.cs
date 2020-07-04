using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace ScheduleGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            /*MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();

            // Test
            foreach (var day in monthsWeekDays)
            {
                Console.WriteLine(day);
            }*/


            string path = "C:/Users/simas/OneDrive/Documents/Programavimas/ScheduleGenerator/ScheduleGenerator/Employees.json";

            EmployeesList employeesList = JsonConvert.DeserializeObject<EmployeesList>(new StreamReader(path).ReadToEnd());
            
            

            Console.WriteLine(employeesList.employee.Count);
        }
    }
}
