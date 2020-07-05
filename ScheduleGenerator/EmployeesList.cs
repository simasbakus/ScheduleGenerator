using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ScheduleGenerator
{
    class EmployeesList
    {
        static string path = @"C:\Users\simas\OneDrive\Documents\Programavimas\ScheduleGenerator\ScheduleGenerator\Employees.json";
        public List<Employee> Employees { get; set; } = JsonConvert.DeserializeObject<List<Employee>>(new StreamReader(path).ReadToEnd());



        //------------------method for testing if json parsed-------------------------------//
        public void test()
        {
            MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();
            foreach (var person in Employees)
            {
                string[] array = person.getMonthSchedule();
                Console.WriteLine(person.Name);
                Console.WriteLine(person.Position);
                int i = 0;
                foreach (var item in array)
                {
                    Console.WriteLine(monthsWeekDays[i] + ":  " + item);
                    i++;
                }
                Console.WriteLine("");
            }
        }
    }
}
