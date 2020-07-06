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

    }
}
