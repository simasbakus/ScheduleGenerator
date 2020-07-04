using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ScheduleGenerator
{
    class JsonParser
    {
        static string path = @"C:\Users\simas\OneDrive\Documents\Programavimas\ScheduleGenerator\ScheduleGenerator\Employees.json";
        public List<Employee> singleEmployee { get; set; } = JsonConvert.DeserializeObject<List<Employee>>(new StreamReader(path).ReadToEnd());

        //------------------method for testing if json parsed-------------------------------//
        public void test()
        {
            for (int i = 0; i < singleEmployee.Count; i++)
            {
                Console.WriteLine(singleEmployee[i].Name);
            }
        }
    }
}
