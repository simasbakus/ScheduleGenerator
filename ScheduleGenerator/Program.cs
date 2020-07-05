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
            EmployeesList employeesList = new EmployeesList();
            employeesList.test();

            //ExcellGenerator excel = new ExcellGenerator();
            //excel.generateExcel();

        }
    }
}
