﻿using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace ScheduleGenerator
{
    class Program
    {
        static void Main()
        {
            //EmployeesList employeesList = new EmployeesList();
            //employeesList.test();

            WordGenerator word = new WordGenerator();
            word.generateWord();

        }
    }
}
