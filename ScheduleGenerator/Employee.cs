using System;
using System.Collections.Generic;
using System.Text;

namespace ScheduleGenerator
{
    class Employee
    {
        public string Name { get; set; }
        public string Position { get; set; }

        public WorkingHours WorkingHours { get; set; } = new WorkingHours();
    }
}
