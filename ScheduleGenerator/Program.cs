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
            WordGenerator word = new WordGenerator();
            word.generateWord();
        }
    }
}
