using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using _Word = Microsoft.Office.Interop.Word;

namespace ScheduleGenerator
{
    class ExcellGenerator
    {
        public void generateExcel()
        {
            var app = new _Excel.Application();
            app.Visible = true;
            app.Workbooks.Add();
            _Excel._Worksheet workSheet = (_Excel.Worksheet)app.ActiveSheet;
            workSheet.Cells[1, 1] = "Hello World";



            var wordApp = new _Word.Application();

            object oMissing = System.Reflection.Missing.Value;

            _Word.Document oDoc = wordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            wordApp.ActiveDocument.SaveAs2("Test");
        }
    }
}
