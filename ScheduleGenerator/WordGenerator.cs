using Microsoft.Office.Interop.Word;
using Nager.Date.Model;
using System;
using System.Collections.Generic;
using System.Text;
using _Word = Microsoft.Office.Interop.Word;

namespace ScheduleGenerator
{
    class WordGenerator
    {
        public void generateWord()
        {
            EmployeesList employeesList = new EmployeesList();
            MonthDays month = new MonthDays();
            string[] monthsWeekDays = month.getNextMonthDays();

            //---Creates a document and sets orientation to landscape---//
            var wordApp = new _Word.Application();

            object oMissing = System.Reflection.Missing.Value;

            _Word.Document oDoc = wordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            oDoc.PageSetup.Orientation = _Word.WdOrientation.wdOrientLandscape;


            //--------------------Inserts a table with borders----------------//

            //-----------------First page---------------------------//

            _Word.Range tableLocation = oDoc.Range(0, 0);

            Table table1 = oDoc.Tables.Add(tableLocation, employeesList.Employees.Count + 1, 20);

            table1.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table1.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;


            //----------Page break-------------------//
            oDoc.Words.Last.InsertBreak(_Word.WdBreakType.wdPageBreak);

            //-----------------Second page-------------------------//
            

            //---------------Sets the height and width of table cells-----------//

            table1.Columns[1].PreferredWidth = wordApp.CentimetersToPoints(1.0f);
            table1.Columns[2].PreferredWidth = wordApp.CentimetersToPoints(2.5f);
            table1.Columns[3].PreferredWidth = wordApp.CentimetersToPoints(2.5f);
            table1.Columns[4].PreferredWidth = wordApp.CentimetersToPoints(1.6f);
            for (int i = 5; i < 21; i++)
            {
                table1.Columns[i].PreferredWidth = wordApp.CentimetersToPoints(1.15f);
            }

            table1.Rows[1].Height = wordApp.CentimetersToPoints(1.8f);
            for (int i = 2; i < 10; i++)
            {
                table1.Rows[i].Height = wordApp.CentimetersToPoints(1.15f);
            }


            table1.Range.Font.Size = 10;

            //-----------------Populates the table----------------------------//

            //----------------First page------------------------------//

            //--------Default values-------------------//
            table1.Cell(1, 1).Range.Text = "Eil. Nr.";
            table1.Cell(1, 2).Range.Text = "Vardas, Pavarde";
            table1.Cell(1, 3).Range.Text = "Pareigos";
            table1.Cell(1, 4).Range.Text = "Nustat. Darbo val. sk.";
            for (int d = 5; d < 21; d++)
            {
                table1.Cell(1, d).Range.Text = (d - 4).ToString();
            }

            //------------Data from EmployeesList-----------------//
            int row = 2;
            foreach (var employee in employeesList.Employees)
            {
                table1.Cell(row, 1).Range.Text = (row - 1).ToString();
                table1.Cell(row, 2).Range.Text = employee.Name;
                table1.Cell(row, 3).Range.Text = employee.Position;
                string[] employeeSchedule = employee.getMonthSchedule();
                int col = 5;
                foreach (var dayHours in employeeSchedule)
                {
                    table1.Cell(row, col).Range.Text = dayHours;
                    col++;
                    if (col > 20)
                    {
                        break;
                    }
                }
                col = 5;
                row++;
            }
            

            //----------------Saves the document and opens it-----------------//
            wordApp.ActiveDocument.SaveAs2("Grafikas_" + MonthNameInLT());
            wordApp.Visible = true;
            oDoc.Activate();
        }

        public string MonthNameInLT()
        {
            MonthDays month = new MonthDays();
            string[] namesInLT = new string[]
            {   
                "Sausis",
                "Vasaris",
                "Kovas",
                "Balandis",
                "Geguze",
                "Birzelis",
                "Liepa",
                "Rugpjutis",
                "Rugsejis",
                "Spalis",
                "Lapkritis",
                "Gruodis"
            };
            return namesInLT[month.nextMonth.Month - 1];

        }
    }
}
