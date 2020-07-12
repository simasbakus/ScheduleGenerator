using Microsoft.Office.Interop.Word;
using Nager.Date.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.WebSockets;
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
            wordApp.Visible = false;

            object oMissing = System.Reflection.Missing.Value;

            _Word.Document oDoc = wordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            oDoc.PageSetup.Orientation = _Word.WdOrientation.wdOrientLandscape;

            //-----------------Set Margins------------------------//

            oDoc.PageSetup.TopMargin = wordApp.CentimetersToPoints(1.5f);

            //-------------------Adding header text in page one-------------------------//

            _Word.Paragraph text1;

            text1 = oDoc.Content.Paragraphs.Add();
            text1.Range.Text = "UAB G. Tarnauskienes odontologijos klinika \t\t\t\tTvirtinu: G. Tarnauskiene";
            text1.Range.Font.Bold = 1;
            text1.Range.Font.Size = 14;
            text1.Range.Font.Name = "Times New Roman";

            //-------sets different font from 43 char-----//
            object diffFontStart = text1.Range.Start + 43;
            object diffFontEnd = text1.Range.End;
            _Word.Range diffFontRng = oDoc.Range(ref diffFontStart, ref diffFontEnd);
            diffFontRng.Font.Bold = 0;
            diffFontRng.Font.Size = 12;
            //-----------------------------------------------//

            text1.Format.SpaceAfter = 24;
            text1.Range.InsertParagraphAfter();

            _Word.Paragraph text2;

            text2 = oDoc.Content.Paragraphs.Add();
            text2.Range.Text = "\t\t\t Darbo grafikas Nr. " + month.nextMonth.Month + "\t\t\t\t" + DateTime.Today.ToShortDateString();
            text2.Range.Font.Bold = 0;
            text2.Format.SpaceAfter = 1;
            text2.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
            text2.Range.InsertParagraphAfter();

            _Word.Paragraph text3;

            text3 = oDoc.Content.Paragraphs.Add();
            text3.Range.Text = "\t\t\t " + month.nextMonth.Year + "m. " + MonthNameInLT() + " men.";
            text3.Range.Font.Size = 12;
            text3.Format.SpaceAfter = 6;
            text3.Range.InsertParagraphAfter();

            //-------------------------table1-----------------------------//

            object start = oDoc.Content.End - 1;
            object end = oDoc.Content.End;
            _Word.Range rng = oDoc.Range(start, end);

            Table table1;

            table1 = oDoc.Tables.Add(rng, employeesList.Employees.Count + 1, 20);

            table1.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table1.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;


            //----------Page break-------------------//

            oDoc.Words.Last.InsertBreak(_Word.WdBreakType.wdPageBreak);


            //---------------------table2-----------------------------------------//

            start = oDoc.Content.End - 1;
            end = oDoc.Content.End;
            rng = oDoc.Range(start, end);

            Table table2;

            table2 = oDoc.Tables.Add(rng, employeesList.Employees.Count + 1, monthsWeekDays.Length - 12);

            table2.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
            table2.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;

            //---------------------Adding text after table in page two------------//

            _Word.Paragraph text4;

            text4 = oDoc.Content.Paragraphs.Add();
            text4.Range.Text = "Suderinta: darbuotoju atstove Vita Kazlauskaite";
            text4.Range.Font.Size = 12;
            text4.Format.SpaceBefore = 24;
            text4.Format.SpaceAfter = 6;
            text4.Range.InsertParagraphAfter();

            _Word.Paragraph text5;

            text5 = oDoc.Content.Paragraphs.Add();
            text5.Range.Text = "Sudare: direktore Giedre Tarnauskiene";
            text5.Range.Font.Size = 12;
            text5.Format.SpaceBefore = 0;
            text5.Format.SpaceAfter = 6;
            text5.Range.InsertParagraphAfter();


            //---------------Sets the height and width of table cells-----------//
            //-------------Table1---------------------//
            table1.Columns[1].PreferredWidth = wordApp.CentimetersToPoints(1.0f);
            table1.Columns[2].PreferredWidth = wordApp.CentimetersToPoints(2.5f);
            table1.Columns[3].PreferredWidth = wordApp.CentimetersToPoints(2.5f);
            table1.Columns[4].PreferredWidth = wordApp.CentimetersToPoints(1.6f);
            for (int i = 5; i < 21; i++)
            {
                table1.Columns[i].PreferredWidth = wordApp.CentimetersToPoints(1.15f);
            }

            table1.Rows[1].Height = wordApp.CentimetersToPoints(1.8f);
            for (int i = 2; i <= employeesList.Employees.Count + 1; i++)
            {
                table1.Rows[i].Height = wordApp.CentimetersToPoints(1.15f);
            }


            table1.Range.Font.Size = 10;

            //-------------Table2---------------------//
            table2.Columns[1].PreferredWidth = wordApp.CentimetersToPoints(1.0f);
            table2.Columns[2].PreferredWidth = wordApp.CentimetersToPoints(2.5f);
            table2.Columns[3].PreferredWidth = wordApp.CentimetersToPoints(2.5f);
            table2.Columns[4].PreferredWidth = wordApp.CentimetersToPoints(1.6f);
            for (int i = 5; i <= monthsWeekDays.Length - 12; i++)
            {
                table2.Columns[i].PreferredWidth = wordApp.CentimetersToPoints(1.15f);
            }

            table2.Rows[1].Height = wordApp.CentimetersToPoints(1.8f);
            for (int i = 2; i <= employeesList.Employees.Count + 1; i++)
            {
                table2.Rows[i].Height = wordApp.CentimetersToPoints(1.15f);
            }


            table2.Range.Font.Size = 10;

            //-----------------Populates the table----------------------------//

            //----------------First page------------------------------//

            //--------Default values-------------------//
            table1.Cell(1, 1).Range.Text = "Eil. Nr.";
            table2.Cell(1, 1).Range.Text = "Eil. Nr.";
            table1.Cell(1, 2).Range.Text = "Vardas, Pavarde";
            table2.Cell(1, 2).Range.Text = "Vardas, Pavarde";
            table1.Cell(1, 3).Range.Text = "Pareigos";
            table2.Cell(1, 3).Range.Text = "Pareigos";
            table1.Cell(1, 4).Range.Text = "Nustat. Darbo val. sk.";
            table2.Cell(1, 4).Range.Text = "Nustat. Darbo val. sk.";
            int d = 5;
            foreach (var day in monthsWeekDays)
            {
                if (d < 21)
                {
                    table1.Cell(1, d).Range.Text = (d - 4).ToString();
                }
                else
                {
                    table2.Cell(1, d - 16).Range.Text = (d - 4).ToString();
                }
                d++;
            }

            //------------Data from EmployeesList-----------------//
            int row = 2;
            foreach (var employee in employeesList.Employees)
            {
                table1.Cell(row, 1).Range.Text = (row - 1).ToString();
                table1.Cell(row, 2).Range.Text = employee.Name;
                table1.Cell(row, 3).Range.Text = employee.Position;
                table2.Cell(row, 1).Range.Text = (row - 1).ToString();
                table2.Cell(row, 2).Range.Text = employee.Name;
                table2.Cell(row, 3).Range.Text = employee.Position;
                string[] employeeSchedule = employee.getMonthSchedule();
                int col = 5;
                foreach (var dayHours in employeeSchedule)
                {
                    if (col <= 20)
                    {
                        table1.Cell(row, col).Range.Text = dayHours;
                    }
                    else
                    {
                        table2.Cell(row, col - 16).Range.Text = dayHours;
                    }
                    col++;
                }
                row++;
            }


            //----------------Saves the document and opens it-----------------//
            wordApp.ActiveDocument.SaveAs2("Grafikas_" + MonthNameInLT());
            wordApp.Visible = true;
            oDoc.Activate();
            oDoc.ActiveWindow.View.ShowParagraphs = false;
        }

        public string MonthNameInLT()
        {
            MonthDays month = new MonthDays();
            string[] namesInLT = new string[]
            {   
                "Sausio",
                "Vasario",
                "Kovo",
                "Balandzio",
                "Geguzes",
                "Birzelio",
                "Liepos",
                "Rugpjucio",
                "Rugsejo",
                "Spalio",
                "Lapkricio",
                "Gruodzio"
            };
            return namesInLT[month.nextMonth.Month - 1];

        }
    }
}
