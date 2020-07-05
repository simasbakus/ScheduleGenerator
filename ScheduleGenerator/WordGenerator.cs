using Microsoft.Office.Interop.Word;
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
            try
            {
                //---Creates a document and sets orientation to landscape---//
                var wordApp = new _Word.Application();

                object oMissing = System.Reflection.Missing.Value;

                _Word.Document oDoc = wordApp.Documents.Add(ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                oDoc.PageSetup.Orientation = _Word.WdOrientation.wdOrientLandscape;


                //--------------------Inserts a table------------------------//
                _Word.Range tableLocation = oDoc.Range(0, 0);

                Table table = oDoc.Tables.Add(tableLocation, 9, 20);

                table.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle;
                table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;


                //----------------Saves the document and opens it-----------------//
                wordApp.ActiveDocument.SaveAs2("TestForScheduleGenerator");
                wordApp.Visible = true;
                oDoc.Activate();
            }
            catch (Exception)
            {
                Console.WriteLine("Word document generatio failed!");
            }
            
            
        }
    }
}
