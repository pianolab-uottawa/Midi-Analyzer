using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Midi_Analyzer.Logic
{
    class ErrorDetector
    {

        public void readFile(string xls_path, string reference_path)
        {
            /*
             * This method is an attempt at error detection. During coding, I realized the problem of potential errors is very broad.
             * 
             */
            FileInfo xlsFile = new FileInfo(xls_path);
            ExcelPackage midiFile = new ExcelPackage(xlsFile);
            FileInfo referenceFile = new FileInfo(reference_path);
            ExcelPackage excerpt = new ExcelPackage(referenceFile);

            //get the first worksheet in the workbook
            ExcelWorksheet midiSheet = midiFile.Workbook.Worksheets[0];
            ExcelWorksheet excerptSheet = excerpt.Workbook.Worksheets[0];

            string header = "";
            string velocity = "";
            int excerptIndex = 2;
            int midiIndex = 2;
            while (header != "end_of_file")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                velocity = midiSheet.Cells[midiIndex, 6].Text;
                if (header == "note_on_c" & velocity != "0")
                {
                    if(midiSheet.Cells[midiIndex, 5].Text != excerptSheet.Cells[excerptIndex, 5].Text) //This is assuming they're in the same format.
                    {
                        //ERROR DETECTED;
                        //Type 1: User pressed wrong key, continued playing as usual.
                        if((midiIndex + 3 < midiSheet.Dimension.End.Row) && (excerptIndex + 3 < excerptSheet.Dimension.End.Row)) //Does this work despite the column?
                        {
                            if ((midiSheet.Cells[midiIndex + 1, 5].Text == excerptSheet.Cells[excerptIndex + 1, 5].Text) &&
                            (midiSheet.Cells[midiIndex + 2, 5].Text == excerptSheet.Cells[excerptIndex + 2, 5].Text) &&
                            (midiSheet.Cells[midiIndex + 3, 5].Text == excerptSheet.Cells[excerptIndex + 3, 5].Text))
                            {
                                midiSheet.Row(midiIndex).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                midiSheet.Row(midiIndex).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                midiIndex++;
                                excerptIndex++;
                            }
                            //Type 2: Pianist presses wrong key, restarts from that key
                            else if((midiSheet.Cells[midiIndex+1, 5].Text == excerptSheet.Cells[excerptIndex, 5].Text) &&
                                (midiSheet.Cells[midiIndex + 2, 5].Text == excerptSheet.Cells[excerptIndex + 1, 5].Text) &&
                                (midiSheet.Cells[midiIndex + 3, 5].Text == excerptSheet.Cells[excerptIndex + 2, 5].Text))
                            {
                                midiSheet.Row(midiIndex).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                midiSheet.Row(midiIndex).Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                midiIndex++;
                            }
                        }
                    }
                }
            }
            excerpt.Save();
        }
    }
}
