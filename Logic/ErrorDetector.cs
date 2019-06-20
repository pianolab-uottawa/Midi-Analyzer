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
             * DEPRECATED AND INCOMPLETE
             * This method was originally made to try and detect errors caused by the user playing. However, there were too many ways
             * a user could possibly make a mistake.
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

        public List<string> ScanWorkbookForErrors(ExcelPackage midiWb, ExcelPackage excerptWb)
        {
            List<string> badSheets = new List<string>();

            ExcelWorksheet midiSheet = null;
            ExcelWorksheet excerptSheet = excerptWb.Workbook.Worksheets[1];
            for(int i = 1; i <= midiWb.Workbook.Worksheets.Count; i++)
            {
                midiSheet = midiWb.Workbook.Worksheets[i];
                if(!DetectGoodPlaythrough(midiSheet, excerptSheet))
                {
                    badSheets.Add(midiSheet.Name);
                }
            }
            midiWb.Save();
            return badSheets;
        }

        public bool DetectGoodPlaythrough(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            string header = "";
            int excerptIndex = 2;
            int midiIndex = 2;
            while (header != "end_of_file")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                if(header == "note_on_c")
                {
                    if(excerptSheet.Cells[excerptIndex, 1].Text.Trim().ToLower() == "end" || excerptSheet.Cells[excerptIndex, 2].Text.Trim().ToLower() == "end"){
                        excerptIndex = 2; //Resets the excerpt, in case the person has multiple attempts on the same track.
                    }
                    if (midiSheet.Cells[midiIndex, 7].Text.Trim().ToLower() == excerptSheet.Cells[excerptIndex, 2].Text.Trim().ToLower())
                    {
                        midiSheet.Cells[midiIndex, 11].Value = excerptSheet.Cells[excerptIndex, 5].Value;
                        midiSheet.Cells[midiIndex, 12].Value = excerptSheet.Cells[excerptIndex, 1].Value;
                        midiSheet.Cells[midiIndex, 13].Value = excerptSheet.Cells[excerptIndex, 4].Value;
                        excerptIndex++;
                    }
                    else
                    {
                        midiSheet.Cells[midiIndex, 13].Value = "ERROR";
                        return DetectGoodPlaythroughReversed(midiSheet, excerptSheet); //error detected
                    }
                }
                midiIndex++;
            }
            return true; //No errors found
        }

        public bool DetectGoodPlaythroughReversed(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            string header = "";
            int excerptIndex = excerptSheet.Dimension.End.Row;
            int midiIndex = midiSheet.Dimension.End.Row;
            while (header != "start_track")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                if (header == "note_on_c")
                {
                    if (excerptSheet.Cells[excerptIndex, 1].Text.Trim().ToLower() == "end" || excerptSheet.Cells[excerptIndex, 2].Text.Trim().ToLower() == "end")
                    {
                        excerptIndex = excerptSheet.Dimension.End.Row; //Resets the excerpt, in case the person has multiple attempts on the same track.
                    }
                    if (midiSheet.Cells[midiIndex, 7].Text.Trim().ToLower() == excerptSheet.Cells[excerptIndex, 2].Text.Trim().ToLower())
                    {
                        midiSheet.Cells[midiIndex, 11].Value = excerptSheet.Cells[excerptIndex, 5].Value;
                        midiSheet.Cells[midiIndex, 12].Value = excerptSheet.Cells[excerptIndex, 1].Value;
                        midiSheet.Cells[midiIndex, 13].Value = excerptSheet.Cells[excerptIndex, 4].Value;
                        excerptIndex--;
                    }
                    else
                    {
                        midiSheet.Cells[midiIndex, 13].Value = "ERROR";
                        return false; //error detected
                    }
                }
                midiIndex--;
            }
            return true; //No errors found (This technically should not be possible.
        }

        public bool GroupingDetection(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            //Convert the entire excel range into an array. Then, scan through the array.
            //Should an error be detected, you convert 10 items into one big string. You also group 3 notes into another string.
            //Then, you search for the index of the small group inside the big group. big yes.
            return false;
        }
    }
}
