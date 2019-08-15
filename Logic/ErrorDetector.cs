using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace Midi_Analyzer.Logic
{
    class ErrorDetector
    {

        private readonly int FROZEN_ROWS = 10;

        /// <summary>
        /// A basic error detection algorithm. It checks if all the notes in the playthrough are correct. Should an error be detected,
        /// it restarts from the end of the sheet back towards the top. This way, as many correct notes as possible can be detected.
        /// </summary>
        /// <param name="midiWb">The workbook to scan for errors.</param>
        /// <param name="excerptWb">Th workbook containing the excerpt.</param>
        /// <returns></returns>
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

        /// <summary>
        /// Checks if the note in the midisheet match the excerpt sheet.
        /// </summary>
        /// <param name="midiSheet">The sheet of the sample, representing the playthrough.</param>
        /// <param name="excerptSheet">The excerpt sheet, representing the score.</param>
        /// <returns></returns>
        public bool DetectGoodPlaythrough(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            string header = "";
            int excerptIndex = 2;
            int midiIndex = FROZEN_ROWS + 1;
            while (header != "end_of_file")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                if(header == "note_on_c" && int.Parse(midiSheet.Cells[midiIndex, 8].Text.Trim()) != 0)
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
                        midiSheet.Cells[midiIndex, 14].Value = "ERROR";
                        return DetectGoodPlaythroughReversed(midiSheet, excerptSheet); //error detected
                    }
                }
                else if(header == "note_on_c" && int.Parse(midiSheet.Cells[midiIndex, 8].Text.Trim()) == 0)
                {
                    Console.WriteLine("NOTE_ON VELOCITY 0 DETECTED AT: " + midiIndex);
                }
                midiIndex++;
            }
            return true; //No errors found
        }

        /// <summary>
        /// Checks if the notes in the midisheet match the excerpt sheet, in reverse order.
        /// </summary>
        /// <param name="midiSheet">The sheet of the sample, representing the playthrough.</param>
        /// <param name="excerptSheet">The excerpt sheet, representing the score.</param>
        /// <returns></returns>
        public bool DetectGoodPlaythroughReversed(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            string header = "";
            int excerptIndex = excerptSheet.Dimension.End.Row;
            int midiIndex = midiSheet.Dimension.End.Row;
            while (header != "start_track")
            {
                header = midiSheet.Cells[midiIndex, 4].Text.Trim().ToLower();
                if (header == "note_on_c" && int.Parse(midiSheet.Cells[midiIndex, 8].Text.Trim()) != 0)
                {
                    if (excerptSheet.Cells[excerptIndex, 1].Text.Trim().ToLower() == "end" || excerptSheet.Cells[excerptIndex, 2].Text.Trim().ToLower() == "end")
                    {
                        //excerptIndex = excerptSheet.Dimension.End.Row; //Resets the excerpt, in case the person has multiple attempts on the same track.
                        excerptIndex--;
                    }
                    else if (midiSheet.Cells[midiIndex, 7].Text.Trim().ToLower() == excerptSheet.Cells[excerptIndex, 2].Text.Trim().ToLower())
                    {
                        midiSheet.Cells[midiIndex, 11].Value = excerptSheet.Cells[excerptIndex, 5].Value;
                        midiSheet.Cells[midiIndex, 12].Value = excerptSheet.Cells[excerptIndex, 1].Value;
                        midiSheet.Cells[midiIndex, 13].Value = excerptSheet.Cells[excerptIndex, 4].Value;
                        excerptIndex--;
                        midiIndex--;
                    }
                    else
                    {
                        midiSheet.Cells[midiIndex, 14].Value = "ERROR";
                        return false; //error detected
                    }
                }
                else
                {
                    midiIndex--;
                }
            }
            return true; //No errors found 
        }

        /// <summary>
        /// Convert the entire excel range into an array. Then, scan through the array.
        /// Should an error be detected, you convert 10 items into one big string. You also group 3 notes into another string.
        /// Then, you search for the index of the small group inside the big group. 
        /// </summary>
        /// <param name="midiSheet">The sheet of the sample, representing the playthrough.</param>
        /// <param name="excerptSheet">The excerpt sheet, representing the score.</param>
        /// <returns></returns>
        public bool GroupingDetection(ExcelWorksheet midiSheet, ExcelWorksheet excerptSheet)
        {
            List<Node> midiNotes = GetColumnAsList(midiSheet, 7);
            List<Node> excerptNotes = GetColumnAsList(excerptSheet, 2);

            int midiIndex = 0;
            int excerptIndex = 0;

            while(excerptIndex < excerptNotes.Count && midiIndex < midiNotes.Count)
            {
                //Still in development.
            }
            return false;
        }

        /// <summary>
        /// Gets the specified column as a List.
        /// </summary>
        /// <param name="sheet">The sheet to get the data from.</param>
        /// <param name="col">The column to read.</param>
        /// <returns></returns>
        public List<Node> GetColumnAsList(ExcelWorksheet sheet, int col)
        {
            List<Node> columnList = new List<Node>();

            int lastRow = sheet.Dimension.End.Row;
            int index = 2; //Skip header.

            //Traverse the sheet.
            while(index <= lastRow)
            {
                if(sheet.Cells[index, col].Value != null)   //If the cell is not empty.
                {
                    columnList.Add(new Node(index, sheet.Cells[index, col].Text.Trim().ToUpper()) );  //Add the note to the list.
                }
                index++;
            }
            //Return the list.
            return columnList;
        }

        public class Node
        {
            private int row;        //The row in the sheet where the note was found.
            private string note;    //The note value itself.

            public Node(int row, string note)
            {
                this.row = row;
                this.note = note;
            }

            public int Row
            {
                get { return row; }
                set { row = value; }
            }

            public string Note
            {
                get { return note; }
                set { note = value; }
            }
        }

        //#############################################################################Deprecated and incomplete methods.
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
                    if (midiSheet.Cells[midiIndex, 5].Text != excerptSheet.Cells[excerptIndex, 5].Text) //This is assuming they're in the same format.
                    {
                        //ERROR DETECTED;
                        //Type 1: User pressed wrong key, continued playing as usual.
                        if ((midiIndex + 3 < midiSheet.Dimension.End.Row) && (excerptIndex + 3 < excerptSheet.Dimension.End.Row)) //Does this work despite the column?
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
                            else if ((midiSheet.Cells[midiIndex + 1, 5].Text == excerptSheet.Cells[excerptIndex, 5].Text) &&
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
