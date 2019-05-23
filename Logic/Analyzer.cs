using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Midi_Analyzer.Logic
{
    class Analyzer
    {
        private int division = 0;

        public void AnalyzeCSVFiles(string[] sourceFiles, string dest)
        {
            foreach (string file in sourceFiles)
            {
                string[] fileSplit = file.Split('\\');
                string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                string csv_path = (dest + "\\" + sFileName + ".csv");
                CreateIoIRow(csv_path);
            }
        }

        public void CreateIoIRow(string csv_path)
        {
            //Create an array that represents each key on the piano
            Node[] keys = new Node[128+1];
            for(int i = 1; i < 128+1; i++)
            {
                keys[i] = new Node();
            }
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            try
            {
                Excel.Workbook file = excelApp.Workbooks.Open(csv_path, Delimiter:";", Format: Excel.XlFileFormat.xlCSV);
                //               Excel.Workbook file = excelApp.Workbooks.Open(csv_path,               // Filename
                //                   Type.Missing, Type.Missing, Excel.XlFileFormat.xlCSV,   // Format
                //                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, ";",          // Delimiter
                //                   Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                //Excel.Workbook file = new Excel.Workbooks.OpenText(csv_path,
                //                        DataType: Excel.XlTextParsingType.xlDelimited,
                //                        TextQualifier: Excel.XlTextQualifier.xlTextQualifierNone,
                //                        ConsecutiveDelimiter: true,
                //                        Semicolon: true);
                Excel._Worksheet workSheet = (Excel.Worksheet)file.ActiveSheet;
                Excel.Range usedRange = workSheet.UsedRange;
                int previous_Track = 0;
                int current_track = 0;
                string header;
                string index, trackID, headerID, velocityID, noteID, timeID, ioiID;
                int last_note_played = -1;
                Console.WriteLine("SIZE OF USED RANGE: " + (usedRange.Rows.Count + 1).ToString());
                for(int i = 1; i < usedRange.Rows.Count+1; i++)
                {
                    index = i.ToString();
                    trackID = "A" + index;
                    if(usedRange[i, 1].Text != "0")
                    {
                        headerID = "C" + index;
                        header = usedRange[i, 3].Text.ToLower();
                        velocityID = "F" + index;
                        if (header == "note_on_c" && usedRange[i, 6].Text != "0")
                        {
                            noteID = "E" + index;
                            int current_note = Int32.Parse(usedRange[i, 5].Text);
                            if (last_note_played != -1)
                            {
                                //IOI for previous note registered
                                timeID = "B" + index;
                                double end_time = Double.Parse(usedRange[i, 2].Text);
                                double ioi = end_time - keys[last_note_played].On_time;
                                ioiID = "I" + keys[last_note_played].ToString();
                                usedRange.Rows[i, 9].Text = ioi;
                                //Start time and row for next note saved
                                keys[current_note].On_time = end_time;
                                keys[current_note].Row = i;
                                last_note_played = current_note;
                            }
                            else
                            {   //This would be for the first note played.
                                timeID = "B" + index;
                                keys[current_note].On_time = Double.Parse(usedRange[i, 2].Text);
                                keys[current_note].Row = i;
                                last_note_played = current_note;
                            }                            
                        }
                    }
                }
                file.Save();
            }
            catch (Exception e){
                Console.WriteLine(e.StackTrace);
                Console.WriteLine(e.Message);
            }
        }

        //Nested class
        public class Node
        {
            int _row;
            bool _checked;
            double _on_time;

            public Node()
            {
                _row = 0;
                _on_time = 0;
                _checked = false;
            }

            public Node(int row, double on_time)
            {
                this._row = row;
                this._on_time = on_time;
            }

            public void ClearNode()
            {
                _row = 0;
                _on_time = 0;
            }

            public int Row
            {
                get { return _row; }
                set { _row = value; }
            }

            public bool Checked
            {
                get { return _checked; }
                set { _checked = value; }
            }

            public double On_time
            {
                get { return _on_time; }
                set { _on_time = value; }
            }
        }
    }
}
