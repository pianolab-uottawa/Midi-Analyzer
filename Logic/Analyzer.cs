using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace Midi_Analyzer.Logic
{
    class Analyzer
    {

        private double tempo = -1;
        private double division = -1;
        private JObject notes;

        public Analyzer()
        {
            var stream = File.OpenText("notes.json");
            string st = stream.ReadToEnd();
            notes = (JObject)JsonConvert.DeserializeObject(st);
        }

        public void AnalyzeCSVFiles(string[] sourceFiles, string dest)
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            foreach (string file in sourceFiles)
            {
                string[] fileSplit = file.Split('\\');
                string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                string csv_path = (dest + "\\" + sFileName + ".csv");
                string xls_path = CreateXLSFile(csv_path);
                CreateTimeRows(xls_path);
                CreateLetterNoteRow(xls_path);
                CreateIOIRowEPP(xls_path);
                CreateGraphs(xls_path);
            }
            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            Console.WriteLine("RunTime " + elapsedTime);
        }

        public void CreateTimeRows(string xls_path)
        {
            /*
             * Creates all columns pertaining to time (Except IOI). Also creates the header (maybe this should be seperated?).
             * 
             */
            FileInfo existingFile = new FileInfo(xls_path);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
                ExcelWorksheet treatedSheet = package.Workbook.Worksheets[2];
                treatedSheet.Cells[1, 3].Value = "Timestamp";
                treatedSheet.Cells[1, 1].Value = "Track Number";
                treatedSheet.Cells[1, 2].Value = "Midi pulses";
                treatedSheet.Cells[1, 4].Value = "Header";
                treatedSheet.Cells[1, 5].Value = "Channel";
                treatedSheet.Cells[1, 6].Value = "Midi Note";
                treatedSheet.Cells[1, 7].Value = "Letter Note";
                treatedSheet.Cells[1, 8].Value = "Velocity";
                treatedSheet.Cells[1, 9].Value = "IOI (pulses)";
                treatedSheet.Cells[1, 10].Value = "IOI (Timestamp)";
                string header = "";
                division = Double.Parse(workSheet.Cells[1, 6].Text);
                int workIndex = 2;
                int treatedIndex = 2;
                while (header != "end_of_file")
                {
                    header = workSheet.Cells[workIndex, 3].Text.Trim().ToLower();
                    if(header == "tempo")
                    {
                        tempo = Double.Parse(workSheet.Cells[workIndex, 4].Text);
                    }
                    if(header == "note_on_c" || header == "note_off_c" || header == "start_track" ||
                        header == "end_track" || header == "end_of_file" || header == "control_c")
                    {
                        double milli = CalculateMilliseconds(Double.Parse(workSheet.Cells[workIndex, 2].Text));
                        string timestamp = ConvertMilliToString(milli);
                        treatedSheet.Cells[treatedIndex, 1].Value = workSheet.Cells[workIndex, 1].Value;
                        treatedSheet.Cells[treatedIndex, 2].Value = workSheet.Cells[workIndex, 2].Value;
                        treatedSheet.Cells[treatedIndex, 3].Value = timestamp;
                        treatedSheet.Cells[treatedIndex, 4].Value = header;
                        treatedSheet.Cells[treatedIndex, 5].Value = workSheet.Cells[workIndex, 4].Value;//Channels
                        treatedSheet.Cells[treatedIndex, 6].Value = workSheet.Cells[workIndex, 5].Value;//note
                        treatedSheet.Cells[treatedIndex, 8].Value = workSheet.Cells[workIndex, 6].Value;//Velocity
                        treatedIndex++;
                    }                    
                    workIndex++;
                }
                package.Save();
            }
        }

        public void CreateLetterNoteRow(string xls_path)
        {
            FileInfo existingFile = new FileInfo(xls_path);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet treatedSheet = package.Workbook.Worksheets[2];

                string header = "";
                int i = 2;
                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[i, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" || header == "note_off_c")
                    {
                        string note = treatedSheet.Cells[i, 6].Text;
                        string newNote = ConvertNumToNote(note);
                        treatedSheet.Cells[i, 7].Value = newNote;
                    }
                    i++;
                }
                package.Save();
            }
            //For this one, I'll play a scale on the piano, and check what each individual note is on the excel file. 
        }

        public void CreateIOIRowEPP(string xls_path)
        {
            /*
             * This method creates an IOI row inside of an xls path. This method only takes max. 3 seconds to run.
             * */
            Node[] keys = new Node[128 + 1];
            for (int i = 1; i < 128 + 1; i++)
            {
                keys[i] = new Node();
            }
            FileInfo existingFile = new FileInfo(xls_path);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet treatedSheet = package.Workbook.Worksheets[2];
                string header = "";
                int last_note_played = -1;
                //Console.WriteLine("SIZE OF USED RANGE: " + (usedRange.Rows.Count + 1).ToString());
                int i = 1;
                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[i, 4].Text.Trim().ToLower();
                    //Console.WriteLine("HEADER NAME: " + header);
                    if (treatedSheet.Cells[i, 1].Text != "0")
                    {
                        //Console.WriteLine("HEADER NAME: " + header);
                        if (header == "note_on_c" && treatedSheet.Cells[i, 8].Text != "0")
                        {
                            //Console.WriteLine("NOTE ON FOUND AT ROW: " + i.ToString());
                            int current_note = Int32.Parse(treatedSheet.Cells[i, 6].Text);
                            if (last_note_played != -1)
                            {
                                //IOI for previous note registered
                                double end_time = Double.Parse(treatedSheet.Cells[i, 2].Text);
                                double ioi = end_time - keys[last_note_played].On_time;
                                double ioi_milli = CalculateMilliseconds(ioi);
                                string ioi_timestamp = ConvertMilliToString(ioi_milli);
                                //Console.WriteLine("CELL I VALUE: "+workSheet.Cel)
                                treatedSheet.Cells[keys[last_note_played].Row, 9].Value = ioi;
                                treatedSheet.Cells[keys[last_note_played].Row, 10].Value = ioi_timestamp;
                                //Start time and row for next note saved
                                keys[current_note].On_time = end_time;
                                keys[current_note].Row = i;
                                last_note_played = current_note;
                            }
                            else
                            {   //This would be for the first note played.
                                keys[current_note].On_time = Double.Parse(treatedSheet.Cells[i, 2].Text);
                                keys[current_note].Row = i;
                                last_note_played = current_note;
                            }
                        }
                    }
                    i++;
                }
                package.Save();
            }
        }

        public string CreateXLSFile(string csv_path)
        {
            string fileName = csv_path.Split('.')[0];
            string newFileName = fileName + ".xlsx";

            var format = new ExcelTextFormat();
            format.Delimiter = ',';
            format.EOL = "\r";
            if (File.Exists(newFileName))
            {
                File.Delete(newFileName);
            }
            using (ExcelPackage package = new ExcelPackage(new FileInfo(newFileName)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Raw Conversion Data");
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csv_path), format, OfficeOpenXml.Table.TableStyles.Medium27, false);
                //worksheet.Cells[1, 9].Value = "IOI Values";
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Treated Data");

                package.Save();
            }
            return newFileName;
        }

        public void CreateGraphs(string xls_path)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(xls_path)))
            {
                ExcelWorksheet treatedSheet = package.Workbook.Worksheets[2];
                ExcelWorksheet graphSheet = package.Workbook.Worksheets.Add("Graphs");
                ExcelChart graph = graphSheet.Drawings.AddChart("lineChart", eChartType.Line);
                graph.Title.Text = "IOI Graph";
                graph.SetPosition(2, 0, 1, 0);
                graph.SetSize(800, 600);
                //Get the last row of the IOI column:
                int numRows = treatedSheet.Dimension.End.Row;
                string labelRange = "0:0";
                string ioiRange = "I2:I100";// + numRows;
                string timeRange = "C2:C100";// + numRows;
                graph.Series.Add(treatedSheet.Cells[ioiRange], treatedSheet.Cells[timeRange]);
                //graph.Series.Add(treatedSheet.Cells[ioiRange], treatedSheet.Cells[labelRange]);
                graph.Series[0].Header = "Time";
                //graph.Series[1].Header = "IOI range";

                package.Save();
            }
        }

        public string ConvertNumToNote(string num)
        {
            return notes[num].Value<string>();
        }

         public double CalculateMilliseconds(double midiPulses)
        {
            double milli = (midiPulses * (tempo/division))/1000;
            //Console.WriteLine("MILLISECONDS: " + milli.ToString());
            return milli;
        }

        public string ConvertMilliToString(double milli)
        {
            TimeSpan t = TimeSpan.FromMilliseconds(milli);
            string answer = string.Format("{0:D2}h:{1:D2}m:{2:D2}s:{3:D3}ms",
                                    t.Hours,
                                    t.Minutes,
                                    t.Seconds,
                                    t.Milliseconds);
            return answer;
        }

        public void CreateIoIRow(string csv_path)
        {
            /*
             * DEPRECATED: This method is significantly slower than its EPPlus counterpart.
             * This method uses the microsoft interop library to write an IOI column into existing csv files. 
             * The main issue is it takes approx. 6 minutes to run.
             */
            Node[] keys = new Node[128+1];
            for(int i = 1; i < 128+1; i++)
            {
                keys[i] = new Node();
            }
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;
            try
            {
                Excel.Workbook file = excelApp.Workbooks.Open(csv_path, Delimiter:";", Format: Excel.XlFileFormat.xlCSV);
                Excel._Worksheet workSheet = (Excel.Worksheet)file.ActiveSheet;
                Excel.Range usedRange = workSheet.UsedRange;
                int previous_Track = 0;
                int current_track = 0;
                string header = "";
                int last_note_played = -1;
                //Console.WriteLine("SIZE OF USED RANGE: " + (usedRange.Rows.Count + 1).ToString());
                int i = 1;
                while(header != "end_of_file")
                {
                    header = workSheet.Cells[i, 3].Text.Trim().ToLower();
                    if (usedRange[i, 1].Text != "0")
                    {
                        //Console.WriteLine("HEADER NAME: " + header);
                        if (header == "note_on_c" && usedRange[i, 6].Text != "0")
                        {
                            //Console.WriteLine("NOTE ON FOUND AT ROW: " + i.ToString());
                            int current_note = Int32.Parse(usedRange[i, 5].Text);
                            if (last_note_played != -1)
                            {
                                //IOI for previous note registered
                                double end_time = Double.Parse(usedRange[i, 2].Text);
                                double ioi = end_time - keys[last_note_played].On_time;
                                //Console.WriteLine("CELL I VALUE: "+workSheet.Cel)
                                workSheet.Cells[keys[last_note_played].Row, 9].Value2 = ioi.ToString();
                                //Start time and row for next note saved
                                keys[current_note].On_time = end_time;
                                keys[current_note].Row = i;
                                last_note_played = current_note;
                            }
                            else
                            {   //This would be for the first note played.
                                keys[current_note].On_time = Double.Parse(usedRange[i, 2].Text);
                                keys[current_note].Row = i;
                                last_note_played = current_note;
                            }                            
                        }
                    }
                    i++;
                }
                file.Save();
                file.Close(0);
                excelApp.Quit();
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
