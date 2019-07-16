using OfficeOpenXml;
using System;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace Midi_Analyzer.Logic
{
    class Analyzer
    {

        private readonly int FROZEN_ROWS = 10;

        private double tempo = -1;
        private double division = -1;
        private JObject notes;

        private ExcelPackage analysisPackage;
        private ExcelPackage excerptPackage;

        private string[] sourceFiles;
        private string destinationFolder;
        private string excerptCSV;
        private string modelMidi;
        private string imagePath;

        public Analyzer(string[] sourceFiles, string destinationFolder, string excerptCSV, string modelMidi, string imagePath)
        {
            this.sourceFiles = sourceFiles;
            this.destinationFolder = destinationFolder;
            this.excerptCSV = excerptCSV;
            this.modelMidi = modelMidi;
            this.imagePath = imagePath;

            var stream = File.OpenText("notes.json");
            string st = stream.ReadToEnd();
            notes = (JObject)JsonConvert.DeserializeObject(st);
        }

        /// <summary>
        /// Runs the first part of the analysis on the given files. This includes creating a rawWorkbook.xls, an analyzed file, as well as necessary rows.
        /// </summary>
        public List<string> AnalyzeCSVFilesStep1()
        {
            //Start a stopwatch to see how long the method takes.
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            //Initialize an array to contain the paths to the xls files.
            string[] xlsPaths = new string[sourceFiles.Length];
            string file = "";
            string[] sourceCSVs = new string[sourceFiles.Length];
            string[] fileSplit = null;
            string sFileName = "";

            //Generate the source file names.
            for (int i = 0; i < sourceFiles.Length; i++)
            {
                file = sourceFiles[i];
                fileSplit = file.Split('\\');
                sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                sourceCSVs[i] = (destinationFolder + "\\" + sFileName + ".csv");
            }
            //Generate the raw workbook.
            string sourceFile = CreateCombinedXLSXFile(sourceCSVs, destinationFolder);

            //Create an excel package from the combined xls file.
            ExcelPackage sourcePackage = new ExcelPackage(new FileInfo(sourceFile));

            //Delete the existing analyzed file, if there is one.
            if (File.Exists(destinationFolder + "\\analyzedFile.xlsx"))
            {
                File.Delete(destinationFolder + "\\analyzedFile.xlsx");
            }
            
            //Create the analysis excel package, as well as a package for the excerpt.
            analysisPackage = new ExcelPackage(new FileInfo(destinationFolder + "\\analyzedFile.xlsx"));
            excerptPackage = new ExcelPackage(new FileInfo(excerptCSV));

            //Run Analysis algorithms.
            for (int j = 1; j <= sourcePackage.Workbook.Worksheets.Count; j++)
            {
                //Add a new sheet for the sample and write headers for it.
                ExcelWorksheet sheet = analysisPackage.Workbook.Worksheets.Add(sourcePackage.Workbook.Worksheets[j].Name);
                WriteHeader(sheet);

                CreateTimeRows(sourcePackage);      //Create the time related rows.
                CreateLetterNoteRow();              //Creates a row with the letter equivalencies of the numbered notes from midi
                HighlightNoteRows();                //Highlights the note_on_c's and note_off_c's.
            }
            
            //Scan each of the sheets for possible errors.
            ErrorDetector errorDetector = new ErrorDetector();
            List<string> badFiles = errorDetector.ScanWorkbookForErrors(analysisPackage, excerptPackage);
            
            //Save the analysis Package.
            analysisPackage.Save();

            //Stop the timer.
            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            Console.WriteLine("RunTime " + elapsedTime);

            //Return all the files where mistakes detected.
            return badFiles;
        }

        /// <summary>
        /// Runs the second part of the csv file analysis. This inludes creating an IOI and articulation row, as well as graph creation.
        /// </summary>
        public void AnalyzeCSVFilesStep2()
        {
            //Start a stopwatch to see how long the method takes.
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            //The analysis package must be re-initialized to take into account the changes the user made.
            analysisPackage = new ExcelPackage(new FileInfo(destinationFolder + "\\analyzedFile.xlsx"));

            //Get the rawworkbook again.
            ExcelPackage sourcePackage = new ExcelPackage(new FileInfo(destinationFolder+"\\rawWorkbook.xlsx"));

            //Run the algorithms on each sheet.
            for (int j = 1; j <= sourcePackage.Workbook.Worksheets.Count; j++)
            {
                CreateIOIRowEPP(j);         //Create an IOI row in sheet j.
                CreateArticulationRow(j);   //Create an articualtion row in sheet j.
                CreateDurationRow(j);       //Creates a note duration row in sheet j.
            }

            //Create graphs and save package changes.
            CreateGraphs();
            analysisPackage.Save();

            //Stop the timer.
            stopWatch.Stop();
            // Get the elapsed time as a TimeSpan value.
            TimeSpan ts = stopWatch.Elapsed;

            // Format and display the TimeSpan value.
            string elapsedTime = String.Format("{0:00}:{1:00}:{2:00}.{3:00}",
                ts.Hours, ts.Minutes, ts.Seconds,
                ts.Milliseconds / 10);
            Console.WriteLine("RunTime " + elapsedTime);
        }

        /// <summary>
        /// Highlights all rows that contain either note_on_c or note_off_c as a label.
        /// </summary>
        public void HighlightNoteRows()
        {
            ExcelWorksheet treatedSheet = analysisPackage.Workbook.Worksheets[analysisPackage.Workbook.Worksheets.Count]; //get last sheet, for last file. 
            int i = FROZEN_ROWS + 1; //Skip header row.
            string header = "";
            while(header != "end_of_file")
            {
                header = treatedSheet.Cells[i, 4].Text.Trim().ToLower();
                if (header == "note_on_c")          //If on, color the include cell for the row yellow.
                {
                    treatedSheet.Cells[i, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    treatedSheet.Cells[i, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                }
                else if(header == "note_off_c")     //If off, color the include cell for the row orange.
                {
                    treatedSheet.Cells[i, 11].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    treatedSheet.Cells[i, 11].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                }
                i++;
            }
        }

        /// <summary>
        /// Creates all columns pertaining to time (Except IOI).
        /// </summary>
        /// <param name="sourcePackage">The source excel package to base the creation on.</param>
        public void CreateTimeRows(ExcelPackage sourcePackage)
        {
            //Initialize the needed sheets.
            ExcelWorksheet treatedSheet = analysisPackage.Workbook.Worksheets[analysisPackage.Workbook.Worksheets.Count]; // Get the last sheet
            ExcelWorksheet workSheet = sourcePackage.Workbook.Worksheets[treatedSheet.Name];

            //Save division, used to convert midi pulses into ms.
            division = Double.Parse(workSheet.Cells[1, 6].Text);

            //Initialize values for sheet traversal.
            string header = "";
            int workIndex = 2;
            int treatedIndex = FROZEN_ROWS + 1;

            while (header != "end_of_file")
            {
                header = workSheet.Cells[workIndex, 3].Text.Trim().ToLower();
                if(header == "tempo")                                           //Save tempo, used to convert midi pulses into ms.
                {
                    tempo = Double.Parse(workSheet.Cells[workIndex, 4].Text);
                }
                //All acceptable headers to include in the treated sheet.
                if(header == "note_on_c" || header == "start_track" || header == "note_off_c" ||
                    header == "end_track" || header == "end_of_file" || header == "control_c") 
                {
                    double milli = CalculateMilliseconds(Double.Parse(workSheet.Cells[workIndex, 2].Text)); //Convert midi pulses to milliseconds.
                    string timestamp = ConvertMilliToString(milli);     //Convert milliseconds to timestamp value.
                    treatedSheet.Cells[treatedIndex, 1].Value = workSheet.Cells[workIndex, 1].Value; //Track number.
                    treatedSheet.Cells[treatedIndex, 2].Value = workSheet.Cells[workIndex, 2].Value; //Midi pulses.
                    treatedSheet.Cells[treatedIndex, 3].Value = milli;
                    treatedSheet.Cells[treatedIndex, 4].Value = header;
                    treatedSheet.Cells[treatedIndex, 5].Value = workSheet.Cells[workIndex, 4].Value;//Channels
                    treatedSheet.Cells[treatedIndex, 6].Value = workSheet.Cells[workIndex, 5].Value;//note
                    treatedSheet.Cells[treatedIndex, 8].Value = workSheet.Cells[workIndex, 6].Value;//Velocity
                    treatedIndex++;
                }                    
                workIndex++;
            }
            //Save the changes to the analysis package.
            analysisPackage.Save();
        }

        /// <summary>
        /// Creates a row of letter notes, converted from the midi numbered notes.
        /// </summary>
        public void CreateLetterNoteRow()
        {
            //get the first worksheet in the workbook
            ExcelWorksheet treatedSheet = analysisPackage.Workbook.Worksheets[analysisPackage.Workbook.Worksheets.Count];

            string header = "";
            int i = FROZEN_ROWS + 1;  //Skip the header.
            while (header != "end_of_file")
            {
                header = treatedSheet.Cells[i, 4].Text.Trim().ToLower();
                if (header == "note_on_c" || header == "note_off_c")    //If the header is a note.
                {
                    string note = treatedSheet.Cells[i, 6].Text;
                    string newNote = ConvertNumToNote(note);    //Get the letter note equivalent to the numbered one.
                    treatedSheet.Cells[i, 7].Value = newNote;
                }
                i++;
            }
            //Save the package.
            analysisPackage.Save();
        }

        /// <summary>
        /// Create the IOI row using the EPPlus library.
        /// </summary>
        /// <param name="workSheetIndex"></param>
        public void CreateIOIRowEPP(int workSheetIndex)
        {
            //Initialize an array of notes, used to keep track of when notes were played.
            Node[] keys = new Node[128 + 1];
            for (int i = 1; i < 128 + 1; i++)
            {
                keys[i] = new Node();   //Each index in the array represents a note (hence the array size being 128 (129 - 1).
            }

            //get the specified worksheet from the package and initialize traversal variables.
            ExcelWorksheet treatedSheet = analysisPackage.Workbook.Worksheets[workSheetIndex];
            string header = "";
            int index = FROZEN_ROWS + 1;
            int last_note_played = -1;      //Set the last note played as -1, meaning we're starting the analysis.

            while (header != "end_of_file")
            {
                header = treatedSheet.Cells[index, 4].Text.Trim().ToLower();
                if (treatedSheet.Cells[index, 1].Text != "0")       //Make sure to skip track 0.
                {
                    //If the note is on an the velocity is not 0 (not a note_off).
                    if (header == "note_on_c" && treatedSheet.Cells[index, 8].Text != "0")
                    {
                        int current_note = Int32.Parse(treatedSheet.Cells[index, 6].Text);      //Get the current note number.
                        if (last_note_played != -1)     //If this is not the first note being analyzed.
                        {
                            if (treatedSheet.Cells[keys[last_note_played].Row, 11].Text.Trim().ToLower() == "y")    //Note is included.
                            {
                                //IOI for previous note registered
                                double end_time = Double.Parse(treatedSheet.Cells[index, 2].Text);      //The start time of the current note is technically the end time of the ioi of the previous note.
                                double ioi = end_time - keys[last_note_played].On_time;                 //Calculate the IOI.
                                double ioi_milli = CalculateMilliseconds(ioi);                          //Convert the IOI into milliseconds.
                                string ioi_timestamp = ConvertMilliToString(ioi_milli);                 //Convert the milliseconds into a timestamp value.
                                treatedSheet.Cells[keys[last_note_played].Row, 9].Value = ioi;          //Write the IOI into the sheet.
                                treatedSheet.Cells[keys[last_note_played].Row, 10].Value = ioi_milli;   //Write the millisecond IOI into the sheet.

                                //Start time and row for next note saved
                                keys[current_note].On_time = end_time;
                                keys[current_note].Row = index;
                                last_note_played = current_note; //The current note becomes the last note.
                            }
                            else
                            {   
                                //Reset the count as if it were the first note being played. This is to prevent the excluded note from affecting the calculation.
                                keys[current_note].On_time = Double.Parse(treatedSheet.Cells[index, 2].Text);
                                keys[current_note].Row = index;
                                last_note_played = current_note;
                            }
                        }
                        else
                        {   //This would be for the first note played.
                            keys[current_note].On_time = Double.Parse(treatedSheet.Cells[index, 2].Text);
                            keys[current_note].Row = index;
                            last_note_played = current_note;
                        }
                    }
                }
                index++;
            }
            //Save the package.
            analysisPackage.Save();
        }

        /// <summary>
        /// Creates a row representing the note duration. 
        /// </summary>
        /// <param name="workSheetIndex"></param>
        public void CreateDurationRow(int workSheetIndex)
        {
            //Initialize an array of notes, used to keep track of when notes were played.
            Node[] keys = new Node[128 + 1];
            for (int i = 1; i < 128 + 1; i++)
            {
                keys[i] = new Node();   //Each index in the array represents a note (hence the array size being 128 (129 - 1).
            }

            //get the specified worksheet from the package and initialize traversal variables.
            ExcelWorksheet treatedSheet = analysisPackage.Workbook.Worksheets[workSheetIndex];
            string header = "";
            int index = FROZEN_ROWS + 1;
            int last_note_played = -1;      //Set the last note played as -1, meaning we're starting the analysis.
            int lastLineNumber = GetLastLineNumber();
            int currentLineNumber = -1;
            string lastNoteValue = "";

            //Queue of notes played
            List<Node> queue = new List<Node>();

            while (header != "end_of_file")
            {
                header = treatedSheet.Cells[index, 4].Text.Trim().ToLower();
                if (treatedSheet.Cells[index, 1].Text != "0")       //Make sure to skip track 0.
                {
                    //If the note is on an the velocity is not 0 (not a note_off).
                    if (header == "note_on_c" && treatedSheet.Cells[index, 8].Text != "0")
                    {
                        int current_note = Int32.Parse(treatedSheet.Cells[index, 6].Text);      //Get the current note number.
                        if (treatedSheet.Cells[index, 11].Text.Trim().ToLower() == "y")    //Note is included.
                        {
                            //Start time and row for next note saved
                            keys[current_note].On_time = Double.Parse(treatedSheet.Cells[index, 2].Text);
                            keys[current_note].Row = index;
                            keys[current_note].Checked = false;
                            last_note_played = current_note; //The current note becomes the last note.
                            lastNoteValue = treatedSheet.Cells[index, 7].Text.Trim();
                            queue.Add(new Node(index, current_note));   //Add it to the queue.
                        }
                    }
                    else if(header == "note_off_c" || treatedSheet.Cells[index, 8].Text == "0") //Calculate the previous note's duration.
                    {
                        int current_note = Int32.Parse(treatedSheet.Cells[index, 6].Text);
                        int qIndex = queue.FindIndex(x => x.Note == current_note);
                        if (qIndex != -1)       //Matches the previous note_on.
                        {
                            Node result = queue[qIndex];
                            double endTime = Double.Parse(treatedSheet.Cells[index, 2].Text); //End time of the last note.
                            double noteDuration = endTime - Double.Parse(treatedSheet.Cells[result.Row, 2].Text.Trim());
                            double durationMilli = CalculateMilliseconds(noteDuration);
                            string durationString = ConvertMilliToString(durationMilli);
                            treatedSheet.Cells[result.Row, 15].Value = noteDuration;
                            treatedSheet.Cells[result.Row, 16].Value = durationMilli;
                            keys[last_note_played].Checked = true; //MUDAMUDAMUDAMUDAMUDAMUDA
                            queue.RemoveAt(qIndex);
                        }
                    }
                }
                index++;
            }
            //Save the package.
            analysisPackage.Save();
        }

        /// <summary>
        /// Creates a row representing the articulation (Time between consecutive notes in ms) in the worksheet specified.
        /// </summary>
        /// <param name="workSheetIndex">The worksheet to analyze from the analysis package.</param>
        public void CreateArticulationRow(int workSheetIndex)
        {
            //get the designated workbook from the package.
            ExcelWorksheet treatedSheet = analysisPackage.Workbook.Worksheets[workSheetIndex];

            //Get an ordered list of the notes played in the sheet.
            List<Node> notesPlayed = GetListForArticulation(treatedSheet);
            
            for (int i = 1; i < notesPlayed.Count; i++)
            {
                double difference = notesPlayed[i].On_time - notesPlayed[i - 1].Off_time;   //Calculate the difference between the start of a note and the end of a previous one.
                treatedSheet.Cells[notesPlayed[i-1].Row, 14].Value = CalculateMilliseconds(difference); //Convert the difference into milliseconds and assign it to the sheet.
            }
            //Save the package.
            analysisPackage.Save();
        }

        /// <summary>
        /// Generates a list of Node items that represent the order in which the notes were played.
        /// </summary>
        /// <param name="treatedSheet">The worksheet to generate the row from.</param>
        /// <returns>A list of nodes representing all notes played in order.</returns>
        public List<Node> GetListForArticulation(ExcelWorksheet treatedSheet)
        {
            //Create the list that will contain the nodes.
            List<Node> notesPlayed = new List<Node>();

            //Create sheet traversal variables.
            string header = "";
            int index = FROZEN_ROWS + 1;
            while (header != "end_of_file")
            {
                header = treatedSheet.Cells[index, 4].Text.Trim().ToLower();
                if (header == "note_on_c")  //A note on value.
                {
                    //Assign current row, note and on_time to a new node in the list.
                    int current_note = Int32.Parse(treatedSheet.Cells[index, 6].Text);
                    double on_time = Double.Parse(treatedSheet.Cells[index, 2].Text);
                    notesPlayed.Add(new Node(index, on_time, current_note));
                }
                else if(header == "note_off_c")
                {
                    //Assign the off_time to the note added previously.
                    int current_note = Int32.Parse(treatedSheet.Cells[index, 6].Text);
                    int listIndex = notesPlayed.FindLastIndex(node => node.Note == current_note && node.Off_time == 0);
                    notesPlayed[listIndex].Off_time = Double.Parse(treatedSheet.Cells[index, 2].Text);
                }
                index++;
            }
            //Return the list of all notes played.
            return notesPlayed;
        }

        /// <summary>
        /// Gets the last line number from the excerpt package.
        /// POSSIBLE EDGE CASE!!!! If the user has excluded the last note in the midi sheet, this method fails.
        /// </summary>
        public int GetLastLineNumber()
        {
            ExcelWorksheet eSheet = excerptPackage.Workbook.Worksheets[1];
            int lastRow = eSheet.Dimension.End.Row;
            int lastLineNumber = -1;
            while (lastRow > 1)
            {
                if (eSheet.Cells[lastRow, 1].Value != null && eSheet.Cells[lastRow, 1].Text.Trim() != "" && eSheet.Cells[lastRow, 1].Text.Trim().ToLower() != "end")
                {
                    lastLineNumber = Int32.Parse(eSheet.Cells[lastRow, 1].Text.Trim());
                    break;
                }
                else
                {
                    lastRow--;
                }
            }
            return lastLineNumber;
        }

        /// <summary>
        /// Creates an XLS file from the given path to an existing csv file. 
        /// </summary>
        /// <param name="csv_path">The path to the csv file.</param>
        /// <returns>The name of the new xlsx file (including its path).</returns>
        public string CreateXLSXFile(string csv_path)
        {
            //Get the file name and create a new name for the xlsx file. 
            string fileName = csv_path.Split('.')[0];
            string newFileName = fileName + ".xlsx";

            //Set formatting and delimiter values.
            var format = new ExcelTextFormat();
            format.Delimiter = ',';
            format.EOL = "\r";

            //Check if the new file exists already. If so, delete it.
            if (File.Exists(newFileName))
            {
                File.Delete(newFileName);
            }

            //Create a new package with the new file name.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(newFileName)))
            {
                //Add a new worksheet into the file.
                ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Raw Conversion Data");
                //Copy the data from the csv int othe new sheet.
                worksheet.Cells["A1"].LoadFromText(new FileInfo(csv_path), format, OfficeOpenXml.Table.TableStyles.Medium27, false);
                ExcelWorksheet worksheet2 = package.Workbook.Worksheets.Add("Treated Data");

                //Save the package changes.
                package.Save();
            }
            //return the name of the new file (including its path).
            return newFileName;
        }

        /// <summary>
        /// Creates an xlsx file containing the data from multiple csv files. Each csv has its own sheet in the xlsx file. 
        /// </summary>
        /// <param name="csvPaths">An array of strings, each being a path to a csv file.</param>
        /// <param name="dest">A path to the folder in which the new xlsx file should be saved.</param>
        /// <returns>The name of the new file, including its path.</returns>
        private string CreateCombinedXLSXFile(string[] csvPaths, string dest)
        {
            //Create new file name.
            string newFileName = dest + "\\rawWorkbook.xlsx";

            //If the file already exists, delete it.
            if (File.Exists(newFileName))
            {
                File.Delete(newFileName);
            }
            
            //Set the formatting parameters.
            var format = new ExcelTextFormat();
            format.Delimiter = ',';
            format.EOL = "\r";

            //Create a new excel package.
            using (ExcelPackage package = new ExcelPackage(new FileInfo(newFileName)))
            {
                //Variables used for each sheet.
                ExcelWorksheet worksheet = null;
                string name = "";

                //For each of the csv files, run the copying process.
                for (int i =0; i < csvPaths.Length; i++)
                {
                    //Get the actual name of the file.
                    string[] path = csvPaths[i].Split(new string[] { "\\" }, StringSplitOptions.None);
                    name = path[path.Length - 1].Split('.')[0];

                    //Add the new sheet into the xlsx file and copy the data from the csv into it. 
                    worksheet = package.Workbook.Worksheets.Add(name);
                    worksheet.Cells["A1"].LoadFromText(new FileInfo(csvPaths[i]), format, OfficeOpenXml.Table.TableStyles.Medium27, false);

                    //Delete the CSV file. 
                    File.Delete(csvPaths[i]);
                }
                //Save the new package name.
                package.Save();
            }
            //Return the file name (including its path).
            return newFileName;
        }

        /// <summary>
        /// Initializes an instance of the grapher, and creates the mean IOI, mean Velocity and articulation graph. 
        /// </summary>
        public void CreateGraphs()
        {
            //Get the number of samples first. This is to prevent added graph sheets to ruin the numbering later on.
            int numSamples = analysisPackage.Workbook.Worksheets.Count;
            
            //Initalize grapher.
            Grapher grapher = new Grapher(analysisPackage, excerptPackage, imagePath, numSamples);

            //Create the graphs.
            grapher.CreateIOIGraph();
            grapher.CreateVelocityGraph();
            grapher.CreateArticulationGraph();
            grapher.CreateNoteDurationGraph();
        }

        /// <summary>
        /// Write the header into the specified sheet.
        /// </summary>
        /// <param name="sheet">Specifies the sheet in which to write the header to.</param>
        public void WriteHeader(ExcelWorksheet sheet)
        {
            sheet.View.FreezePanes(FROZEN_ROWS + 1, 14 + 1);
            sheet.Cells[FROZEN_ROWS, 1].Value = "Track Number";
            sheet.Cells[FROZEN_ROWS, 2].Value = "Midi pulses";
            sheet.Cells[FROZEN_ROWS, 3].Value = "Timestamp";
            sheet.Cells[FROZEN_ROWS, 4].Value = "Header";
            sheet.Cells[FROZEN_ROWS, 5].Value = "Channel";
            sheet.Cells[FROZEN_ROWS, 6].Value = "Midi Note";
            sheet.Cells[FROZEN_ROWS, 7].Value = "Letter Note";
            sheet.Cells[FROZEN_ROWS, 8].Value = "Velocity";
            sheet.Cells[FROZEN_ROWS, 9].Value = "IOI (pulses)";
            sheet.Cells[FROZEN_ROWS, 10].Value = "IOI (Milliseconds)";
            sheet.Cells[FROZEN_ROWS, 11].Value = "Include? (Y/N)";
            sheet.Cells[FROZEN_ROWS, 12].Value = "Line Number";
            sheet.Cells[FROZEN_ROWS, 13].Value = "Duration";
            sheet.Cells[FROZEN_ROWS, 14].Value = "Articulation";
            sheet.Cells[FROZEN_ROWS, 15].Value = "Note duration (midi)";
            sheet.Cells[FROZEN_ROWS, 16].Value = "Note duration (ms)";
        }

        /// <summary>
        /// Converts a given note in midi numbering into a letter note.
        /// </summary>
        /// <param name="num">The note to convert.</param>
        /// <returns>A string representing the note as a letter note (e.g. G5).</returns>
        public string ConvertNumToNote(string num)
        {
            return notes[num].Value<string>();
        }

        /// <summary>
        /// Converts the number of midi pulses into milliseconds. 
        /// </summary>
        /// <param name="midiPulses">Number of midipulses.</param>
        /// <returns>Time in milliseconds equivalent to the number of midi pulses.</returns>
         public double CalculateMilliseconds(double midiPulses)
        {
            double milli = (midiPulses * (tempo/division))/1000;    //Formula derived from the midicsv tool website.
            return milli;
        }

        /// <summary>
        /// Converts the milliseconds into a timestamp format.
        /// </summary>
        /// <param name="milli">The number of milliseconds we want to convert.</param>
        /// <returns>A string of the number as a timestamp.</returns>
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

        /// <summary>
        /// This class represents a note that was played. 
        /// </summary>
        public class Node
        {
            int _row;           //The row on which the note is found (can be either on or off, depending on usage).
            int _note;          //The number of the note itself in midi format.
            bool _checked;      //Marks that the note was checked.
            double _on_time;    //The note_on_c time of the note. 
            double _off_time;   //The note_off_c time of the note.

            public Node()
            {
                _note = -1;
                _row = 0;
                _on_time = 0;
                _checked = false;
                _off_time = 0;
            }

            /// <summary>
            /// Generates a node using only the row and on_time information.
            /// </summary>
            /// <param name="row">The row number at which the on time was detected.</param>
            /// <param name="on_time">The on time, which is when the note was pressed down.</param>
            public Node(int row, double on_time)
            {
                _row = row;
                _on_time = on_time;
            }

            /// <summary>
            /// Generates a node using only the row and note information.
            /// </summary>
            /// <param name="row">The row number at which the on time was detected.</param>
            /// <param name="note">The note number in midi format.</param>
            public Node(int row, int note)
            {
                _row = row;
                _note = note;
            }

            /// <summary>
            /// Generates a node using the row, on_time and note information.
            /// </summary>
            /// <param name="row">The row number at which the on_time was detected.</param>
            /// <param name="on_time">The on time, which is when the note was pressed down.</param>
            /// <param name="note">The note number in midi format.</param>
            public Node(int row, double on_time, int note)
            {
                _row = row;
                _on_time = on_time;
                _note = note;
            }

            /// <summary>
            /// Generates a node using the row, on_time, off_time and note information.
            /// </summary>
            /// <param name="row">The row number at which either the on_time or off_time was detected.</param>
            /// <param name="on_time">The on time, which is when the note was pressed down.</param>
            /// <param name="off_time">The off time, which is when the note was released.</param>
            /// <param name="note">The note number in midi format.</param>
            public Node(int row, double on_time, double off_time, int note)
            {
                _row = row;
                _on_time = on_time;
                _off_time = off_time;
                _note = note;
            }

            /// <summary>
            /// Clears all information from the node.
            /// </summary>
            public void ClearNode()
            {
                _row = 0;
                _on_time = 0;
                _note = -1;
                _off_time = 0;
                _checked = false;
            }

            public int Row
            {
                get { return _row; }
                set { _row = value; }
            }

            public int Note
            {
                get { return _note; }
                set { _note = value; }
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

            public double Off_time
            {
                get { return _off_time; }
                set { _off_time = value; }
            }
        }

        //DEPRECATED METHODS#############################################################################################
        [Obsolete("Please use CreateIOIRowEPP instead. CreateIoIRow is deprecated given how slow the excel interop needs to run (approx. 5 mins) " +
            "compared to its Epplus counterpart (2 seconds). Furthermore, it requires that excel be installed on the target machine.")]
        public void CreateIoIRow(string csv_path)
        {
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
                int i = FROZEN_ROWS + 1;
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
    }
}
