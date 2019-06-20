using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Midi_Analyzer.Logic
{
    class Grapher
    {

        private IDictionary<int, string> columnAssignment;
        private string imagePath;

        public Grapher(string imagePath)
        {
            columnAssignment = new Dictionary<int, string>();
            InitializeDictionary();
            this.imagePath = imagePath;
        }

        /*
         * This method will compare each individual IOI of each sample with the teacher's IOI for their note. The data points on the graph
         * represent how much the student's note deviated from their teachers.
         * 
         */
        public void CreateTeacherIOIGraph(ExcelPackage analysisPackage, ExcelPackage excerptPackage, int numSamples)
        {
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Teacher Tone Lengthening");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];

            //Create Header
            int columnIndex = 1;
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);
            int seriesIndex = numSamples;
            int modelIOI = 4;
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            int markerIndex = 0;

            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);

            for (int i = numSamples; i > 0; i--)
            {
                //Header writing works no problem.
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                graphSheet.Cells[1, columnIndex].Value = treatedSheet.Name;
                graphSheet.Cells[1, columnIndex + 1].Value = "Line Number";
                graphSheet.Cells[1, columnIndex + 2].Value = "Timestamp";
                graphSheet.Cells[1, columnIndex + 3].Value = "IOI Deviation (%)";
                graphSheet.Cells[1, columnIndex + 4].Value = "Spacing";

                string header = "";
                int treatedIndex = 2;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;

                if (i == seriesIndex) //This is the model
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y" ||
                            treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == ""))
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = treatedSheet.Cells[treatedIndex, 10].Value;
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 6].Value;
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                }
                else       //These are the samples
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y" ||
                            treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == ""))
                        {
                            
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            if (graphSheet.Cells[graphIndex, modelIOI].Value != null)
                            {
                                //graphSheet.Cells[graphIndex, columnIndex + 3].Style.Numberformat.Format = "#0.00%";
                                double ioiDeviation = CalculateIOIDeviation((double)(graphSheet.Cells[graphIndex, modelIOI].Value),
                                                                                                            (double)(treatedSheet.Cells[treatedIndex, 10].Value)); //IOI
                                graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(ioiDeviation, 2);
                            }
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 6].Value;
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                    //Get the last row of the IOI column:
                    int numRows = graphSheet.Dimension.End.Row;

                    string lineNumColLetter = ConvertIndexToLetter(columnIndex + 4);
                    string ioiColLetter = ConvertIndexToLetter(columnIndex + 3);
                    string ioiRange = ioiColLetter + "2:" + ioiColLetter + lastValidRow;
                    string timeRange = lineNumColLetter + "2:" + lineNumColLetter + lastValidRow;
                    var series = graph.Series.Add(graphSheet.Cells[ioiRange], graphSheet.Cells[timeRange]);
                    graph.Series[seriesIndex - i - 1].Header = sheetNames[seriesIndex - i - 1];
                    graph.DisplayBlanksAs = eDisplayBlanksAs.Span;
                    markerIndex = SelectMarker(markerIndex, markerTypes.Length);
                    ((ExcelScatterChartSerie)series).Marker = (eMarkerStyle)markerTypes.GetValue(markerIndex);
                    markerIndex++;
                }
                columnIndex += 6;
            }
            graph.Title.Text = "Teacher Tone Lengthening - "+excerptPackage.File.Name.Split('.')[0];
            graph.SetSize(933, 410);
            graph.SetPosition(2, 0, columnIndex, 0);
            analysisPackage.Save();
        }

        /*
        * This method will compare the IOIs of each note with the mean IOI of the same sample, then graph the deviation.
        * 
        */
        public void CreateIOIGraph(ExcelPackage analysisPackage, ExcelPackage excerptPackage, int numSamples)
        {
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Tone Lengthening");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];

            //Create Header
            int columnIndex = 1;
            int markerIndex = 0;
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);

            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);
            for (int i = 1; i <= numSamples; i++)
            {
                //Header writing works no problem.
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                graphSheet.Cells[1, columnIndex].Value = treatedSheet.Name;
                graphSheet.Cells[1, columnIndex + 1].Value = "Line Number";
                graphSheet.Cells[1, columnIndex + 2].Value = "Timestamp";
                graphSheet.Cells[1, columnIndex + 3].Value = "IOI (Milliseconds)";

                //Variables
                double meanIOI = CalculateMeanIOI(treatedSheet);
                graphSheet.Cells[1, columnIndex + 4].Value = meanIOI;
                string header = "";
                int treatedIndex = 2;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;

                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y"))
                    {

                        graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                        graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                        int lineNumber = int.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                        double ioiDeviation = CalculateMeanIOIDeviation(meanIOI, (double)(treatedSheet.Cells[treatedIndex, 10].Value), 
                                                                        (double)excerptSheet.Cells[lineNumber + 1, 4].Value);
                        graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(ioiDeviation, 2);
                        graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 6].Value;
                        lastValidRow = graphIndex;
                        graphIndex++;
                    }
                    else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                    {
                        graphIndex++;
                    }
                    treatedIndex++;
                }
                //Get the last row of the IOI column:
                int numRows = graphSheet.Dimension.End.Row;

                string lineNumColLetter = ConvertIndexToLetter(columnIndex + 4);
                string ioiColLetter = ConvertIndexToLetter(columnIndex + 3);
                string ioiRange = ioiColLetter + "2:" + ioiColLetter + lastValidRow;
                string timeRange = lineNumColLetter + "2:" + lineNumColLetter + lastValidRow;
                var series = graph.Series.Add(graphSheet.Cells[ioiRange], graphSheet.Cells[timeRange]);
                markerIndex = SelectMarker(markerIndex, markerTypes.Length);
                ((ExcelScatterChartSerie)series).Marker = (eMarkerStyle)markerTypes.GetValue(markerIndex);
                graph.Series[i - 1].Header = sheetNames[i - 1];
                markerIndex++;
                columnIndex += 6;
            }
            graph.Title.Text = "Tone Lengthening - "+excerptPackage.File.Name.Split('.')[0];
            graph.SetSize(933, 410);
            graph.SetPosition(2, 0, columnIndex, 0);
            analysisPackage.Save();
        }

        public void CreateVelocityGraph(ExcelPackage analysisPackage, ExcelPackage excerptPackage, int numSamples)
        {
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Dynamics");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];

            //Create Header
            int columnIndex = 1;
            int markerIndex = 0;
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);
            for (int i = 1; i <= numSamples; i++)
            {
                //Header writing works no problem.
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                graphSheet.Cells[1, columnIndex].Value = treatedSheet.Name;
                graphSheet.Cells[1, columnIndex + 1].Value = "Line Number";
                graphSheet.Cells[1, columnIndex + 2].Value = "Timestamp";
                graphSheet.Cells[1, columnIndex + 3].Value = "Velocity Deviation (%)";

                //Variables
                double meanVel = CalculateMeanVelocity(treatedSheet);
                graphSheet.Cells[1, columnIndex + 4].Value = meanVel;
                string header = "";
                int treatedIndex = 2;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;

                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y"))
                    {

                        graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                        graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                        int lineNumber = int.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                        double velDeviation = CalculateMeanVelDeviation(meanVel, (double)(treatedSheet.Cells[treatedIndex, 8].Value));
                        graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(velDeviation, 2);
                        graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 6].Value;
                        lastValidRow = graphIndex;
                        graphIndex++;
                    }
                    else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                    {
                        graphIndex++;
                    }
                    treatedIndex++;
                }
                //Get the last row of the IOI column:
                int numRows = graphSheet.Dimension.End.Row;

                string lineNumColLetter = ConvertIndexToLetter(columnIndex + 4);
                string velColLetter = ConvertIndexToLetter(columnIndex + 3);
                string velRange = velColLetter + "2:" + velColLetter + lastValidRow;
                string timeRange = lineNumColLetter + "2:" + lineNumColLetter + lastValidRow;
                var series = graph.Series.Add(graphSheet.Cells[velRange], graphSheet.Cells[timeRange]);
                markerIndex = SelectMarker(markerIndex, markerTypes.Length);
                ((ExcelScatterChartSerie)series).Marker = (eMarkerStyle)markerTypes.GetValue(markerIndex);
                markerIndex++;
                Console.WriteLine("SHEET NAME: " + sheetNames[i - 1]);
                graph.Series[i - 1].Header = sheetNames[i - 1];
                columnIndex += 6;
            }
            graph.Title.Text = "Dynamics Graph - "+excerptPackage.File.Name.Split('.')[0];
            graph.SetSize(933, 410);
            graph.SetPosition(2, 0, columnIndex, 0);
            InsertImageIntoSheet(graphSheet, 23, columnIndex);
            analysisPackage.Save();
        }

        public void CreateTeacherVelocityGraph(ExcelPackage analysisPackage, ExcelPackage excerptPackage, int numSamples)
        {
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Teacher Dynamics Graph");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];

            //Create Header
            int columnIndex = 1;
            int markerIndex = 0;
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);
            int seriesIndex = numSamples;
            int modelVelCol = 4;

            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);
            for (int i = numSamples; i > 0; i--)
            {
                //Header writing works no problem.
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                graphSheet.Cells[1, columnIndex].Value = treatedSheet.Name;
                graphSheet.Cells[1, columnIndex + 1].Value = "Line Number";
                graphSheet.Cells[1, columnIndex + 2].Value = "Timestamp";
                graphSheet.Cells[1, columnIndex + 3].Value = "Velocity";

                string header = "";
                int treatedIndex = 2;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;

                if (i == seriesIndex) //This is the model
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y"))
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = treatedSheet.Cells[treatedIndex, 8].Value;
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 6].Value;
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                }
                else       //These are the samples
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y" ||
                            treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == ""))
                        {

                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            if (graphSheet.Cells[graphIndex, modelVelCol].Value != null)
                            {
                                //graphSheet.Cells[graphIndex, columnIndex + 3].Style.Numberformat.Format = "#0.00%";
                                double velDeviation = CalculateVelDeviation((double)(graphSheet.Cells[graphIndex, modelVelCol].Value),
                                                                                                            (double)(treatedSheet.Cells[treatedIndex, 8].Value)); //IOI
                                graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(velDeviation, 2);
                            }
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 6].Value;
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                    //Get the last row of the IOI column:
                    int numRows = graphSheet.Dimension.End.Row;

                    string lineNumColLetter = ConvertIndexToLetter(columnIndex + 4);
                    string velColLetter = ConvertIndexToLetter(columnIndex + 3);
                    string velRange = velColLetter + "2:" + velColLetter + lastValidRow;
                    string timeRange = lineNumColLetter + "2:" + lineNumColLetter + lastValidRow;
                    var series = graph.Series.Add(graphSheet.Cells[velRange], graphSheet.Cells[timeRange]);
                    markerIndex = SelectMarker(markerIndex, markerTypes.Length);
                    ((ExcelScatterChartSerie)series).Marker = (eMarkerStyle)markerTypes.GetValue(markerIndex);
                    markerIndex++;
                    graph.Series[seriesIndex - i - 1].Header = sheetNames[seriesIndex - i - 1];
                }
                columnIndex += 6;
            }
            graph.Title.Text = "Teacher Dynamic Graph - "+excerptPackage.File.Name.Split('.')[0];
            graph.SetSize(933, 410);
            graph.SetPosition(2, 0, columnIndex, 0);
            InsertImageIntoSheet(graphSheet, 23, columnIndex);
            analysisPackage.Save();
        }

        public double CalculateMeanIOIDeviation(double meanIOI, double sampleIOI, double noteLength)
        {
            if (sampleIOI != meanIOI)
            {
                double noteBeats = noteLength * 4; // This scales it back to the correct amount. 
                double deviation = ((sampleIOI - (meanIOI * noteBeats))/(meanIOI * noteBeats)) * 100;
                return deviation;
            }
            else
            {
                return 0;
            }
        }

        public double CalculateIOIDeviation(double modelIOI, double sampleIOI)
        {
            if(sampleIOI != modelIOI)
            {
                double deviation = ((sampleIOI / modelIOI) - 1)*100;
                return deviation;
            }
            else
            {
                return 0;
            }
        }

        /*
         * This method calculates the mean IOI of a given worksheet (The worksheet MUST be in the format generated by the tool)! 
         */
        public double CalculateMeanIOI(ExcelWorksheet analyzedSheet)
        {
            int index = 2;
            string header = "";
            double totalIOI = 0.0;
            while(header != "end_of_file")
            {
                header = analyzedSheet.Cells[index, 4].Text.Trim().ToLower().ToLower();
                if(header == "note_on_c" && analyzedSheet.Cells[index, 11].Text.Trim().ToLower() == "y")
                {
                    totalIOI += (double)(analyzedSheet.Cells[index, 10].Value);
                }
                index++;
            }
            double totalBeats = CalculateTotalBeats(analyzedSheet);
            return totalIOI / totalBeats;
        }

        public double CalculateMeanVelDeviation(double meanVel, double sampleVel)
        {
            if (sampleVel != meanVel)
            {
                double deviation = ((sampleVel - meanVel) / meanVel) * 100;
                return deviation;
            }
            else
            {
                return 0;
            }
        }

        public double CalculateMeanVelocity(ExcelWorksheet analyzedSheet)
        {
            int index = 2;
            string header = "";
            double totalVel = 0.0;
            int numNotes = 0;
            while (header != "end_of_file")
            {
                header = analyzedSheet.Cells[index, 4].Text.Trim().ToLower().ToLower();
                if (header == "note_on_c" && analyzedSheet.Cells[index, 11].Text.Trim().ToLower() == "y")
                {
                    totalVel += (double)(analyzedSheet.Cells[index, 8].Value);
                    numNotes++;
                }
                index++;
            }
            return totalVel / numNotes;
        }

        public double CalculateVelDeviation(double modelVel, double sampleVel)
        {
            if (sampleVel != modelVel)
            {
                double deviation = ((sampleVel / modelVel) - 1) * 100;
                return deviation;
            }
            else
            {
                return 0;
            }
        }

        public int SelectMarker(int markerIndex, int limit)
        {
            while (markerIndex == 1 || markerIndex == 3 || markerIndex == 4 || markerIndex == 5 || markerIndex == 6 || markerIndex == 10)
            {
                markerIndex++;
            }
            if (markerIndex == limit)
            {
                markerIndex = 0;
            }
            return markerIndex;
        }

        public string[] CreateSeriesNames(ExcelPackage package, int numSamples)
        {
            string[] names = new string[numSamples];
            string longestName = "";
            ExcelWorksheet sheet = null;
            for(int i = 1; i <= numSamples; i++)
            {
                sheet = package.Workbook.Worksheets[i];
                if(sheet.Name.Length > longestName.Length)
                {
                    longestName = sheet.Name;
                }
            }
            for(int j = 1; j <= numSamples; j++)
            {
                sheet = package.Workbook.Worksheets[j];
                string name = sheet.Name;
                while(name.Length != longestName.Length)
                {
                    name = name + "_";
                }
                if(name.LastIndexOf('_') == name.Length - 1)
                {
                    //name = name + "__";
                }
                names[j - 1] = name;
            }
            return names;
        }

        public void InsertImageIntoSheet(ExcelWorksheet sheet, int row, int col)
        {
            Image image = Image.FromFile(imagePath);
            var scorePicture = sheet.Drawings.AddPicture("Score", image);
            scorePicture.SetPosition(row, 0, col, 0);
            int horizontalCoord = scorePicture.
            scorePicture.SetSize(785, 71);
        }

        public int CalculateMeanIOIV2(ExcelWorksheet sheet)
        {
            double totalBeats = CalculateTotalBeats(sheet);
            int index = sheet.Dimension.End.Row;
            int firstY = -1;             //Represents the row on which we detected a Y.
            int lastN = -1;              //Represents the row on which we detected an N.
            string noteOffLetter = "";
            int noteOffY = -1;
            string header = "";
            while(firstY == -1 && index != 0)
            {
                header = sheet.Cells[index, 4].Text.Trim().ToLower();
                if(header == "note_on_c")
                {
                    if(sheet.Cells[index, 11].Text.Trim().ToLower() == "y")
                    {
                        firstY = index;
                        if(sheet.Cells[index, 7].Text.Trim().ToLower() == sheet.Cells[noteOffY, 7].Text.Trim().ToLower())
                        {
                            //Here, you're supposed to match both the notes. The problem is there may have been multiple note_off_c's,
                            //so if they don't match, you'd have to retroactively search for the matching note_off. 
                        }
                        break;
                    }
                    else if(sheet.Cells[index, 11].Text.Trim().ToLower() == "n")
                    {
                        lastN = index;
                    }
                }
                else if (header == "note_off_c")
                {
                    noteOffY = index;
                }
                index--;
            }
            if(lastN == -1)
            {

            }
            return -1;
        }

        public double CalculateTotalBeats(ExcelWorksheet analyzedSheet)
        {
            double totalBeats = 0.0;
            int index = 2;
            string header = analyzedSheet.Cells[index, 4].Text.Trim().ToLower();
            while (header != "end_of_file")
            {
                header = analyzedSheet.Cells[index, 4].Text.Trim().ToLower();
                if(analyzedSheet.Cells[index, 11].Text.Trim().ToLower() == "y")
                {
                    totalBeats += (double)(analyzedSheet.Cells[index, 13].Value);
                }
                index++;
            }
            return totalBeats * 4; //This is to scale it from quarter notes.
        }

        public string ConvertIndexToLetter(int index)
        {
            if(index < 27)  //There is only one leader in the column ID.
            {
                return columnAssignment[index];
            }
            else
            {
                int firstLetter = index / 26;
                int secondLetter = index % 26;
                string word = columnAssignment[firstLetter] + columnAssignment[secondLetter];
                return word;
            }
        }

        public void InitializeDictionary()
        {
            columnAssignment.Add(1, "A");
            columnAssignment.Add(2, "B");
            columnAssignment.Add(3, "C");
            columnAssignment.Add(4, "D");
            columnAssignment.Add(5, "E");
            columnAssignment.Add(6, "F");
            columnAssignment.Add(7, "G");
            columnAssignment.Add(8, "H");
            columnAssignment.Add(9, "I");
            columnAssignment.Add(10, "J");
            columnAssignment.Add(11, "K");
            columnAssignment.Add(12, "L");
            columnAssignment.Add(13, "M");
            columnAssignment.Add(14, "N");
            columnAssignment.Add(15, "O");
            columnAssignment.Add(16, "P");
            columnAssignment.Add(17, "Q");
            columnAssignment.Add(18, "R");
            columnAssignment.Add(19, "S");
            columnAssignment.Add(20, "T");
            columnAssignment.Add(21, "U");
            columnAssignment.Add(22, "V");
            columnAssignment.Add(23, "W");
            columnAssignment.Add(24, "X");
            columnAssignment.Add(25, "Y");
            columnAssignment.Add(26, "Z");
        }
    }
}
