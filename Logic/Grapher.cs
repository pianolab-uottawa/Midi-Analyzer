using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Drawing;

namespace Midi_Analyzer.Logic
{
    /// <summary>
    /// This class contains all the graphing methods necessary to create graphs into a sheet.
    /// </summary>
    class Grapher
    {
        private readonly int FROZEN_ROWS = 10;
        private IDictionary<int, string> columnAssignment;
        private string imagePath;
        private ExcelPackage analysisPackage;
        private ExcelPackage excerptPackage;
        private int numSamples;

        public Grapher(ExcelPackage analysisPackage, ExcelPackage excerptPackage, string imagePath, int numSamples)
        {
            columnAssignment = new Dictionary<int, string>();
            InitializeDictionary();
            this.imagePath = imagePath;
            this.analysisPackage = analysisPackage;
            this.excerptPackage = excerptPackage;
            this.numSamples = numSamples;
        }

        /// <summary>
        /// This method will compare each individual IOI of each sample with the model's IOI for their note. The data points on the graph
        /// represent how much the sample's note deviated from the model's.
        /// </summary>
        public void CreateModelIOIGraph()
        {
            //1. Initialize all the sheets and add the graph to the new sheet.
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Teacher Tone Lengthening");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);

            //2. Initalize indexes.
            int columnIndex = 1;
            int seriesIndex = numSamples;
            int modelIOI = 4;
            int markerIndex = 0;

            //3. Get the series names, and a list of series markers.
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);

            //Get series from each sample's treated sheet.
            for (int i = numSamples; i > 0; i--)
            {
                //Get corresponding sheet and write header. 
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                WriteHeader(graphSheet, treatedSheet.Name, columnIndex, "IOI Deviation (%)");

                //Set variables and indexes for sheet traversal.
                string header = "";
                int treatedIndex = FROZEN_ROWS + 1;   //Skip header
                int graphIndex = 2;                   //Skip header
                int lastValidRow = 2;   //This index is used to only take into account the range of valid notes (prevents ending N's being included).

                if (i == seriesIndex) //This is the model. We don't calculate deviation here.
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        //Note is included
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y" ||
                            treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == ""))
                        {
                            //Write values from treated sheet into graph sheet. 
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = treatedSheet.Cells[treatedIndex, 10].Value;
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;  //Write spacing from excerpt.
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        //Note was played, but excluded.
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                }
                else       //These are the samples. These will be compared to the model found above.
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        //Note is included.
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y" ||
                            treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == ""))
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            if (graphSheet.Cells[graphIndex, modelIOI].Value != null)       //Ensures the data gotten from the treated sheet wasn't empty/null.
                            {
                                //Calculate IOI deviation and assign it.
                                double ioiDeviation = CalculateIOIDeviation((double)(graphSheet.Cells[graphIndex, modelIOI].Value),
                                                                                                            (double)(treatedSheet.Cells[treatedIndex, 10].Value));
                                graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(ioiDeviation, 2);
                            }
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        //Note was played, but it was excluded.
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                    //Create and add series to the graph.
                    CreateAndAddSeries(graphSheet, graph, columnIndex, lastValidRow, markerIndex);
                    markerIndex++;
                }
                columnIndex += 6;
            }
            //Finalize graph and save.
            string title = "Teacher Tone Lengthening - "+excerptPackage.File.Name.Split('.')[0];
            string yLabel = "Deviation of IOI (%)";
            SetGraphProperties(graph, title, columnIndex, yLabel);
            InsertImageIntoSheet(graphSheet, 23, columnIndex + 1);
            analysisPackage.Save();
        }
        
        /// <summary>
        /// This method will compare the IOIs of each note with the mean IOI of the same sample, then graph the deviation.
        /// </summary>
        public void CreateIOIGraph()
        {
            //1. Initialize all the sheets and add the graph to the new sheet.
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Tone Lengthening");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);

            //2. Initalize indexes.
            int columnIndex = 1;
            int markerIndex = 0;

            //3. Get the series names, and a list of series markers.
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));

            //Get series from each sample's treated sheet. 
            for (int i = 1; i <= numSamples; i++)
            {
                //Get corresponding sheet and write header. 
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                WriteHeader(graphSheet, treatedSheet.Name, columnIndex, "IOI (milliseconds)");

                //Calculate mean and set variables and indexes for sheet traversal.
                double meanIOI = CalculateMeanIOI(treatedSheet, excerptSheet);
                graphSheet.Cells[1, columnIndex + 5].Value = meanIOI;
                string header = "";
                int treatedIndex = FROZEN_ROWS + 1;   //Skip header
                int graphIndex = 2;                   //Skip header
                int lastValidRow = 2;   //This index is used to only take into account the range of valid notes (prevents ending N's being included).

                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" && 
                        treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y")   //Include note.
                    {
                        int lineNumber = int.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                        if (excerptSheet.Cells[lineNumber + 1, 6].Text.Trim().ToLower() == "y")
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            double ioiDeviation = CalculateMeanIOIDeviation(meanIOI, (double)(treatedSheet.Cells[treatedIndex, 10].Value),
                                                                            (double)excerptSheet.Cells[lineNumber + 1, 4].Value);       //Calculate IOI deviation
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(ioiDeviation, 2);                          //Assign deviation
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;          //Write Spacing from excerpt.
                            lastValidRow = graphIndex;
                        }
                        graphIndex++;
                    }
                    //Note was played, but it was excluded.
                    else if (header == "note_on_c" && 
                        treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                    {
                        graphIndex++;
                    }
                    treatedIndex++;
                }
                //Create the series and add it to the graph.
                CreateAndAddSeries(graphSheet, graph, columnIndex, lastValidRow, markerIndex);
                graph.Series[i - 1].Header = sheetNames[i - 1];
                markerIndex++;
                columnIndex += 7;
            }
            //Finalize graph and save.
            string title = "Tone Lengthening - "+excerptPackage.File.Name.Split('.')[0];
            string yLabel = "Deviation of IOI (%)";
            SetGraphProperties(graph, title, columnIndex, yLabel);
            InsertImageIntoSheet(graphSheet, 23, columnIndex + 1);
            analysisPackage.Save();
        }

        /// <summary>
        /// This method will compare the velocities of each note with the mean velocity of the same sample, then graph the deviation.
        /// </summary>
        public void CreateVelocityGraph()
        {
            //1. Initialize all the sheets and add the graph to the new sheet.
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Dynamics");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];
            ExcelChart graph = graphSheet.Drawings.AddChart("MeanVelocityGraph", eChartType.XYScatterLines);

            //1A. This parralel graph will graph the real velocities. This was at Melinas request.
            ExcelChart graph2 = graphSheet.Drawings.AddChart("VelocityGraph", eChartType.XYScatterLines);

            //2. Initalize indexes.
            int columnIndex = 1;
            int markerIndex = 0;

            //3. Get the series names, and a list of series markers.
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));

            //4. Get series from each sample's treated sheet.
            for (int i = 1; i <= numSamples; i++)
            {
                //Get corresponding sheet and write header. 
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                WriteHeader(graphSheet, treatedSheet.Name, columnIndex, "Velocity Deviation (%)");

                //Calculate mean and set variables and indexes for sheet traversal.
                double meanVel = CalculateMeanVelocity(treatedSheet);
                graphSheet.Cells[1, columnIndex + 4].Value = meanVel;
                string header = "";
                int treatedIndex = FROZEN_ROWS + 1;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;   //This index is used to only take into account the range of valid notes (prevents ending N's being included).

                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" && 
                        treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y")   //Include note
                    {
                        int lineNumber = int.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                        if (excerptSheet.Cells[lineNumber + 1, 7].Text.Trim().ToLower() == "y")
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            double velDeviation = CalculateMeanVelDeviation(meanVel, (double)(treatedSheet.Cells[treatedIndex, 8].Value));  //Calculate velocity deviation
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(velDeviation, 2);                              //Assign deviation into sheet
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;              //Write spacing from excerpt.

                            //4A. Thisgraph will show the real velocities.
                            graphSheet.Cells[graphIndex, columnIndex + 5].Value = treatedSheet.Cells[treatedIndex, 8].Value;
                            lastValidRow = graphIndex;
                        }
                        graphIndex++;
                    }
                    //Note was played, but excluded.
                    else if (header == "note_on_c" && 
                        treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                    {
                        graphIndex++;
                    }
                    treatedIndex++;
                }
                //Create and add series to the graph.
                CreateAndAddSeries(graphSheet, graph, columnIndex, lastValidRow, markerIndex);
                graph.Series[i - 1].Header = sheetNames[i - 1];
                SetXAxisMax(graph, 35);

                //Add the velocity series to the non-mean graph.
                CreateAndAddVelSeries(graphSheet, graph2, columnIndex, lastValidRow, markerIndex);
                graph2.Series[i - 1].Header = sheetNames[i - 1];
                SetXAxisMax(graph2, 35);
                
                markerIndex++;
                columnIndex += 7;
            }
            //Finalize graph and save the package.
            string title = "Dynamics - " + excerptPackage.File.Name.Split('.')[0];
            string yLabel = "Deviation of velocity (%)";
            SetGraphProperties(graph, title, columnIndex, yLabel);
            InsertImageIntoSheet(graphSheet, 23, columnIndex + 1);

            //Finalize vel graph
            string title2 = "Dynamics (Raw) - " + excerptPackage.File.Name.Split('.')[0];
            string yLabel2 = "Velocity (in midi units)";
            SetVelGraphProperties(graph2, title2, columnIndex, yLabel2);
            InsertImageIntoSheet(graphSheet, 50, columnIndex + 1);

            analysisPackage.Save();
        }

        /// <summary>
        /// This method will compare each individual velocity of each sample with the model's velocity for their note. The data points on the graph
        /// represent how much the sample's note deviated from the model's.
        /// </summary>
        public void CreateModelVelocityGraph()
        {
            //1. Initialize all the sheets and add the graph to the new sheet.
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Teacher Dynamics Graph");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);

            //2. Initalize indexes.
            int columnIndex = 1;
            int markerIndex = 0;
            int seriesIndex = numSamples;
            int modelVelCol = 4;

            //3. Get the series names, and a list of series markers.
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));

            //Get series from each sample's treated sheet.
            for (int i = numSamples; i > 0; i--)
            {
                //Get corresponding sheet and write header. 
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                WriteHeader(graphSheet, treatedSheet.Name, columnIndex, "Velocity");

                //Set variables and indexes for sheet traversal.
                string header = "";
                int treatedIndex = FROZEN_ROWS + 1;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;   //This index is used to only take into account the range of valid notes (prevents ending N's being included).

                if (i == seriesIndex)   //This is the model
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y"))       //Include note.
                        {
                            //Write model values from treated sheet into graph sheet. 
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = treatedSheet.Cells[treatedIndex, 8].Value;
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;  //Write spacing from excerpt.
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        //Note was played, but excluded.
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                }
                else       //These are the samples. These will be compared to the model found above.
                {
                    while (header != "end_of_file")
                    {
                        header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                        //Note is included.
                        if (header == "note_on_c" && (treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y" ||
                            treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == ""))
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            if (graphSheet.Cells[graphIndex, modelVelCol].Value != null)                //Ensures the data gotten from the treated sheet wasn't empty/null.
                            {
                                //Calculate velocity deviation and assign it.
                                double velDeviation = CalculateVelDeviation((double)(graphSheet.Cells[graphIndex, modelVelCol].Value),
                                                                                                            (double)(treatedSheet.Cells[treatedIndex, 8].Value));
                                graphSheet.Cells[graphIndex, columnIndex + 3].Value = Math.Round(velDeviation, 2);
                            }
                            int lineNumber = Int32.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;      //Write spacing from excerpt.
                            lastValidRow = graphIndex;
                            graphIndex++;
                        }
                        //Note was played, but it was excluded.
                        else if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                        {
                            graphIndex++;
                        }
                        treatedIndex++;
                    }
                    //Create and add series to the graph.
                    CreateAndAddSeries(graphSheet, graph, columnIndex, lastValidRow, markerIndex);
                    markerIndex++;
                    graph.Series[seriesIndex - i - 1].Header = sheetNames[seriesIndex - i - 1];
                }
                columnIndex += 6;
            }
            //Finalize graph and save.
            string title = "Teacher Dynamic Graph - "+excerptPackage.File.Name.Split('.')[0];
            string yLabel = "Deviation of velocity (%)";
            SetGraphProperties(graph, title, columnIndex, yLabel);
            InsertImageIntoSheet(graphSheet, 23, columnIndex + 1);
            analysisPackage.Save();
        }

        /// <summary>
        /// This method will graph the time between each note being played (articulation). No deviation calculation are done here.
        /// </summary>
        public void CreateArticulationGraph()
        {
            //1. Initialize all the sheets and add the graph to the new sheet.
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("Articulation");
            ExcelWorksheet excerptSheet = excerptPackage.Workbook.Worksheets[1];
            ExcelChart graph = graphSheet.Drawings.AddChart("scatterChart", eChartType.XYScatterLines);

            //2. Initalize indexes.
            int columnIndex = 1;
            int markerIndex = 0;
            int seriesIndex = numSamples;
            int modelArtCol = 4;

            //3. Get the series names, and a list of series markers.
            string[] sheetNames = CreateSeriesNames(analysisPackage, numSamples);

            //Get series from each sample's treated sheet.
            for (int i = 1; i <= numSamples; i++)
            {
                //Get corresponding sheet and write header. 
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                WriteHeader(graphSheet, treatedSheet.Name, columnIndex, "Time between notes (ms)");

                //Set variables and indexes for sheet traversal.
                string header = "";
                int treatedIndex = FROZEN_ROWS + 1;   //Skip header
                int graphIndex = 2;     //Skip header
                int lastValidRow = 2;   //This index is used to only take into account the range of valid notes (prevents ending N's being included).

                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" && 
                        treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y")       //Note is included.
                    {
                        int lineNumber = int.Parse(treatedSheet.Cells[treatedIndex, 12].Text);
                        if (excerptSheet.Cells[lineNumber + 1, 8].Text.Trim().ToLower() == "y")
                        {
                            graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                            graphSheet.Cells[graphIndex, columnIndex + 3].Value = treatedSheet.Cells[treatedIndex, 14].Value;   //Assign articulation into sheet
                            graphSheet.Cells[graphIndex, columnIndex + 4].Value = excerptSheet.Cells[lineNumber + 1, 9].Value;  //Write spacing from excerpt.
                            lastValidRow = graphIndex;
                        }
                        graphIndex++;
                    }
                    //Note was played, but excluded.
                    else if (header == "note_on_c" && 
                        treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "n")
                    {
                        graphIndex++;
                    }
                    treatedIndex++;
                }
                //Create and add series to the graph.
                CreateAndAddSeries(graphSheet, graph, columnIndex, lastValidRow, markerIndex);
                markerIndex++;
                graph.Series[i - 1].Header = sheetNames[i - 1];
                columnIndex += 6;
            }
            //Finalize graph and save.
            string title = "Articulation - " + excerptPackage.File.Name.Split('.')[0];
            string yLabel = "Time between notes (ms)";
            SetGraphProperties(graph, title, columnIndex, yLabel);
            InsertImageIntoSheet(graphSheet, 23, columnIndex + 1);
            analysisPackage.Save();
        }

        /// <summary>
        /// Writes the header into the graphsheet.
        /// </summary>
        /// <param name="graphSheet">The excel sheet containing the graph and data.</param>
        /// <param name="sheetName">The name of the sheet (hence, the sample) that represents the current columns.</param>
        /// <param name="columnIndex">Specifies which column we're currently at.</param>
        /// <param name="variableName">The variable that we'll be writing to the column.</param>
        public void WriteHeader(ExcelWorksheet graphSheet, string sheetName, int columnIndex, string variableName)
        {
            graphSheet.Cells[1, columnIndex].Value = sheetName;
            graphSheet.Cells[1, columnIndex + 1].Value = "Line Number";
            graphSheet.Cells[1, columnIndex + 2].Value = "Timestamp";
            graphSheet.Cells[1, columnIndex + 3].Value = variableName;
            graphSheet.Cells[1, columnIndex + 4].Value = "Spacing";
        }

        /// <summary>
        /// Creates a series from a column variable and adds it to the supplied graph.
        /// </summary>
        /// <param name="graphSheet">The excel sheet containing the graph and data.</param>
        /// <param name="graph">The graph to add the series to.</param>
        /// <param name="columnIndex">Specifies which column we're currently at.</param>
        /// <param name="lastValidRow">The last row to include in our series.</param>
        /// <param name="markerIndex">The current marker index to get from the marker types enum.</param>
        public void CreateAndAddSeries(ExcelWorksheet graphSheet, ExcelChart graph, int columnIndex, int lastValidRow, int markerIndex)
        {
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            string lineNumColLetter = ConvertIndexToLetter(columnIndex + 4);
            string varColLetter = ConvertIndexToLetter(columnIndex + 3);
            string varRange = varColLetter + "2:" + varColLetter + lastValidRow;
            string timeRange = lineNumColLetter + "2:" + lineNumColLetter + lastValidRow;
            var series = graph.Series.Add(graphSheet.Cells[varRange], graphSheet.Cells[timeRange]);
            markerIndex = SelectMarker(markerIndex, markerTypes.Length);
            ((ExcelScatterChartSerie)series).Marker = (eMarkerStyle)markerTypes.GetValue(markerIndex);
        }

        /// <summary>
        /// This method creates a series and adds it to the supplied graph. It is intended to be used with the modifications for velocity.
        /// </summary>
        /// <param name="graphSheet"></param>
        /// <param name="graph"></param>
        /// <param name="columnIndex"></param>
        /// <param name="lastValidRow"></param>
        /// <param name="markerIndex"></param>
        public void CreateAndAddVelSeries(ExcelWorksheet graphSheet, ExcelChart graph, int columnIndex, int lastValidRow, int markerIndex)
        {
            Array markerTypes = Enum.GetValues(typeof(eMarkerStyle));
            string lineNumColLetter = ConvertIndexToLetter(columnIndex + 4);
            string varColLetter = ConvertIndexToLetter(columnIndex + 5);        //The velocity variable.
            string varRange = varColLetter + "2:" + varColLetter + lastValidRow;
            string timeRange = lineNumColLetter + "2:" + lineNumColLetter + lastValidRow;
            var series = graph.Series.Add(graphSheet.Cells[varRange], graphSheet.Cells[timeRange]);
            markerIndex = SelectMarker(markerIndex, markerTypes.Length);
            ((ExcelScatterChartSerie)series).Marker = (eMarkerStyle)markerTypes.GetValue(markerIndex);
        }

        /// <summary>
        /// Sets all the graph properties for the velocity graph.
        /// </summary>
        /// <param name="graph"></param>
        /// <param name="title"></param>
        /// <param name="column"></param>
        /// <param name="yLabel"></param>
        public void SetVelGraphProperties(ExcelChart graph, string title, int column, string yLabel)
        {
            graph.Title.Text = title;
            graph.SetSize(958, 410);
            graph.SetPosition(29, 0, column, 0);
            graph.Legend.Position = eLegendPosition.Top;
            graph.XAxis.Fill.Style = eFillStyle.NoFill;
            graph.XAxis.TickLabelPosition = eTickLabelPosition.None;
            graph.XAxis.MajorTickMark = eAxisTickMark.None;
            graph.XAxis.MinorTickMark = eAxisTickMark.None;
            graph.YAxis.Title.Text = yLabel;
            graph.YAxis.Title.Font.Size = 10;
        }

        /// <summary>
        /// Sets all the graph properties. 
        /// </summary>
        /// <param name="graph">The graph to modify the properties of.</param>
        /// <param name="title">The name to assign to the graph.</param>
        /// <param name="column">The column to set the graph at.</param>
        /// <param name="yLabel">The name to assign to the y-axis.</param>
        public void SetGraphProperties(ExcelChart graph, string title, int column, string yLabel)
        {
            graph.Title.Text = title;
            graph.SetSize(994, 410);
            graph.SetPosition(2, 0, column, 0);
            graph.Legend.Position = eLegendPosition.Top;
            graph.XAxis.Fill.Style = eFillStyle.NoFill;
            graph.XAxis.TickLabelPosition = eTickLabelPosition.None;
            graph.XAxis.MajorTickMark = eAxisTickMark.None;
            graph.XAxis.MinorTickMark = eAxisTickMark.None;
            graph.YAxis.Title.Text = yLabel;
            graph.YAxis.Title.Font.Size = 10;
        }

        /// <summary>
        /// Sets the maximum value for the graph's x axis.
        /// </summary>
        /// <param name="graph"></param>
        /// <param name="max"></param>
        public void SetXAxisMax(ExcelChart graph, int max)
        {
            graph.XAxis.MaxValue = max;
        }

        /// <summary>
        /// Calculates the deviation of the sample IOI from the mean IOI.
        /// </summary>
        /// <param name="meanIOI">The mean IOI to compare to.</param>
        /// <param name="sampleIOI">The sample IOI that will be compared.</param>
        /// <param name="noteLength">
        /// The original duration of the note (quarter note, eighth note, etc). Possible values include 0.25, 0.125, 0.0625.
        /// </param>
        /// <returns>
        /// The deviation of the sampleIOI from the meanIOI, in %.
        /// </returns>
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

        /// <summary>
        /// Calculates the deviation of the sampleIOI value from the modelIOI value (no mean here).
        /// </summary>
        /// <param name="modelIOI">The model IOI to compare to.</param>
        /// <param name="sampleIOI">The sample IOI that will be compared.</param>
        /// <returns>
        /// The deviation of the sampleIOI from the modelIOI, in %.
        /// </returns>
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

        /// <summary>
        /// Calculates the mean IOI.
        /// </summary>
        /// <param name="analyzedSheet">The worksheet from which the meanIOI will be calculated.</param>
        /// <returns>
        /// The mean IOI of the sheet, in ms.
        /// </returns>
        public double CalculateMeanIOI(ExcelWorksheet analyzedSheet, ExcelWorksheet excerptSheet)
        {
            int aIndex = 2;
            int eIndex = 2;
            string header = "";
            double totalIOI = 0.0;
            while(header != "end_of_file")
            {
                header = analyzedSheet.Cells[aIndex, 4].Text.Trim().ToLower().ToLower();
                if(header == "note_on_c")
                {
                    if(analyzedSheet.Cells[aIndex, 11].Text.Trim().ToLower() == "y" &&
                        excerptSheet.Cells[eIndex, 6].Text.Trim().ToLower() == "y")
                    {
                        totalIOI += (double)(analyzedSheet.Cells[aIndex, 10].Value);
                    }
                    eIndex++;
                }
                aIndex++;
            }
            double totalBeats = CalculateTotalBeats(analyzedSheet);
            return totalIOI / totalBeats;
        }

        /// <summary>
        /// Calculates the mean velocity deviation of the sample velocity from the mean velocity.
        /// </summary>
        /// <param name="meanVel">The mean velocity to compare to.</param>
        /// <param name="sampleVel">The sample velocity that will be compared.</param>
        /// <returns>
        /// The deviation of the sampleIOI from the meanIOI, in %.
        /// </returns>
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

        /// <summary>
        /// Calculates the mean velocity.
        /// </summary>
        /// <param name="analyzedSheet">The worksheet from which the mean velocity will be calculated.</param>
        /// <returns>
        /// The mean velocity of the sheet.
        /// </returns>
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

        /// <summary>
        /// Calculates the deviation of the sample velocity value from the model velocity value (no mean here).
        /// </summary>
        /// <param name="modelVel">The model velocity to compare to.</param>
        /// <param name="sampleVel">The sample velocity that will be compared.</param>
        /// <returns>
        /// The deviation of the sample velocity from the model velocity, in %.
        /// </returns>
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

        /// <summary>
        /// Used to select a marker for the series. Avoid markers deemed illegible or hard to read. 
        /// </summary>
        /// <param name="markerIndex">The current index the program is at.</param>
        /// <param name="limit">The total number of markers available to use.</param>
        /// <returns>
        /// The current value of markerIndex.
        /// </returns>
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

        /// <summary>
        /// Generates a list of names for the series from the given package. It uses the given number of samples to only read them.
        /// It generates the names by reading each sheet's name.
        /// </summary>
        /// <param name="package">The pakcage to generate the names from.</param>
        /// <param name="numSamples">The number of samples that were generated.</param>
        /// <returns>
        /// An array of strings containing the package names.
        /// </returns>
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

        /// <summary>
        /// Inserts an image from the global variable into the sheet at a specified row and column value.
        /// </summary>
        /// <param name="sheet">The work sheet to insert the image into.</param>
        /// <param name="row">The row where to insert the image.</param>
        /// <param name="col">The column where to insert the image.</param>
        public void InsertImageIntoSheet(ExcelWorksheet sheet, int row, int col)
        {
            sheet.Column(col).Width = 4;               //This is to line up the image with the graph as much as possible.
            Image image = Image.FromFile(imagePath);
            var scorePicture = sheet.Drawings.AddPicture("Score " + row, image);
            scorePicture.SetPosition(row, 0, col, 0);
            //int horizontalCoord = scorePicture.
            scorePicture.SetSize(933, 71);
        }

        /// <summary>
        /// Calculates the total number of beats in the sample.
        /// </summary>
        /// <param name="analyzedSheet">The work sheet to read from and calculate beats.</param>
        /// <returns>
        /// The total number of beats in the sheet.
        /// </returns>
        public double CalculateTotalBeats(ExcelWorksheet analyzedSheet)
        {
            double totalBeats = 0.0;
            int index = 2;      //Skip the header.
            string header = analyzedSheet.Cells[index, 4].Text.Trim().ToLower();
            while (header != "end_of_file")
            {
                header = analyzedSheet.Cells[index, 4].Text.Trim().ToLower();
                if(analyzedSheet.Cells[index, 11].Text.Trim().ToLower() == "y")     //If the note is included, add the beat.
                {
                    totalBeats += (double)(analyzedSheet.Cells[index, 13].Value);
                }
                index++;
            }
            return totalBeats * 4; //This is to scale it from quarter notes.
        }

        /// <summary>
        /// Converts a number index into a letter that would be used for getting excel ranges.
        /// </summary>
        /// <param name="index">The number index of the column.</param>
        /// <returns>
        /// The equivalent letter of the column index as a string.
        /// </returns>
        public string ConvertIndexToLetter(int index)
        {
            if(index < 27)  //There is only one leader in the column ID.
            {
                return columnAssignment[index];
            }
            else
            {
                int firstLetter = index / 26;   //The first letter is the quotient of the division, should the index be larger than 26.
                int secondLetter = index % 26;  //The second letter is the remainder of the division, should the index be larger than 26.
                string word = columnAssignment[firstLetter] + columnAssignment[secondLetter];
                return word;
            }
        }

        /// <summary>
        /// Initializes the global variable dictionary to contain the number and letter equivalencies for the alphabet.
        /// </summary>
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
