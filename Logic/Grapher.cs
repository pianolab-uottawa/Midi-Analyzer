using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Midi_Analyzer.Logic
{
    class Grapher
    {

        private IDictionary<int, string> columnAssignment;

        public Grapher()
        {
            columnAssignment = new Dictionary<int, string>();
            InitializeDictionary();
        }


        public void CreateIOIGraph(ExcelPackage analysisPackage, ExcelPackage excerptPackage)
        {
            ExcelWorksheet treatedSheet = null;
            ExcelWorksheet graphSheet = analysisPackage.Workbook.Worksheets.Add("IOI Graph");

            //Create Header
            int columnIndex = 1;
            ExcelChart graph = graphSheet.Drawings.AddChart("lineChart", eChartType.Line);
            for (int i = 1; i < analysisPackage.Workbook.Worksheets.Count; i++)
            {
                //Header writing works no problem.
                treatedSheet = analysisPackage.Workbook.Worksheets[i];
                graphSheet.Cells[1, columnIndex].Value = treatedSheet.Name;
                graphSheet.Cells[1, columnIndex + 1].Value = "Line Number";
                graphSheet.Cells[1, columnIndex + 2].Value = "Timestamp";
                graphSheet.Cells[1, columnIndex + 3].Value = "IOI (Milliseconds)";

                string header = "";
                int treatedIndex = 2;   //Skip header
                int graphIndex = 2;     //Skip header

                while (header != "end_of_file")
                {
                    header = treatedSheet.Cells[treatedIndex, 4].Text.Trim().ToLower();
                    if (header == "note_on_c" && treatedSheet.Cells[treatedIndex, 11].Text.Trim().ToLower() == "y")
                    {
                        Console.WriteLine("WRITING LINE");
                        graphSheet.Cells[graphIndex, columnIndex + 1].Value = treatedSheet.Cells[treatedIndex, 12].Value;
                        graphSheet.Cells[graphIndex, columnIndex + 2].Value = treatedSheet.Cells[treatedIndex, 3].Value;
                        graphSheet.Cells[graphIndex, columnIndex + 3].Value = treatedSheet.Cells[treatedIndex, 10].Value;
                        graphIndex++;
                    }
                    treatedIndex++;
                }
                //Get the last row of the IOI column:
                int numRows = graphSheet.Dimension.End.Row;

                string lineNumColLetter = ConvertIndexToLetter(columnIndex + 1);
                string ioiColLetter = ConvertIndexToLetter(columnIndex + 3);
                string labelRange = "0:0";
                string ioiRange = ioiColLetter + "2:" + ioiColLetter + (excerptPackage.Workbook.Worksheets[1].Dimension.End.Row - 1).ToString();
                string timeRange = lineNumColLetter + "2:" + lineNumColLetter + (excerptPackage.Workbook.Worksheets[1].Dimension.End.Row - 1).ToString();
                graph.Series.Add(graphSheet.Cells[ioiRange], graphSheet.Cells[timeRange]);
                graph.Series[i-1].Header = treatedSheet.Name;
                columnIndex += 6;
            }
            graph.Title.Text = "IOI Graph";
            graph.SetSize(800, 600);
            graph.SetPosition(2, 0, columnIndex, 0);
            analysisPackage.Save();
        }

        public void CreateVelocityGraph(ExcelPackage analysisPackage)
        {
            
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
