using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Globalization;
using System.Threading;

namespace Midi_Analyzer.Logic
{
    class Converter
    {
        public bool RunCSVBatchFile(string[] sourceFiles, string dest)
        {
            try
            {
                string fileName = "CsvGenerator.bat";
                string fullCommand = "";
                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    foreach (string file in sourceFiles)
                    {
                        string[] fileSplit = file.Split('\\');
                        string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                        fullCommand = "Midicsv \"" + file + "\" \"" + dest + "\\" + sFileName + ".csv\"";
                        Console.WriteLine("CONVERTER.CS: COMMAND NAME: " + fullCommand);
                        writer.WriteLine(fullCommand);
                    }
                    writer.WriteLine("EXIT /B");
                }
                Process process = Process.Start(fileName);
                process.WaitForExit();
                File.Delete(fileName);
                Process.Start(@"" + dest);
                return true;
            }
            catch(Exception e){
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }
        public bool RunMIDIBatchFile(string[] sourceFiles, string dest)
        {
            try
            {
                string fileName = "MIDIGenerator.bat";
                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    foreach (string file in sourceFiles)
                    {
                        string[] fileSplit = file.Split('\\');
                        string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                        writer.WriteLine("Csvmidi \"" + file + "\" \"" + dest + "\\" + sFileName + ".mid\"");
                    }
                    writer.WriteLine("EXIT /B");
                }
                Process process = Process.Start(fileName);
                process.WaitForExit();
                File.Delete(fileName);
                Process.Start(@"" + dest);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }

        public bool ConvertFilesToXls(string[] sourceFiles, string dest)
        {
            foreach (string file in sourceFiles)
            {
                string[] fileSplit = file.Split('\\');
                string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                string csv_path = (dest + "\\" + sFileName + ".csv");
                //set the formatting options
                ExcelTextFormat format = new ExcelTextFormat();
                format.Delimiter = ';';
                //            format.Culture = new CultureInfo(Thread.CurrentThread.CurrentCulture.ToString());
                //            format.Culture.DateTimeFormat.ShortDatePattern = "dd-mm-yyyy";
                format.Encoding = new UTF8Encoding();

                //read the CSV file from disk
                FileInfo newFile = new FileInfo(csv_path);

                //create a new Excel package
                using (ExcelPackage excelPackage = new ExcelPackage())
                {
                    //create a WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");

                    //load the CSV data into cell A1
                    worksheet.Cells["A1"].LoadFromText(newFile, format);
                    string newFileName = dest + "\\" + sFileName + ".xlsx";
                    excelPackage.SaveAs(new FileInfo(@"c:\workbooks\myworkbook.xlsx"));
                }
            }
            return true;
        }

        public bool ConvertCSVToXls(string sourceFile)
        {
            return false;
        }

    }
}
