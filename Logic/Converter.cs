using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Midi_Analyzer.Logic
{
    class Converter
    {
        public bool RunCSVBatchFile(string[] sourceFiles, string dest)
        {
            try
            {
                string fileName = "CsvGenerator.bat";
                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    foreach (string file in sourceFiles)
                    {
                        string[] fileSplit = file.Split('\\');
                        string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];
                        writer.WriteLine("Midicsv \"" + file + "\" \"" + dest +"\\"+ sFileName + ".csv\"");
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

    }
}
