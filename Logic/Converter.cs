using System;
using System.Diagnostics;
using System.IO;

namespace Midi_Analyzer.Logic
{
    class Converter
    {

        /// <summary>
        /// Runs a batch file that performs the conversion from midi to csv.
        /// </summary>
        /// <param name="sourceFiles">The source midi file paths.</param>
        /// <param name="dest">The destination path to save the new files to.</param>
        /// <param name="openDest">Specifies if the destination path should be opened in windows explorer. Default to on.</param>
        /// <returns>Boolean value that represents succesful conversion.</returns>
        public bool RunCSVBatchFile(string[] sourceFiles, string dest, bool openDest=true)
        {
            try
            {
                string fileName = "CsvGenerator.bat";   //Name of bat file.
                string fullCommand = "";

                //Start writing to the bat file.
                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    //For each source file, add a line specifying the conversion in the document.
                    foreach (string file in sourceFiles)
                    {
                        //Get just the file name, to name the output file.
                        string[] fileSplit = file.Split('\\');
                        string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];

                        //Write the midicsv command into the bat file.
                        fullCommand = "Midicsv \"" + file + "\" \"" + dest + "\\" + sFileName + ".csv\"";
                        writer.WriteLine(fullCommand);
                    }
                    writer.WriteLine("EXIT /B");
                }
                //Runs the bat file.
                Process process = Process.Start(fileName);
                process.WaitForExit();

                //Delete the bat file after its done running.
                File.Delete(fileName);
                if (openDest)
                {
                    //Open windows explorer at the destination.
                    Process.Start(@"" + dest);
                }
                return true;
            }
            //Should an exception be thrown when writing the bat file.
            catch(Exception e){
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }

        /// <summary>
        /// Runs a batch file that performs the conversion from csv to midi.
        /// </summary>
        /// <param name="sourceFiles">The source csv file paths.</param>
        /// <param name="dest">The destination path to save the new files to.</param>
        /// <param name="openDest">Specifies if the destination path should be opened in windows explorer. Default to on.</param>
        /// <returns>Boolean value that represents succesful conversion.</returns>
        public bool RunMIDIBatchFile(string[] sourceFiles, string dest, bool openDest=true)
        {
            try
            {
                string fileName = "MIDIGenerator.bat";  //Name of bat file.

                //Start writing to the bat file.
                using (StreamWriter writer = new StreamWriter(fileName))
                {
                    //For each source file, add a line specifying the conversion in the document.
                    foreach (string file in sourceFiles)
                    {
                        //Get just the file name, to name the output file.
                        string[] fileSplit = file.Split('\\');
                        string sFileName = fileSplit[fileSplit.Length - 1].Split('.')[0];

                        //Write the midicsv command into the bat file.
                        writer.WriteLine("Csvmidi \"" + file + "\" \"" + dest + "\\" + sFileName + ".mid\"");
                    }
                    writer.WriteLine("EXIT /B");
                }

                //Runs the bat file.
                Process process = Process.Start(fileName);
                process.WaitForExit();

                //Delete the bat file after its done running.
                File.Delete(fileName);
                
                if (openDest)
                {
                    //Open windows explorer at the destination.
                    Process.Start(@"" + dest);
                }
                return true;
            }
            //Should an exception be thrown when writing the bat file.
            catch (Exception e)
            {
                Console.WriteLine(e.StackTrace);
                return false;
            }
        }
    }
}
