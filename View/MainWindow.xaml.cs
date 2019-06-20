using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.WindowsAPICodePack.Dialogs;
using Midi_Analyzer.Logic;

namespace Midi_Analyzer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string sourceFileType;
        private Analyzer analyzer;

        public MainWindow()
        {
            InitializeComponent();
            sourceFileType = "MIDI";
        }

        private void buttonCheckChange(object sender, RoutedEventArgs e)
        {
            /*
             * This method is meant to clear the contents of the source path and array should the user pick a different file type.
             * It also checks which radio button is now checked, and assigns that to the sourceFileType variable.
             * */
            //TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            ListBox path = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            path.Items.Clear();
            RadioButton midiButton = (RadioButton)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("midiButton");
            if(midiButton.IsChecked == true)
            {
                sourceFileType = "MIDI";
            }
            else
            {
                sourceFileType = "CSV";
            }
        }

        private void PopulateSourceListbox(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            if (sourceFileType == "MIDI")
            {
                dlg.Filter = "MIDI files|*.MID;*.MIDI";
            }
            else
            {
                dlg.DefaultExt = ".csv";
                dlg.Filter = "CSV Files (*.csv)|*.csv";
            }
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true && dlg.FileNames.Length != 0)
            {
                ListBox sourcePathBox = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
                foreach(string file in dlg.FileNames)
                {
                    sourcePathBox.Items.Add(file);
                }
            }
        }

        private void BrowseForFile(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            if (sourceFileType == "MIDI")
            {
                dlg.Filter = "MIDI files|*.MID;*.MIDI";
            }
            else
            {
                dlg.DefaultExt = ".csv";
                dlg.Filter = "CSV Files (*.csv)|*.csv";
            }
            Nullable<bool> result = dlg.ShowDialog();
            if(result == true && dlg.FileNames.Length != 0)
            {
                TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
                path.Text = String.Join(";\n", dlg.FileNames);
            }
        }

        private void BrowseForMidi(object sender, RoutedEventArgs e)
        {
            /*
             * Opens a dialog that can only select .mid files.
             */
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = false;
            dlg.Filter = "MIDI files|*.MID;*.MIDI";
            Nullable<bool> result = dlg.ShowDialog();
            if(result == true && dlg.FileNames.Length != 0)
            {
                TextBox modelBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("modelBox");
                modelBox.Text = dlg.FileName;
            }
        }

        private void BrowseForCSV(object sender, RoutedEventArgs e)
        {
            /*
             * Opens a dialog that can only select .csv files. 
             */
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.DefaultExt = ".csv";
            dlg.Filter = "XLSX Files (*.xlsx)|*.xlsx";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true && dlg.FileNames.Length != 0)
            {
                TextBox excerptBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("excerptBox");
                excerptBox.Text = dlg.FileName;
            }
        }

        private void BrowseForImage(object sender, RoutedEventArgs e)
        {
            /*
             * Opens a dialog that can only select image files. 
             */
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Image Files |*.jpg;*.jpeg;*.png;*.bmp";
            Nullable<bool> result = dlg.ShowDialog();
            if (result == true && dlg.FileNames.Length != 0)
            {
                TextBox imageBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("imageBox");
                imageBox.Text = dlg.FileName;
            }
        }

        private void BrowseForFolder(object sender, RoutedEventArgs e)
        {
            var dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true;
            CommonFileDialogResult result = dialog.ShowDialog();
            //var dlg = new System.Windows.Forms.FolderBrowserDialog();
            //System.Windows.Forms.DialogResult result = dlg.ShowDialog();
            if (result == CommonFileDialogResult.Ok)
            {
                TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
                path.Text = dialog.FileName;
            }
        }
        private void ConvertFile(object sender, RoutedEventArgs e)
        {
            ListBox sPath = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
            string[] sourceFiles = new string[sPath.Items.Count];
            sPath.Items.CopyTo(sourceFiles, 0);
            string destinationFolder = destPath.Text;
            Converter converter = new Converter();
            if(sourceFileType == "CSV")
            {
                Console.WriteLine("Running conversion to MIDI...");
                converter.RunMIDIBatchFile(sourceFiles, destinationFolder);
            }
            else if(sourceFileType == "MIDI")
            {
                Console.WriteLine("Running conversion to CSV...");
                converter.RunCSVBatchFile(sourceFiles, destinationFolder);
            }
            else
            {
                Console.WriteLine("There was an error with the source file type selection.");
            }
        }

        private void AnalyzeFile(object sender, RoutedEventArgs e)
        {
            ListBox sPath = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");

            //added:
            TextBox excerptBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("excerptBox");
            TextBox modelBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("modelBox");
            TextBox imageBox = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("imageBox");
            string excerptCSV = excerptBox.Text;
            string modelMidi = modelBox.Text;
            string image = imageBox.Text;

            string[] sourceFiles = new string[sPath.Items.Count + 1];
            sPath.Items.CopyTo(sourceFiles, 0);
            sourceFiles[sourceFiles.Length - 1] = modelMidi;
            foreach (string file in sourceFiles)
            {
                Console.WriteLine("FILE NAME: " + file);
            }
            string destinationFolder = destPath.Text;
            Converter converter = new Converter();
            converter.RunCSVBatchFile(sourceFiles, destinationFolder, false);
            analyzer = new Analyzer(sourceFiles, destinationFolder, excerptCSV, modelMidi, image);
            List<string> badSheets = analyzer.AnalyzeCSVFilesStep1();

            //Populate next tab with data
            ListBox xlsList = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("xlsFileList");
            xlsList.Items.Clear();
            foreach(string name in badSheets)
            {
                xlsList.Items.Add(name);
            }
            TabControl tabControl = (TabControl)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("tabController");
            tabControl.Items.OfType<TabItem>().SingleOrDefault(n => n.Name == "errorDetection").Focus();
        }

        private void OpenFile(object sender, MouseButtonEventArgs e)
        {
            Console.WriteLine("SENDER TYPE: " + sender.ToString());
            var list = sender as ListBoxItem;
            TextBox destPath = this.destinationPath;
            //TextBox destPath = FindResource("destinationPath") as TextBox;
            //TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent.Parent as FrameworkElement).FindName("destinationPath");
            string file = destPath.Text + "//analyzedFile.xlsx";

            Process.Start(@"" + file);
        }

        private void GenerateGraphs(object sender, RoutedEventArgs e)
        {
            analyzer.AnalyzeCSVFilesStep2();
            TabControl tabControl = (TabControl)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("tabController");
            tabControl.Items.OfType<TabItem>().SingleOrDefault(n => n.Name == "results").Focus();
        }

        private void DeleteItem(object sender, System.Windows.Input.KeyEventArgs e)
        {
            ListBox sourcePath = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            if(e.Key.Equals(Key.Delete) || e.Key.Equals(Key.Back))
            {
                if(sourcePath.SelectedItems.Count != 0)
                {
                    var selectedItems = sourcePath.SelectedItems;
                    for (int i = selectedItems.Count - 1; i > -1; i--)
                    {
                        sourcePath.Items.Remove(selectedItems[i]);
                    }
                }
            }
        }
    }
}
