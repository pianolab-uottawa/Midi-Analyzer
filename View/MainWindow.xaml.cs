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
using Midi_Analyzer.Logic;

namespace Midi_Analyzer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string sourceFileType;

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
            TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            path.Text = "";
            RadioButton midiButton = (RadioButton)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("midiButton");
            if(midiButton.IsChecked == true)
            {
                sourceFileType = "MIDI";
            }
            else
            {
                sourceFileType = "CSV";
            }
            Console.WriteLine(sourceFileType);
        }

        private void BrowseForFile(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            if(sourceFileType == "MIDI")
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
        private void BrowseForFolder(object sender, RoutedEventArgs e)
        {
            var dlg = new System.Windows.Forms.FolderBrowserDialog();
            System.Windows.Forms.DialogResult result = dlg.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                TextBox path = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
                path.Text = dlg.SelectedPath;
            }
        }
        private void ConvertFile(object sender, RoutedEventArgs e)
        {
            TextBox sPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
            string[] sourceFiles = sPath.Text.Replace("\n", "").Split(';');
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
            TextBox sPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("sourcePath");
            TextBox destPath = (TextBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("destinationPath");
            string[] sourceFiles = sPath.Text.Replace("\n", "").Split(';');
            string destinationFolder = destPath.Text;
            Converter converter = new Converter();
            converter.RunCSVBatchFile(sourceFiles, destinationFolder, false);
            Analyzer analyzer = new Analyzer();
            string xlsPath = analyzer.AnalyzeCSVFiles(sourceFiles, destinationFolder);

            //Populate next tab with data
            ListBox xlsList = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("xlsFileList");
            xlsList.Items.Add(xlsPath);
            TabControl tabControl = (TabControl)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("tabController");
            tabControl.Items.OfType<TabItem>().SingleOrDefault(n => n.Name == "errorDetection").Focus();
        }

        private void OpenFile(object sender, RoutedEventArgs e)
        {
            ListBox xlsList = (ListBox)(((FrameworkElement)sender).Parent as FrameworkElement).FindName("xlsFileList");
            string file = xlsList.SelectedItem.ToString();
            Process.Start(@"" + file);
        }
    }
}
