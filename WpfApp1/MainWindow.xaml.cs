using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using Microsoft.WindowsAPICodePack.Dialogs;
using OfficeOpenXml;

namespace WpfApp1
{
    /// <summary>
    /// Logique d'interaction pour MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public string ficherNoms;
        public string pathCible;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new CommonOpenFileDialog();
            dlg.Title = "Sélectionner le dossier ou vous voulez supprimer des fichiers";
            dlg.IsFolderPicker = true;
            dlg.InitialDirectory = ".";

            dlg.AddToMostRecentlyUsedList = false;
            dlg.AllowNonFileSystemItems = false;
            dlg.DefaultDirectory = ".";
            dlg.EnsureFileExists = true;
            dlg.EnsurePathExists = true;
            dlg.EnsureReadOnly = false;
            dlg.EnsureValidNames = true;
            dlg.Multiselect = false;
            dlg.ShowPlacesList = true;

            if (dlg.ShowDialog() == CommonFileDialogResult.Ok)
            {
                pathCible = dlg.FileName;
                // Do something with selected folder string
                dossierCible.Text = pathCible;
            }
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx";
            bool? result = dlg.ShowDialog();

            if (result != true) return;

            ficherNoms = dlg.FileName;
            fichierExcel.Text = ficherNoms;
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            
            MessageBoxResult result = MessageBox.Show("voulez vous confirmer ?",
            "Confirmation", MessageBoxButton.YesNoCancel);
            if (result == MessageBoxResult.Yes)
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(@ficherNoms)))
                {
                    var myWorksheet = xlPackage.Workbook.Worksheets.First(); //select sheet here
                    var totalRows = myWorksheet.Dimension.End.Row;
                    var totalColumns = myWorksheet.Dimension.End.Column;

                    var sb = new StringBuilder(); //this is your data
                    var fichier="";
                    for (int rowNum = 1; rowNum <= totalRows; rowNum++) //select starting row here
                    {
                        //var row = myWorksheet.Cells[rowNum, 1, rowNum, totalColumns].Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                        //sb.AppendLine(string.Join(",", row));
                        fichier = myWorksheet.Cells[rowNum, 1].Text + ".pdf";
                        
                        try
                        {
                            // Check if file exists with its full path    
                            if (File.Exists(Path.Combine(pathCible, fichier)))
                            {
                                // If file found, delete it    
                                File.Delete(Path.Combine(pathCible, fichier));
                                listFichers.Items.Add(fichier);
                                // Console.WriteLine("File deleted.");
                            }
                           // else Console.WriteLine("File not found");
                        }
                        catch (IOException ioExp)
                        {
                           // Console.WriteLine(ioExp.Message);
                        }

                    }
                }

            }
        }
    }
}
