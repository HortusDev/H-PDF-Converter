using Microsoft.Win32;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SautinSoft.Document;
using static SautinSoft.PdfFocus.CWordOptions;
using static SautinSoft.PdfFocus;
using static System.Collections.Specialized.BitVector32;
using Path = System.IO.Path;

namespace H_PDF_Converter
{
    //[System.Runtime.InteropServices.ComVisible(true)] // if need to merge VBA into c#
    //[System.Runtime.InteropServices.ClassInterface(
    //System.Runtime.InteropServices.ClassInterfaceType.None)]
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string FilePath;
        public MainWindow()
        {
            InitializeComponent();
        }

        //private void SelectFileClick(object sender, RoutedEventArgs e)
        //{
        //    MessageBox.Show("Done");
        //}

        
        private void UpdateFilePath(string path) //updating file path
        {
            FilePath = path;
        }
        private void SelectFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            bool? result = openFileDialog.ShowDialog();

            if (result.HasValue && result.Value)
            {
                pathTextBox.Text = openFileDialog.FileName;
                UpdateFilePath(pathTextBox.Text);
            }
        }
        private void ConvertButton_Click(object sender, RoutedEventArgs e)
        {
            if (pathTextBox.Text == String.Empty)
            {
                MessageBox.Show("Select a file");
                return;
            }
        }

        private void ConvertToDoc(object sender, RoutedEventArgs e)
        {
            string inpFile = FilePath;
            string outFile = FilePath + ".docx";

            // Specifying PdfLoadOptions we explicitly set that a loadable document is PDF.
            PdfLoadOptions pdfLO = new PdfLoadOptions()
            {
                // 'false' - means to load vector graphics as is. Don't transform it to raster images.
                RasterizeVectorGraphics = false,

                // The PDF format doesn't have real tables, in fact it's a set of orthogonal graphic lines.
                // In case of 'true' the component will detect and recreate tables from graphic lines.
                DetectTables = false,

                // 'false' - Never load embedded fonts in PDF. Use the fonts with the same name installed at the system or similar by font metrics.
                // 'true' - Always load embedded fonts in PDF.
                PreserveEmbeddedFonts = true
            };

            DocumentCore dc = DocumentCore.Load(inpFile, pdfLO);
            dc.Save(outFile);

            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        private void ConvertToJPEG(object sender, RoutedEventArgs e)
        {
            string inpFile = FilePath;
            string outFile = FilePath + ".jpeg";

            DocumentCore dc = DocumentCore.Load(inpFile);

            // PaginationOptions allow to know, how many pages we have in the document.
            SautinSoft.Document.DocumentPaginator dp = dc.GetPaginator(new PaginatorOptions());

            // Each document page will be saved in its own image format: PNG, JPEG, TIFF with different DPI.
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                dp.Pages[i].Save(outFile, new ImageSaveOptions() { DpiX = 400, DpiY = 800 });
            }
        }

        private void ConvertToPnG(object sender, RoutedEventArgs e)
        {
            string inpFile = FilePath;
            string outFile = FilePath + ".png";

            DocumentCore dc = DocumentCore.Load(inpFile);

            // PaginationOptions allow to know, how many pages we have in the document.
            SautinSoft.Document.DocumentPaginator dp = dc.GetPaginator(new PaginatorOptions());

            // Each document page will be saved in its own image format: PNG, JPEG, TIFF with different DPI.
            for (int i = 0; i < dp.Pages.Count; i++)
            {
                dp.Pages[i].Save(outFile, new ImageSaveOptions() { DpiX = 400, DpiY = 800 });
            }
        }
        private void ConvertToTXT(object sender, RoutedEventArgs e)
        {
            string inpFile = FilePath;
            string outFile = FilePath + "txt";

            DocumentCore dc = DocumentCore.Load(inpFile);
            dc.Save(outFile);
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        } 
    }
}