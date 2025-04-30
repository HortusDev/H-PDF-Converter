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
    //[System.Runtime.InteropServices.ComVisible(true)] // to merge VBA and c#
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

        private void SelectFileClick(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Действие выполнено");
        }

        //private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{

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

            //switch (conversionDropDown.SelectedIndex)
            //{
            //    //case 0: // Convert PDF to Doc
            //    //    ConvertToDoc(FilePath);
            //    //    break;
            //    //    case 1:
            //    //        // Convert PDFTODOC
            //    //        ConvertPDFtoDoc(pathTextBox.Text);
            //    //        break;
            //    //    case 2:
            //    //        ConvertPNGToPDF(pathTextBox.Text);
            //    //        break;
            //    //default:
            //    //    MessageBox.Show("Select an option");
            //    //    return;
            //}
            
        }

        private void ConvertToDoc(object sender, RoutedEventArgs e)
        {
            string inpFile = FilePath;
            string outFile = @"D:\Файлы\Result228.docx";

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
            string outFile = @"D:\Файлы\Result5.jpeg";

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
            string outFile = @"D:\Файлы\ResultPNG.png";

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
            string outFile = @"D:\Файлы\ResultText.txt";

            DocumentCore dc = DocumentCore.Load(inpFile);
            dc.Save(outFile);
            // Important for Linux: Install MS Fonts
            // sudo apt install ttf-mscorefonts-installer -y

            // Open the result for demonstration purposes.
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outFile) { UseShellExecute = true });
        }
        //    private void ConvertDocToPDF(string docPath)
        //    {
        //        WordDocument wordDocument = new WordDocument(docPath, FormatType.Automatic);
        //        DocToPDFConverter converter = new DocToPDFConverter();
        //        PdfDocument pdfDocument = converter.ConvertToPDF(wordDocument);

        //        string newPDFPath = docPath.Split('.')[0] + ".pdf";
        //        pdfDocument.Save(newPDFPath);

        //        pdfDocument.Close(true);
        //        wordDocument.Close();
        //    }


        //    private void ConvertPNGToPDF(string pngPath)
        //    {
        //        PdfDocument pdfDoc = new PdfDocument();
        //        PdfImage pdfImage = PdfImage.FromStream(new FileStream(pngPath, FileMode.Open));
        //        PdfPage pdfPage = new PdfPage();
        //        PdfSection pdfSection = pdfDoc.Sections.Add();
        //        pdfSection.Pages.Insert(0, pdfPage);
        //        pdfPage.Graphics.DrawImage(pdfImage, 0, 0);

        //        string newPNGPath = pngPath.Split('.')[0] + ".pdf";
        //        pdfDoc.Save(newPNGPath);
        //        pdfDoc.Close(true);
        //    }

        //    private void ConvertPDFtoDoc(string pdfPath)
        //    {
        //        WordDocument wordDocument = new WordDocument();
        //        IWSection section = wordDocument.AddSection();
        //        section.PageSetup.Margins.All = 0;
        //        IWParagraph firstParagraph = section.AddParagraph();

        //        SizeF defaultPageSize = new SizeF(wordDocument.LastSection.PageSetup.PageSize.Width,
        //            wordDocument.LastSection.PageSetup.PageSize.Height);

        //        using (PdfLoadedDocument loadedDocument = new PdfLoadedDocument(pdfPath))
        //        {
        //            for (int i = 0; i < loadedDocument.Pages.Count; i++)
        //            {
        //                using (var image = loadedDocument.ExportAsImage(i, defaultPageSize, false))
        //                {
        //                    IWPicture picture = firstParagraph.AppendPicture(image);
        //                    picture.Width = image.Width;
        //                    picture.Height = image.Height;
        //                }
        //            }
        //        }
        //        ;

        //        string newPDFPath = pdfPath.Split('.')[0] + ".docx";
        //        wordDocument.Save(newPDFPath);
        //        wordDocument.Dispose();

        //    }



        //    private void SelectFile_Click(object sender, RoutedEventArgs e)
        //    {
        //        OpenFileDialog openFileDialog = new OpenFileDialog();
        //        bool? result = openFileDialog.ShowDialog();

        //        if (result.HasValue && result.Value)
        //        {
        //            pathTextBox.Text = openFileDialog.FileName;
        //        }
        //    }

        //    private void OpenFolder(string folderPath)
        //    {
        //        ProcessStartInfo startInfo = new ProcessStartInfo()
        //        {
        //            Arguments = folderPath.Substring(0, folderPath.LastIndexOf('\\')),
        //            FileName = "explorer.exe"
        //        };
        //        Process.Start(startInfo);

        //    }


        //    private void myMainWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        //    {
        //        {
        //            myMainWindow.Width = e.NewSize.Width;
        //            myMainWindow.Height = e.NewSize.Height;

        //            double xChange = 1, yChange = 1;

        //            if (e.PreviousSize.Width != 0)
        //                xChange = (e.NewSize.Width / e.PreviousSize.Width);

        //            if (e.PreviousSize.Height != 0)
        //                yChange = (e.NewSize.Height / e.PreviousSize.Height);

        //            foreach (FrameworkElement fe in myGrid.Children)
        //            {
        //                if (fe is Grid == false)
        //                {
        //                    fe.Height = fe.ActualHeight * yChange;
        //                    fe.Width = fe.ActualWidth * xChange;

        //                    Canvas.SetTop(fe, Canvas.GetTop(fe) * yChange);
        //                    Canvas.SetLeft(fe, Canvas.GetLeft(fe) * xChange);

        //                }
        //            }
        //        }
        //    }
        //}
    }
}