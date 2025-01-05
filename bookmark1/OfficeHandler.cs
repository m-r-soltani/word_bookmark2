using Aspose.Words;
using WordToPDF;
using System;
using System.Data;
using System.IO;


namespace OfficeHandler
{
    public class WordHandler
    {
        //Libre office shayad behtarin rahe hal bashe
        //bayad barnamasho nasb koni "https://www.libreoffice.org/"
        //age na fek konam bayad docx->html->pdf ba openxml ya ...
        public static void ConvertWordToPdfWithLibreOffice(string inputFilePath, string outputFilePath)
        {
            var process = new System.Diagnostics.Process
            {
                StartInfo = new System.Diagnostics.ProcessStartInfo
                {
                    FileName = "soffice",
                    Arguments = $"--headless --convert-to pdf --outdir \"{System.IO.Path.GetDirectoryName(outputFilePath)}\" \"{inputFilePath}\"",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            process.WaitForExit();
        }

        //working docx to pdf1 slow solution using interop
        public static void ConvertWordToPdf()
        {
            Word2Pdf objWorPdf = new Word2Pdf();

            // Path where the input file is located
            string inputFolderPath = @"C:\mydocs\";  // This is the path of the input file (DOCX)

            // Input file name
            string inputFileName = "f1.docx";

            // Full input file path (combining folder path and file name)
            object inputFilePath = inputFolderPath + "\\" + inputFileName;

            // Get the file extension (e.g., .docx)
            string fileExtension = System.IO.Path.GetExtension(inputFileName);

            // Generate the output file name by replacing the extension with .pdf
            string outputFileName = inputFileName.Replace(fileExtension, ".pdf");

            // Check if the file is a .doc or .docx file
            if (fileExtension == ".doc" || fileExtension == ".docx")
            {
                // Full output file path (combining folder path and new file name)
                object outputFilePath = inputFolderPath + "\\" + outputFileName;

                // Set input and output paths for conversion
                objWorPdf.InputLocation = inputFilePath;
                objWorPdf.OutputLocation = outputFilePath;

                // Perform the conversion
                objWorPdf.Word2PdfCOnversion();
            }
        }

        public static void DocxToPdfApose()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document(@"C:\mydocs\f1.docx");
            doc.Save(@"C:\mydocs\pdf\aaa.pdf", SaveFormat.Pdf);
        }
    }
}
