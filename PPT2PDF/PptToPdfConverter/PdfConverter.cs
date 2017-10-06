using Aspose.Slides;
using System;
using System.IO;

namespace PptToPdfConverter
{
    public class PdfConverter
    {
        private readonly string OutDir;

        public PdfConverter(string outDir)
        {
            this.OutDir = outDir;
        }


        public bool Process(string inFilePath)
        {
            try
            {
                var fi = new FileInfo(inFilePath);
                var fileName = fi.Name;
                var fileBaseName = Path.GetFileNameWithoutExtension(inFilePath);
                var ext = Path.GetExtension(inFilePath);
                var outPdf = Path.Combine(OutDir, $"{fileBaseName}.pdf");

                Console.Write($"Processing {fileName} ... ");

                if (ext != ".ppt" && ext != ".pptx")
                {
                    Console.WriteLine("Invalid filetype.");
                    return false;
                }

                // Load the document
                var ppDoc = new Presentation(inFilePath);
                // Save the document in output format.
                ppDoc.Save(outPdf, Aspose.Slides.Export.SaveFormat.Pdf);

                Console.WriteLine("Completed.");

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }
    }

}
