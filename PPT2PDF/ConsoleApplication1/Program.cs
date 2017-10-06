using System;
using System.Collections.Generic;
using System.IO;

namespace PptToPdfConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var files = new List<string>()
                            {
                                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "08.ppt"),
                                Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "09.pptx")
                            };


            var outDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Out");
            Directory.CreateDirectory(outDir);
            var pdfConverter = new PdfConverter(outDir);

            foreach (var f in files)
            {
                if (!pdfConverter.Process(f))
                {
                    Console.WriteLine("Failed.");
                }
            }

            Console.WriteLine("Process Completed. Hit enter to Exit.");
            Console.Read();
        }
    }
}
