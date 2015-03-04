using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using PresentationML = DocumentFormat.OpenXml.Presentation;
using DrawingML = DocumentFormat.OpenXml.Drawing;

namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            var watch = Stopwatch.StartNew();

            string originalPath = @"C:\Users\ex1\downloads\Imi.pptx";
            string path = @"C:\Users\ex1\desktop\randomStuffNotANYHAVEBOfDY12723489";

            //Kopierar filen
            File.Copy(originalPath, path + ".pptx");

            //Change to .zip
            path.Replace(Path.GetExtension(path + ".pptx"), ".zip");

            FileInfo f1 = new FileInfo(path + ".pptx");
            f1.MoveTo(Path.ChangeExtension(path, ".zip"));

            using (ZipArchive zip = ZipFile.Open(path + ".zip", ZipArchiveMode.Update))
            {
                for (int i = zip.Entries.Count - 1; i >= 0; i--)
                {
                    ZipArchiveEntry e = zip.Entries[i];

                    if (e.Name == "app.xml" || e.Name == "core.xml")
                    {
                        //e.Delete();
                    }
                }
            }

            //Change back to .pptx
            path.Replace(Path.GetExtension(path + ".zip"), ".pptx");

            FileInfo f2 = new FileInfo(path + ".zip");
            f2.MoveTo(Path.ChangeExtension(path, ".pptx"));

            OpenXMLReader reader = new OpenXMLReader(path + ".pptx");

            reader.read();

            reader.PresentationObject.getXMLTree().Save(@"C:\Users\ex1\Desktop\out.xml");

            //File.Delete(path + ".pptx");


            Console.WriteLine("\nCompilation time: " + (double)watch.ElapsedMilliseconds/1000 + "s");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
