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

            string originalPath = @"C:\Users\ex1\downloads\grouprot.pptx";
            
            //Copy file
            string path = @"C:\Users\ex1\desktop\randomStuffNotANYHAVEBOfDY12723489";
            File.Copy(originalPath, path + ".pptx");

            //Change to .zip
            FileInfo f1 = new FileInfo(path + ".pptx");
            f1.MoveTo(Path.ChangeExtension(path, ".zip"));

            //Open zip file
            using (ZipArchive zip = ZipFile.Open(path + ".zip", ZipArchiveMode.Update)) { }

            //Change back to .pptx
            FileInfo f2 = new FileInfo(path + ".zip");
            f2.MoveTo(Path.ChangeExtension(path, ".pptx"));

            //Do the read
            OpenXMLReader reader = new OpenXMLReader(originalPath);
            reader.read();
            reader.PresentationObject.getXMLTree().Save(@"C:\Users\ex1\Desktop\out.xml");

            //Delete file
            File.Delete(path + ".pptx");

            Console.WriteLine("\nCompilation time: " + (double)watch.ElapsedMilliseconds/1000 + "s");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
