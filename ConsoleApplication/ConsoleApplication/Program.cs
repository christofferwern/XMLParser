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
            string imagesPath = @"C:\Users\ex1\Desktop\PPTImages";
            //Copyy file
            File.Copy(originalPath, path + ".pptx");

            //Change to .zip
            FileInfo f1 = new FileInfo(path + ".pptx");
            f1.MoveTo(Path.ChangeExtension(path, ".zip"));

            //Open zip file
            using (ZipArchive zip = ZipFile.Open(path + ".zip", ZipArchiveMode.Update)) {
                
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (entry.FullName.EndsWith(".jpeg", StringComparison.OrdinalIgnoreCase) || (entry.FullName.EndsWith(".png", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (Directory.Exists(imagesPath))
                        {
                            string[] fileEntries = Directory.GetFiles(imagesPath);
                            int count = 0;
                            bool exists = false;

                            while (count < fileEntries.Length)
                            {
                                string last = fileEntries[count].Substring(fileEntries[count].LastIndexOf('\\') + 1);
                                if (last == entry.Name.ToString())
                                {
                                    exists = true;
                                    break;
                                }
                                count++;
                            }

                            if (!exists)
                                entry.ExtractToFile(Path.Combine(imagesPath, entry.Name));

                        }
                        Console.WriteLine("-------------");
                    }
                    
                }

            }

            //Change back to .pptx
            FileInfo f2 = new FileInfo(path + ".zip");
            f2.MoveTo(Path.ChangeExtension(path, ".pptx"));

            //Do the read
            OpenXMLReader reader = new OpenXMLReader(path + ".pptx");
            reader.read();
            reader.PresentationObject.getXMLTree().Save(@"C:\Users\ex1\Desktop\out.xml");

            //Delete copied file
            File.Delete(path + ".pptx");

            Console.WriteLine("\nCompilation time: " + (double)watch.ElapsedMilliseconds/1000 + "s");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
