using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.IO.Compression;
using System.Web;



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

            string originalPath = @"C:\Users\ex1\desktop\Performance_Out.pptx";

            //Get's the pptx name
            string originalFileName = originalPath.Substring(originalPath.LastIndexOf('\\') + 1);

            string path = @"C:\Users\ex1\desktop\randomStuffNotANYHAVEBOfDY12723489";

            //Image folder location
            string imagesPath = @"C:\Users\ex1\Desktop\" + originalFileName.Split('.')[0] + "_img";

            File.Copy(originalPath, path + ".pptx");

            //Change to .zip
            FileInfo f1 = new FileInfo(path + ".pptx");
            f1.MoveTo(Path.ChangeExtension(path, ".zip"));

            if(!Directory.Exists(imagesPath)){
                DirectoryInfo di = Directory.CreateDirectory(imagesPath);
            }

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
                    }
                    
                }

            }

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
