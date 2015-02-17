using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

            string path = @"C:\Users\ex1\desktop\Compare.pptx";            


            OpenXMLReader reader = new OpenXMLReader(path);

            reader.read();

            reader.PresentationObject.getXMLTree().Save(Console.Out);

            //// Write the XML to a file.
            System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\ex1\Desktop\test.xml", true);
            file.WriteLine(reader.PresentationObject.getXMLTree().InnerXml);

            file.Close();
            
            watch.Stop();
            Console.WriteLine("\nCompilation time: " + (double)watch.ElapsedMilliseconds/1000 + "s");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
