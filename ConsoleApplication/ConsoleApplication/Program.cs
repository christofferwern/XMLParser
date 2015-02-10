﻿using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            var watch = Stopwatch.StartNew();

            SimpleSceneObject SSO = new SimpleSceneObject();
            
            TextObject TO = new TextObject(SSO);

            TO.getXMLTree();

            watch.Stop();
            Console.WriteLine("Compilation time: " + (double)watch.ElapsedMilliseconds/1000 + "s");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }
}
