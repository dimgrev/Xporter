using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using Xporter.Core.Extensions;
using Xporter.Core;

namespace XporterConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello Xporter World!");



            var stds = new List<object>();

            var std = new Students();
            std.FirstName = new List<string>();
            std.LastName = new List<string>();

            std.FirstName.Add("DimitrisA" + 1);
            std.FirstName.Add("DimitrisB" + 1);
            std.FirstName.Add("DimitrisC" + 1);
            std.FirstName.Add("DimitrisD" + 1);
            std.FirstName.Add("DimitrisE" + 1);

            std.LastName.Add("GrevenosA" + 1);
            std.LastName.Add("GrevenosB" + 1);
            std.LastName.Add("GrevenosC" + 1);
            //std.LastName.Add("GrevenosD"+1);
            //std.LastName.Add("GrevenosE"+1);

            stds.Add(std);

            var std2 = new Students();
            std2.FirstName = new List<string>();
            std2.LastName = new List<string>();

            std2.FirstName.Add("DimitrisA" + 2);
            std2.FirstName.Add("DimitrisB" + 2);
            std2.FirstName.Add("DimitrisC" + 2);
            std2.FirstName.Add("DimitrisD" + 2);
            std2.FirstName.Add("DimitrisE" + 2);

            std2.LastName.Add("GrevenosA" + 2);
            std2.LastName.Add("GrevenosB" + 2);

            stds.Add(std2);

            var cp = new CellProperties();
            cp.Add("A2", "Stats");
            cp.Add("B4", "TypeOfProduct");
            cp.Add("B6", "Images");

            //Xport.Load(new FileStream("C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\Test.xlsx", FileMode.Open))
            //        .LoadTempl(new FileStream("C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Templates\\Templ.xlsx", FileMode.Open))
            //        .InsertData(stds, 8);

            var x = Xport.CreateOrLoad("C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports", "TestFileName01", "TestSheetName01")
                    .LoadTempl(new FileStream("C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Templates\\Templ.xlsx", FileMode.Open))
                    .InsertData(stds, 8)
                    .WriteToCells(cp);
        }
    }
}
