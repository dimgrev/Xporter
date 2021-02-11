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



            //Properties...

            var cp = new CellProperties();
            cp.Add("A2", "Stats02");
            cp.Add("B4", "TypeOfProduct02");
            cp.Add("B6", "Images02");

            var exportPath = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports";

            var filePath = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\Test.xlsx";

            var templFullPath = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Templates\\Templ.xlsx";



            //USAGE EXAMPLE...

            Xport.Load(filePath)
                 .Clear()
                 //.LoadTempl(new FileStream(templFullPath, FileMode.Open))
                 .InsertData(stds);
            
            //USAGE EXAMPLE2...

            Xport.Create(exportPath, "TestFileName01.xlsx", "TestSheetName01")
                 .LoadTempl(new FileStream(templFullPath, FileMode.Open))
                 .InsertData(stds, 8, 2)
                 .WriteToCells(cp);
        }
    }
}
