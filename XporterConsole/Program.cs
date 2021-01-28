using OfficeOpenXml;
using System;
using System.IO;
using Xporter.Core.Extensions;

namespace XporterConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello Xporter World!");

            Xporter.Core.Xporter.Load(new FileStream("C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\Test.xlsx", FileMode.OpenOrCreate));
            var x = Xporter.Core.Xporter.CreateOrLoad("C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\Test.xlsx", "TestFileName01", "TestSheetName01");


        }
    }
}
