using System;

namespace XporterConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello Xporter World!");

            Xporter.Core.Xporter.CreateOrLoadFile("Test");
        }
    }
}
