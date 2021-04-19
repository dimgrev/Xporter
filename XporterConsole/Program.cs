using System;
using System.Collections.Generic;
using System.IO;
using Xporter;

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

            //Adding Anonymous type List<object> in List<object> for InsertData
            var nestedObj = new List<object>();
            nestedObj.Add(new { userUSERUSER = 1, firstNameNameName = 2, lastNameNAMENAMENAME = 3 });
            nestedObj.Add(new { user = (string)null, firstName = 5, lastName = 6 });
            nestedObj.Add(new { });
            nestedObj.Add(new { user = 7, firstName = 8, lastName = 9 });
            nestedObj.Add(new { user = 10, firstName = (string)null, lastName = 12 });
            nestedObj.Add(new { user = 10, firstName = 11, lastname = (string)null});

            //var query = from Students in stds
            //            where 

            stds.Add(nestedObj);

            var nullObj = new List<object>();
            stds.Add(new { });
            stds.Add(nestedObj);
            stds.Add(new { Sum = "100"});
            stds.Add(new {});
            stds.Add(nestedObj);

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

            //stds.Add(std2);

            var obj2 = new List<object>();

            var std3 = new Students();
            std3.FirstName = new List<string>();
            std3.LastName = new List<string>();

            std3.FirstName.Add("DimitrisA" + 1);
            std3.FirstName.Add("DimitrisB" + 1);
            std3.FirstName.Add("DimitrisC" + 1);
            std3.FirstName.Add("DimitrisD" + 1);
            std3.FirstName.Add("DimitrisE" + 1);

            std3.LastName.Add("GrevenosA" + 1);
            std3.LastName.Add("GrevenosB" + 1);
            std3.LastName.Add("GrevenosC" + 1);

            obj2.Add(std3);

            var tch = new Teachers();
            tch.FirstName = new List<string>();
            tch.LastName = new List<string>();

            tch.FirstName.Add("DimitrisA" + 1);
            tch.FirstName.Add("DimitrisB" + 1);
            tch.FirstName.Add("DimitrisC" + 1);
            tch.FirstName.Add("DimitrisD" + 1);
            tch.FirstName.Add("DimitrisE" + 1);

            tch.LastName.Add("GrevenosA" + 1);
            tch.LastName.Add("GrevenosB" + 1);
            tch.LastName.Add("GrevenosC" + 1);

            tch.TeacherNumber = 2;

            obj2.Add(tch);

            //Properties...

            var cp = new CellProperties();
            cp.Add("E2", "Stats");
            cp.Add("E3", "TypeOfProduct");
            cp.Add("E4", "Images");
            cp.Add("I2", DateTime.Now.ToString());

            var exportPath = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\TestFileName1.xlsx";

            var exportPath2 = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\TestFileName2.xlsx";

            var filePath = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\Test.xlsx";

            var templFullPath = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Templates\\templ.xlsx";

            var templ = @"C:\Users\dgrevenos\Desktop\Template.xlsx";


            //USAGE EXAMPLE1...
            var templStream = new FileStream(templ, FileMode.Open);
            Xport.LoadFromFileInfo(filePath)
                 .Clear()
                 .LoadTempl(templStream)
                 .InsertData(stds, 8, 2)
                 .LoadTempl("report2", templStream)
                 .InsertData("report2", stds, 8, 2)
                 .WriteToCells(cp)
                 .Save();


            //USAGE EXAMPLE2...

            var fileStream = new FileStream(exportPath, FileMode.OpenOrCreate);
            Xport.LoadFromStream(fileStream, "TestSheetName")
                 .InsertData(stds, 8, 2)
                 .WriteToCells(cp)
                 .Save();

            fileStream.Close();


            //USAGE EXAMPLE3...

            var exportPath3 = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\TestFileName3.xlsx";

            var fileStream2 = new FileStream(exportPath2, FileMode.OpenOrCreate);
            Xport.LoadFromStream(fileStream2, "TestSheetName")
                .InsertData(obj2).Save();

            fileStream2.Close();


            //var fileStream3 = new FileStream(exportPath3, FileMode.OpenOrCreate);

            Xport.LoadFromFileInfo(exportPath2)
                .Clear()
                .SaveAs( new FileInfo(exportPath3));

            //USAGE EXAMPLE4...
            var exportPath4 = "C:\\Users\\dgrevenos\\source\\repos\\Xporter\\XporterConsole\\Exports\\";

            Xport.CreateNewPackage().InsertData(stds, 8, 2).SaveAs(new FileInfo(exportPath4 + "New Microsoft Excel Worksheet.xlsx"));


            //USAGE EXAMPLE5...

            //var fileStream = new FileStream(exportPath, FileMode.OpenOrCreate);
            //Xport.CreateOrLoad(fileStream, "TestSheetName")
            //     .LoadTempl(new FileStream(templFullPath, FileMode.Open))
            //     .InsertData(stds, 8, 2)
            //     .WriteToCells(cp);


        }
    }
}
