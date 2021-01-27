using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace Xporter.Core
{
    /// <summary>
    /// This class helps you export any kind of data to an xlsx file
    /// </summary>
    public static class Xporter
    {
        /// <summary>
        /// Creates a new SpreadSheet file
        /// </summary>
        /// <returns></returns>
        public static FileInfo Create()
        {


            throw new NotImplementedException();
        }

        public static FileInfo CreateOrLoadFile( string fileName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                string exportPath = Path.Combine((new System.Uri(Assembly.GetExecutingAssembly().CodeBase))
                    .AbsolutePath.Split(new string[] { "/bin" }
                    , StringSplitOptions.None)[0]) 
                    + $"/Exports";


                if (!Directory.Exists(exportPath))
                {
                    Directory.CreateDirectory(exportPath);
                }

                var exportFilename = fileName + ".xlsx";
                var file = new FileInfo(Path.Combine(exportPath, exportFilename));

                //CreateExcelSpreadSheet
                //CreateXlsxSpreadSheet(file, sheetName, startingIndex, obj);
                var package = new ExcelPackage(file);

                var activeSheet = package.Workbook.Worksheets.Add("TestSheetName");

                package.Save();

                return file;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
