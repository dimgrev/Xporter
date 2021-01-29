using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Xporter.Core
{
    /// <summary>
    /// This class helps you export any kind of data to an xlsx file
    /// </summary>
    public static class Xport
    {
        /// <summary>
        /// Load an existing xlsx Filestream
        /// </summary>
        /// <returns>ExcelPackage</returns>
        //public static ExcelPackage Load(Stream stream)
        //{
        //    var package = LoadPackage(stream);

        //    stream.Close();

        //    return package;
        //}

        /// <summary>
        /// Creates or Loads an xlsx file
        /// </summary>
        /// <param name="path">Export or Load path</param>
        /// <param name="fileName">File name</param>
        /// <param name="sheetName">SpreadSheet name</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage CreateOrLoad(string path, string fileName, string sheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                string export = Path.Combine(path);


                if (!File.Exists(path))
                {
                    if (!Directory.Exists(export))
                    {
                        Directory.CreateDirectory(export);
                    }

                    var exportFilename = fileName + ".xlsx";

                    //var file = new FileInfo(Path.Combine(export, exportFilename));
                    var file = new FileInfo(export +"\\"+ exportFilename);
                    

                    var package = new ExcelPackage(file);

                    if (!package.Workbook.Worksheets.Select(s=>s.Name = sheetName).Any())
                    {
                        var activeSheet = package.Workbook.Worksheets.Add(sheetName); 
                    }

                    package.Save();

                    return package;
                }
                else
                {
                    var package = new ExcelPackage(new FileInfo(path));
                    
                    return package;
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// Loads package from stream
        /// </summary>
        /// <param name="stream"></param>
        /// <returns>ExcelPackage</returns>
        //private static ExcelPackage LoadPackage(Stream stream)
        //{
        //    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        //    var package = new ExcelPackage(stream);

        //    return package;
        //}
    }
}
