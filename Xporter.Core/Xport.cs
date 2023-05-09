using OfficeOpenXml;
using System;
using System.IO;
using System.Linq;

namespace Xporter
{
    /// <summary>
    /// This class helps you export any kind of data to an xlsx file
    /// </summary>
    public static class Xport
    {
        /// <summary>
        /// Load an existing xlsx File
        /// </summary>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage LoadFromFileInfo(string path)
        {
            var package = new ExcelPackage(new FileInfo(path));

            return package;
        }

        /// <summary>
        /// Creates or Loads an xlsx file
        /// </summary>
        /// <param name="path">Export or Load path</param>
        /// <param name="fileName">(Nullable) File name</param>
        /// <param name="sheetName">(Nullable) SpreadSheet name</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage LoadFromStream(FileStream fileStream, string sheetName)
        {
            try
            {
                var package = new ExcelPackage(fileStream);

                    if (!package.Workbook.Worksheets.Select(s=>s.Name = sheetName).Any())
                    {
                        var activeSheet = package.Workbook.Worksheets.Add(sheetName); 
                    }
                    return package;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// Creates a new XlsxPackage <br></br>
        /// >  Do not forget to use SaveAs() method
        /// </summary>
        /// <returns></returns>
        public static ExcelPackage CreateNewPackage()
        {
            return new ExcelPackage();
        }
    }
}
