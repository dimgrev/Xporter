using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Xporter.Core.Extensions
{
    public static class Ext
    {
        /// <summary>
        /// Load an existing template in the current file
        /// </summary>
        /// <param name="package">This xlsx package as extension method</param>
        /// <param name="stream">Put the Template file stream</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage LoadTempl(this ExcelPackage package, Stream stream)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var templPackage = new ExcelPackage(stream);

            var templSheet = LoadSheet(templPackage);

            var activeSheet = LoadSheet(package);

            if (activeSheet != null)
            {
                package.Workbook.Worksheets.Copy(activeSheet.Name, activeSheet.Name + "D");
                package.Workbook.Worksheets.Delete(activeSheet.Name);

                package.Workbook.Worksheets.Add(activeSheet.Name, templSheet);
                package.Workbook.Worksheets.Delete(activeSheet.Name + "D");

                package.Save();

                stream.Close();

                return package;
            }
            else
            {
                package.Workbook.Worksheets.Add(activeSheet.Name, templSheet);

                package.Save();

                stream.Close();

                return package;
            }
        }

        /// <summary>
        /// Insert any object type or list of properties to the current package
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <param name="startingRow">In which row you want the program to start inserting data</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage InsertData(this ExcelPackage pack, List<object> objs, int startingRow)
        {
            var sheet = LoadSheet(pack);

            //Takes the type of the first object
            var firstObjType = objs.First().GetType();

            //Get all Properties from that type class
            var props = firstObjType.GetProperties();

            for (int i = 0; i < props.Length; i++)
            {
                //var newI = startingIndex + i;

                sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + startingRow.ToString()].Value = props[i].Name;
            }

            var row = startingRow + 2;
            var rowf = 0;

            foreach (var item in objs)
            {
                var rowb = row;
                for (int i = 0; i < props.Length; i++)
                {

                    var prop = item.GetType().GetProperty(props[i].Name).GetValue(item);


                    //Alternative....
                    //List<Object> collection = new List<Object>((IEnumerable<Object>)prop);
                    //.....Works.....

                    if (prop is IEnumerable<Object>)
                    {
                        List<object> list = new List<object>();
                        var enumerator = ((IEnumerable<Object>)prop).GetEnumerator();
                        while (enumerator.MoveNext())
                        {
                            list.Add(enumerator.Current);
                        }
                        foreach (var ad in list)
                        {
                            //do what you want here
                            sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + rowb].Value = ad.ToString();


                            sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + rowb].
                                Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + rowb].
                                Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            rowb++;
                            rowf = rowf < rowb ? rowb : rowf;
                        }

                    }
                    else
                    {
                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + row].Value = item.GetType()
                            .GetProperty(props[i].Name)
                            .GetValue(item, null).ToString();

                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + row].
                            Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + 1) + row].
                            Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    }
                    rowb = row;
                    //foreach (var ad in collection)
                    //{
                    //    //do what you want here
                    //}

                }
                row = rowf + 1;
            }
            var allCells = sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column];

            var cellFont = allCells.Style.Font;
            cellFont.Name = "Bahnschrift Light SemiCondensed";

            pack.Save();

            return pack;
        }

        /// <summary>
        /// 
        /// </summary>
        public static ExcelPackage WriteToCells(this ExcelPackage pack, CellProperties cp)
        {
            var sheet = LoadSheet(pack);

            foreach (var item in cp)
            {
                sheet.Cells[item.Key].Value = item.Value;
            }

            pack.Save();

            return pack;
        }

        /// <summary>
        /// Clears all data of the xlsx file
        /// </summary>
        /// <param name="package"></param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage Clear(this ExcelPackage package)
        {
            var sheet = LoadSheet(package);

            sheet.Cells.Clear();

            package.Save();

            return package;
        }


        /// <summary>
        /// Loads Sheet from package
        /// </summary>
        /// <param name="pack"></param>
        /// <returns>ExcelWorksheet</returns>
        private static ExcelWorksheet LoadSheet(ExcelPackage pack)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var activeSheet = pack.Workbook.Worksheets.First();

            return activeSheet;
        }
    }
}
