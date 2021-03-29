using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Xporter
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
            var templPackage = new ExcelPackage(stream);

            var templSheet = LoadSheet(templPackage);

            var activeSheet = LoadSheet(package);

            if (activeSheet != null)
            {
                package.Workbook.Worksheets.Copy(activeSheet.Name, activeSheet.Name + "D");
                package.Workbook.Worksheets.Delete(activeSheet.Name);

                package.Workbook.Worksheets.Add(activeSheet.Name, templSheet);
                package.Workbook.Worksheets.Delete(activeSheet.Name + "D");

                stream.Close();

                return package;
            }
            else
            {
                package.Workbook.Worksheets.Add(activeSheet.Name, templSheet);

                stream.Close();

                return package;
            }
        }

        /// <summary>
        /// Insert any object type or list of properties to the current package
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <param name="startingRow">In which row you want the program to start inserting data (starts at 1)</param>
        /// <param name="startingCol">In which column you want the program to start inserting data (starts at 1)</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage InsertData(this ExcelPackage pack, List<object> objs, int startingRow, int startingCol)
        {
            var sheet = LoadSheet(pack);

            var rowFlag = 0;
            var row = startingRow;

            for (int j = 0; j < objs.Count; j++)
            {
                var item = objs[j];

                if (item is IEnumerable<Object>)
                {
                    foreach (var it in ((IEnumerable<Object>)item))
                    {
                        var r = Insert(sheet, rowFlag, row, startingCol, it);
                        row = r;
                    }
                }
                else
                {
                    Insert(sheet, rowFlag, row, startingCol, item);
                }
            }
            return pack;
        }

        //Under load
        public static ExcelPackage InsertData(this ExcelPackage pack, List<object> objs)
        {
            var sheet = LoadSheet(pack);

            var rowFlag = 0;
            var row = 1;

            for (int j = 0; j < objs.Count; j++)
            {
                var item = objs[j];

                if (item is IEnumerable<Object>)
                {
                    foreach (var it in ((IEnumerable<Object>)item))
                    {
                        var r = Insert(sheet, rowFlag, row, 1, it);
                        row = r;
                    }
                }
                else
                {
                    Insert(sheet, rowFlag, row, 1, item);
                }
            }
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

            return pack;
        }

        /// <summary>
        /// Clears all data of the xlsx file 
        /// <br></br>
        /// (Works only with FileInfo NOT Stream)
        /// </summary>
        /// <param name="package"></param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage Clear(this ExcelPackage package)
        {
            package.Workbook.Worksheets.ToList().ForEach(f=>f.Cells.Clear());

            return package;
        }

        /// <summary>
        /// Clears all data of the xlsx WorkSheet
        /// <br></br>
        /// (Works only with FileInfo NOT Stream)
        /// </summary>
        /// <param name="package"></param>
        /// <param name="SheetName">WorkSheet Name to clear</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage Clear(this ExcelPackage package, string SheetName)
        {
            package.Workbook.Worksheets.Where(w => w.Name == SheetName)
                                       .FirstOrDefault().Cells
                                       .Clear();

            return package;
        }


        /// <summary>
        /// Loads Sheet from package
        /// </summary>
        /// <param name="pack"></param>
        /// <returns>ExcelWorksheet</returns>
        private static ExcelWorksheet LoadSheet(ExcelPackage pack)
        {
            var activeSheet = pack.Workbook.Worksheets.FirstOrDefault();

            if (activeSheet is null)
            {
                activeSheet = pack.Workbook.Worksheets.Add("Report");
            }

            return activeSheet;
        }


        private static int Insert(ExcelWorksheet sheet, int rowFlag, int row, int startingCol, object item)
        {
            //Takes the type of the first object
            var firstObjType = item.GetType();

            //Get all Properties from that type class
            var props = firstObjType.GetProperties();

            for (int i = 0; i < props.Length; i++)
            {
                //var newI = startingIndex + i;

                sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row.ToString()].Value = (props[i].Name ?? "null").ToString() ?? "null";

                sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row.ToString()]
                     .Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row.ToString()]
                     .Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            row += 2;

            var rowb = row;

            for (int i = 0; i < props.Length; i++)
            {
                var prop = item.GetType().GetProperty(props[i].Name).GetValue(item) ?? "null";

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
                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + rowb.ToString()].Value = (ad ?? "null").ToString() ?? "null";

                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + rowb.ToString()]
                             .Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + rowb.ToString()]
                             .Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        rowb++;
                        rowFlag = rowFlag < rowb ? rowb : rowFlag;
                    }

                }
                else
                {
                    sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row].Value = (item.GetType()
                        .GetProperty(props[i].Name)
                        .GetValue(item, null) ?? "null")
                        .ToString() ?? "null";

                    sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row.ToString()]
                         .Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                    sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row.ToString()]
                         .Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                    rowFlag = rowFlag < row ? row : rowFlag;
                }
                rowb = row;
            }
            row = rowFlag + 2;

            return row;
        }
    }
}
