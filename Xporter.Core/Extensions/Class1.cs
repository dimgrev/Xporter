using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Xporter
{
    public static class Class1
    {
        /// <summary>
        /// Insert any object type or list of properties to the current package
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <param name="startingRow">In which row you want the program to start inserting data</param>
        /// <param name="startingCol">In which column you want the program to start inserting data</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage InsertData2(this ExcelPackage pack, List<object> objs, int startingRow, int startingCol)
        {
            var sheet = LoadSheet(pack);
            
            var rowFlag = 0;
            var row = startingRow + 1;

            for (int j = 0; j < objs.Count; j++)
            {
                var item = objs[j];

                //Takes the type of the first object
                var firstObjType = item.GetType();

                //Get all Properties from that type class
                var props = firstObjType.GetProperties();

                for (int i = 0; i < props.Length; i++)
                {
                    //var newI = startingIndex + i;

                    sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row.ToString()].Value = props[i].Name;
                }
                row += 2;

                var rowb = row;

                for (int i = 0; i < props.Length; i++)
                {
                    var prop = item.GetType().GetProperty(props[i].Name).GetValue(item);

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
                            sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + rowb].Value = ad.ToString();


                            sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + rowb].
                                Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                            sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + rowb].
                                Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            rowb++;
                            rowFlag = rowFlag < rowb ? rowb : rowFlag;
                        }

                    }
                    else
                    {
                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row].Value = item.GetType()
                            .GetProperty(props[i].Name)
                            .GetValue(item, null).ToString();

                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row].
                            Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Center;

                        sheet.Cells[ExcelCellAddress.GetColumnLetter(i + startingCol) + row].
                            Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        rowFlag = rowFlag < row ? row : rowFlag;
                    }
                    rowb = row;
                }
                row = rowFlag + 2;
            }
            return pack;
        }

        /// <summary>
        /// Loads Sheet from package
        /// </summary>
        /// <param name="pack"></param>
        /// <returns>ExcelWorksheet</returns>
        private static ExcelWorksheet LoadSheet(ExcelPackage pack)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var activeSheet = pack.Workbook.Worksheets.FirstOrDefault();

            if (activeSheet is null)
            {
                activeSheet = pack.Workbook.Worksheets.Add("Report");
            }

            return activeSheet;
        }
    }
}
