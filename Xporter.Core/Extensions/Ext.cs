using OfficeOpenXml;
using System.Collections.Generic;
using System.Drawing;
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
            return LoadTempl(package, null, stream);
        }

        /// <summary>
        /// Loads an existing template in the current file in the specified sheet
        /// </summary>
        /// <param name="package">This xlsx package as extension method</param>
        /// <param name="SheetName">Specify Sheet Name for templ to be loaded</param>
        /// <param name="stream">Put the Template file stream</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage LoadTempl(this ExcelPackage package, string SheetName, Stream stream)
        {
            var templPackage = new ExcelPackage(stream);

            var templSheet = LoadSheet(templPackage);

            var activeSheet = LoadSheet(package, SheetName);

            if (activeSheet != null)
            {
                package.Workbook.Worksheets.Copy(activeSheet.Name, activeSheet.Name + "D");
                package.Workbook.Worksheets.Delete(activeSheet.Name);

                package.Workbook.Worksheets.Add(activeSheet.Name, templSheet);
                package.Workbook.Worksheets.Delete(activeSheet.Name + "D");

                return package;
            }
            else
            {
                package.Workbook.Worksheets.Add(activeSheet.Name, templSheet);

                return package;
            }
        }


        /// <summary>
        /// Insert any object type or list of properties to a (new/existing) Sheet
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="SheetName">Insert data to a new or existing Sheet Name</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <param name="startingRow">In which row you want the program to start inserting data (starts at 1)</param>
        /// <param name="startingCol">In which column you want the program to start inserting data (starts at 1)</param>
        /// <returns>ExcelPackage</returns>
        public static ExcelPackage InsertData(this ExcelPackage pack, List<object> objs)
        {
            return InsertData(pack, objs, 1, 1);
        }
        public static ExcelPackage InsertData(this ExcelPackage pack, List<object> objs, int startingRow, int startingCol)
        {
            return InsertData(pack, null, objs, startingRow, startingCol);
        }

        public static ExcelPackage InsertData(this ExcelPackage pack, string SheetName, List<object> objs, int startingRow, int startingCol)
        {
            if (objs == null || objs.Count == 0)
                return pack;

            var sheet = LoadSheet(pack);

            if (SheetName != null)
                sheet = LoadSheet(pack, SheetName);

            var rowFlag = 0;
            var row = startingRow;

            if (objs.Count <= 1)
            {
                InsertDataObjs(pack, objs);
            }
            else
            {
                var first = objs[0];
                var commonType = first?.GetType();
                var allSame = true;
                if (commonType is null)
                {
                    // if first is null, ensure all others are null to be treated as same
                    for (int i = 1; i < objs.Count; i++)
                    {
                        if (objs[i] != null)
                        {
                            allSame = false;
                            break;
                        }
                    }
                }
                else
                {
                    for (int i = 1; i < objs.Count; i++)
                    {
                        var o = objs[i];
                        if (o == null || o.GetType() != commonType)
                        {
                            allSame = false;
                            break;
                        }
                    }
                }

                if (allSame && commonType != null)
                {
                    row = InsertProperties(sheet, rowFlag, row, startingCol, first);

                    for (int j = 0; j < objs.Count; j++)
                    {
                        var item = objs[j];

                        if (item == null)
                        {
                            row++;
                            continue;
                        }

                        var itemType = item.GetType();
                        var props = itemType.GetProperties();

                        // Skip plain System.Object placeholders
                        if (itemType == typeof(object) && props.Length == 0)
                        {
                            row++;
                            continue;
                        }

                        if (props.Length == 0)
                        {
                            // primitive or empty object -> write single cell in first column
                            if (item is string s && string.IsNullOrWhiteSpace(s))
                            {
                                row++;
                                continue;
                            }

                            var col = ExcelCellAddress.GetColumnLetter(1);
                            if (item is int || item is long || item is short || item is decimal || item is double || item is float)
                                sheet.Cells[col + row.ToString()].Value = item;
                            else
                                sheet.Cells[col + row.ToString()].Value = item.ToString();

                            sheet.Cells[col + row.ToString()]
                                 .Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                            sheet.Cells[col + row.ToString()]
                                 .Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            sheet.Cells[col + row.ToString()]
                                 .AutoFitColumns();

                            rowFlag = rowFlag < row ? row : rowFlag;
                            row++;
                        }
                        else
                        {
                            row = InsertItems(sheet, rowFlag, row, startingCol, item);
                        }
                    }
                }
                else
                {
                    InsertDataObjs(pack, SheetName, objs, startingRow, startingCol);
                }
            }
            return pack;
        }

        /// <summary>
        /// Insert any object type or list of properties to the current package
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <returns>ExcelPackage</returns>
        private static ExcelPackage InsertDataObjs(this ExcelPackage pack, List<object> objs)
        {
            return InsertDataObjs(pack, objs, 1, 1);
        }

        //Over Load method

        /// <summary>
        /// Insert any object type or list of properties to the current package
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <param name="startingRow">In which row you want the program to start inserting data (starts at 1)</param>
        /// <param name="startingCol">In which column you want the program to start inserting data (starts at 1)</param>
        /// <returns>ExcelPackage</returns>
        private static ExcelPackage InsertDataObjs(this ExcelPackage pack, List<object> objs, int startingRow, int startingCol)
        {
            return InsertDataObjs(pack, null, objs, startingRow, startingCol);
        }

        /// <summary>
        /// Insert any object type or list of properties to a (new/existing) Sheet
        /// </summary>
        /// <param name="pack">This xlsx package as extension method</param>
        /// <param name="SheetName">Insert data to a new or existing Sheet Name</param>
        /// <param name="objs">The list of your Data that you want to insert</param>
        /// <param name="startingRow">In which row you want the program to start inserting data (starts at 1)</param>
        /// <param name="startingCol">In which column you want the program to start inserting data (starts at 1)</param>
        /// <returns>ExcelPackage</returns>
        private static ExcelPackage InsertDataObjs(this ExcelPackage pack, string SheetName, List<object> objs, int startingRow, int startingCol)
        {
            if (objs == null || objs.Count == 0)
                return pack;

            var sheet = LoadSheet(pack);

            if (SheetName != null)
                sheet = LoadSheet(pack, SheetName);

            var rowFlag = 0;
            var row = startingRow;

            foreach (var item in objs)
            {
                if (item is System.Collections.IEnumerable && !(item is string))
                {
                    var enumerable = (System.Collections.IEnumerable)item;

                    object firstNonNull = null;
                    foreach (var e in enumerable)
                    {
                        if (e == null) continue;
                        var eProps = e.GetType().GetProperties();
                        if (eProps.Length > 0 || !(e is string && string.IsNullOrWhiteSpace((string)e)))
                        {
                            firstNonNull = e;
                            break;
                        }
                    }

                    if (firstNonNull != null)
                    {
                        row = InsertProperties(sheet, rowFlag, row, startingCol, firstNonNull);

                        foreach (var it in enumerable)
                        {
                            if (it == null)
                            {
                                row++;
                                continue;
                            }

                            var itType = it.GetType();
                            var itProps = itType.GetProperties();

                            if (itProps.Length == 0)
                            {
                                // primitive or empty object -> write single cell where reasonable, else skip a row
                                if (it is string s && string.IsNullOrWhiteSpace(s))
                                {
                                    row++;
                                    continue;
                                }

                                var cell = sheet.Cells[row, startingCol];
                                if (it is int || it is long || it is short || it is decimal || it is double || it is float)
                                    cell.Value = it;
                                else
                                    cell.Value = it.ToString();

                                cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                                cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                                rowFlag = rowFlag < row ? row : rowFlag;
                                row++;
                                continue;
                            }

                            row = InsertItems(sheet, rowFlag, row, startingCol, it);
                        }
                        row += 2;
                    }
                    else
                        row++;
                }
                else
                {
                    if (item == null)
                    {
                        row++;
                        continue;
                    }

                    var itemType = item.GetType();
                    var props = itemType.GetProperties();

                    if (props.Length == 0)
                    {
                        // primitive or empty object -> write single cell where reasonable, else skip a row
                        if (item is string s && string.IsNullOrWhiteSpace(s))
                        {
                            row++;
                            continue;
                        }

                        var cell = sheet.Cells[row, startingCol];
                        if (item is int || item is long || item is short || item is decimal || item is double || item is float)
                            cell.Value = item;
                        else
                            cell.Value = item.ToString();

                        cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                        cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        rowFlag = rowFlag < row ? row : rowFlag;
                        row++;
                    }
                    else
                    {
                        row = InsertProperties(sheet, rowFlag, row, startingCol, item);
                        row = InsertItems(sheet, rowFlag, row, startingCol, item);
                        row++;
                    }
                }
            }
            return pack;
        }

        private static int InsertItems(ExcelWorksheet sheet, int rowFlag, int row, int startingCol, object item)
        {
            if (item == null)
                return row;

            //Takes the type of the first object
            var firstObjType = item.GetType();

            //Get all Properties from that type class
            var props = firstObjType.GetProperties();

            row += 1;

            if (props.Length > 0)
            {
                var rowb = row;

                for (int i = 0; i < props.Length; i++)
                {
                    var prop = props[i].GetValue(item, null) ?? "null";
                    var col = ExcelCellAddress.GetColumnLetter(i + startingCol);

                    if (prop is System.Collections.IEnumerable && !(prop is string))
                    {
                        foreach (var adObj in (System.Collections.IEnumerable)prop)
                        {
                            var ad = adObj ?? "null";

                            if (ad is int || ad is long || ad is short || ad is decimal || ad is double || ad is float)
                                sheet.Cells[col + rowb.ToString()].Value = ad;
                            else
                                sheet.Cells[col + rowb.ToString()].Value = ad.ToString();

                            sheet.Cells[col + rowb.ToString()]
                                 .Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                            sheet.Cells[col + rowb.ToString()]
                                 .Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                            sheet.Cells[col + rowb.ToString()]
                                 .AutoFitColumns();

                            rowb++;
                            rowFlag = rowFlag < rowb ? rowb : rowFlag;
                        }
                    }
                    else
                    {
                        if (prop is int || prop is long || prop is short || prop is decimal || prop is double || prop is float)
                            sheet.Cells[col + row.ToString()].Value = prop;
                        else
                            sheet.Cells[col + row.ToString()].Value = prop.ToString();

                        sheet.Cells[col + row.ToString()]
                             .Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;

                        sheet.Cells[col + row.ToString()]
                             .Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;

                        rowFlag = rowFlag < row ? row : rowFlag;
                    }
                    rowb = row;
                }
                row = rowFlag;
            }
            return row;
        }

        /// <summary>
        /// Write data in specific cells in all sheets
        /// </summary>
        /// <param name="pack"></param>
        /// <param name="cp">Create a new CellProperties() var cp<br></br>
        /// cp.Add (cells and values) -repeat .Add<br></br>
        /// and insert it to this method</param>
        /// <returns></returns>
        public static ExcelPackage WriteToCells(this ExcelPackage pack, CellProperties cp)
        {
            return WriteToCells(pack, null, cp);
        }

        /// <summary>
        /// Write data in specific cells in a sheet
        /// </summary>
        /// <param name="pack"></param>
        /// <param name="sheetName">The sheet name where to write cp data</param>
        /// <param name="cp">Create a new CellProperties() var cp<br></br>
        /// cp.Add (cells and values) -repeat .Add<br></br>
        /// and insert it to this method</param>
        /// <returns></returns>
        public static ExcelPackage WriteToCells(this ExcelPackage pack, string sheetName, CellProperties cp)
        {
            var sheetList = pack.Workbook.Worksheets.ToList();

            if (sheetName != null)
                sheetList = new List<ExcelWorksheet>() { LoadSheet(pack, sheetName) };

            sheetList.ForEach(f =>
            {
                foreach (var item in cp)
                {
                    f.Cells[item.Key].Value = item.Value;
                    f.Cells[item.Key].Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    f.Cells[item.Key].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                }
            });
            return pack;
        }

        /// <summary>
        /// Inserts the specified value into cells within the Excel package that match the given cell value.
        /// </summary>
        /// <param name="pack">The ExcelPackage instance containing the worksheet to modify. Cannot be null.</param>
        /// <param name="cellValue">The value to search for in the cells. Cells matching this value will be updated.</param>
        /// <param name="valueToInsert">The value to insert into the matching cells.</param>
        /// <returns>The ExcelPackage instance with the updated cells.</returns>
        public static ExcelPackage InsertToCells(this ExcelPackage pack, string cellValue, string valueToInsert)
        {
            return InsertToCells(pack, null, cellValue, valueToInsert);
        }

        /// <summary>
        /// Replaces the value of all cells matching a specified value in one or more worksheets with a new value and
        /// updates their alignment and column width.
        /// </summary>
        /// <remarks>For each cell matching the specified value, the method sets the cell's value to the
        /// provided value, aligns the content vertically to center and horizontally to left, and adjusts the column
        /// width. If multiple worksheets are processed, all matching cells across those worksheets are
        /// updated.</remarks>
        /// <param name="pack">The ExcelPackage instance containing the workbook and worksheets to modify.</param>
        /// <param name="sheetName">The name of the worksheet to update. If null, all worksheets in the workbook are processed.</param>
        /// <param name="cellValue">The value to search for in cells. Only cells with this exact value are updated.</param>
        /// <param name="valueToInsert">The new value to insert into matching cells.</param>
        /// <returns>The ExcelPackage instance with updated cell values and formatting.</returns>
        public static ExcelPackage InsertToCells(this ExcelPackage pack, string sheetName, string cellValue, string valueToInsert)
        {
            var sheetList = pack.Workbook.Worksheets.ToList();
            if (sheetName != null)
                sheetList = new List<ExcelWorksheet>() { LoadSheet(pack, sheetName) };
            sheetList.ForEach(f =>
            {
                var cells = f.Cells.Where(c => c.Value != null && c.Value.ToString() == cellValue);
                foreach (var cell in cells)
                {
                    cell.Value = valueToInsert;
                    cell.Style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
                    cell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                }
            });
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
            return Clear(package, null);
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
            if (SheetName == null)
                package.Workbook.Worksheets.ToList().ForEach(f => f.Cells.Clear());
            else
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
            return LoadSheet(pack, null);
        }

        /// <summary>
        /// Loads specific Sheet from package
        /// </summary>
        /// <param name="pack"></param>
        /// <param name="SheetName">Specify SheetName</param>
        /// <returns>ExcelWorksheet</returns>
        private static ExcelWorksheet LoadSheet(ExcelPackage pack, string SheetName)
        {
            var activeSheet = pack.Workbook.Worksheets.FirstOrDefault();
            if (SheetName != null)
            {
                activeSheet = pack.Workbook.Worksheets.Where(w => w.Name == SheetName).FirstOrDefault();
            }
            if (activeSheet is null)
            {
                if (SheetName == null)
                    SheetName = "Sheet1";
                try
                {
                    activeSheet = pack.Workbook.Worksheets.Add(SheetName);
                }
                catch (System.InvalidOperationException ex)
                {
                    if (ex.Message.StartsWith("A worksheet with this name already exists in the workbook"))
                    {
                        pack.Workbook.Worksheets.Delete(SheetName);
                        activeSheet = pack.Workbook.Worksheets.Add(SheetName);
                    }
                    else
                        throw;
                }
            }

            return activeSheet;
        }

        private static int InsertProperties(ExcelWorksheet sheet, int rowFlag, int row, int startingCol, object item)
        {
            if (item == null)
                return row;

            //Takes the type of the first object
            var firstObjType = item.GetType();

            //Get all Properties from that type class
            var props = firstObjType.GetProperties();

            var bgColor = Color.FromArgb(13684430);

            for (int i = 0; i < props.Length; i++)
            {
                var colIndex = i + startingCol;
                var cell = sheet.Cells[row, colIndex];
                var header = props[i].Name ?? "null";

                cell.Value = header;

                var style = cell.Style;
                style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Left;
                style.Font.Bold = true;
                style.Font.Size = 12;
                style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                style.Fill.BackgroundColor.SetColor(bgColor);
                style.VerticalAlignment = OfficeOpenXml.Style.ExcelVerticalAlignment.Center;
            }
            //row += 1;
            sheet.Cells[row, startingCol, row, (props.Length + startingCol) - 1].AutoFilter = true;

            return row;
        }

    }
}
