using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using Nuget.Org.Npoi.ExcelExportHelper.Model;
using Nuget.Org.Npoi.ExcelExportHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace Nuget.Org.Npoi.ExcelExportHelper.Npoi
{
    public static class NpoiUtils
    {
        public static MemoryStream ExportDataTableMultiple(List<Tuple<DataTable, string>> sheetTable, bool isOoxml)
        {
            IWorkbook workbook = CreateWorkbook(isOoxml);
            if (sheetTable?.Count > 0)
            {
                foreach (var item in sheetTable)
                {
                    CreateSheet(workbook, item.Item1, item.Item2);
                }
            }

            return WriteWorkbookToMemoryStream(workbook);
        }

        private static void CreateSheet(IWorkbook workbook, DataTable dataTable1, string dtName1)
        {
            ICellStyle style = workbook.CreateCellStyle();
            //设置单元格的样式：水平对齐居中
            style.Alignment = HorizontalAlignment.Center;
            //新建一个字体样式对象
            IFont font = workbook.CreateFont();
            //设置字体加粗样式
            font.Boldweight = short.MaxValue;
            //使用SetFont方法将字体样式添加到单元格样式中
            style.SetFont(font);

            ISheet sheet = workbook.CreateSheet(dtName1);
            IRow headerRow = sheet.CreateRow(0);

            // handling header.
            foreach (DataColumn column in dataTable1.Columns)
            {
                headerRow.CreateCell(column.Ordinal).SetCellValue(column.Caption);//If Caption not set, returns the ColumnName value
                headerRow.GetCell(column.Ordinal).CellStyle = style;
            }

            // handling value.
            int rowIndex = 1;

            foreach (DataRow row in dataTable1.Rows)
            {
                IRow dataRow = sheet.CreateRow(rowIndex);

                foreach (DataColumn column in dataTable1.Columns)
                {
                    if (row[column].GetType() == typeof(decimal)
                        || row[column].GetType() == typeof(int)
                        || row[column].GetType() == typeof(float)
                        || row[column].GetType() == typeof(double))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellNumberValue(row[column]);
                    }
                    else if (row[column].GetType() == typeof(string))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                    }
                    else if (row[column].GetType() == typeof(DateTime))
                    {
                        var value = (DateTime)row[column];
                        if (value != null)
                        {
                            if (value.Hour == 0 && value.Second == 0 && value.Minute == 0)
                            {
                                dataRow.CreateCell(column.Ordinal).SetCellValue(((DateTime)row[column]).ToString("yyyy-MM-dd"));
                            }
                            else
                            {
                                dataRow.CreateCell(column.Ordinal).SetCellValue(((DateTime)row[column]).ToString("yyyy-MM-dd HH:mm:ss"));
                            }
                        }
                        else
                        {
                            dataRow.CreateCell(column.Ordinal).SetCellValue(((DateTime)row[column]));
                        }
                    }
                    else if (row[column].GetType() == typeof(bool))
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue((bool)row[column]);
                    }
                    else
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                    }
                }

                rowIndex++;
            }
        }

        public static IWorkbook CreateWorkbook(bool isOoxml)
        {
            if (isOoxml)
            {
                return new NPOI.XSSF.UserModel.XSSFWorkbook();
            }
            else
            {
                return new NPOI.HSSF.UserModel.HSSFWorkbook();
            }
        }

        public static IWorkbook ReadWorkbook(bool isOoxml, Stream stream)
        {
            using (stream)
            {
                if (isOoxml)
                {
                    return new NPOI.XSSF.UserModel.XSSFWorkbook(stream);
                }
                else
                {
                    return new NPOI.HSSF.UserModel.HSSFWorkbook(stream);
                }
            }
        }

        public static IWorkbook ReadWorkbook(bool isOoxml, string path)
        {
            var stream = File.OpenRead(path);

            return ReadWorkbook(isOoxml, stream);
        }

        public static void WriteWorkbook(Stream stream, IWorkbook workbook, bool seekBegin = true)
        {
            workbook.Write(stream);
            if (seekBegin)
            {
                stream.Seek(0, SeekOrigin.Begin);
            }
        }

        public static MemoryStream WriteWorkbookToMemoryStream(IWorkbook workbook)
        {
            var ms = new NpoiMemoryStream();
            ms.AllowClose = false;
            WriteWorkbook(ms, workbook);
            ms.AllowClose = true;
            return ms;
        }

        public static byte[] WriteWorkbookToBytes(IWorkbook workbook)
        {
            using (var stream = WriteWorkbookToMemoryStream(workbook))
            {
                return stream.ToArray();
            }
        }

        public static void CopyCellValue(ICell srcCell, ICell dstCell)
        {
            switch (srcCell.CellType)
            {
                case CellType.Blank:
                    dstCell.SetCellValue(srcCell.StringCellValue);
                    break;

                case CellType.Boolean:
                    dstCell.SetCellValue(srcCell.BooleanCellValue);
                    break;

                case CellType.Error:
                    dstCell.SetCellErrorValue(srcCell.ErrorCellValue);
                    break;

                case CellType.Formula:
                    dstCell.SetCellFormula(srcCell.CellFormula);
                    break;

                case CellType.Numeric:
                    dstCell.SetCellValue(srcCell.NumericCellValue);
                    break;

                case CellType.String:
                    dstCell.SetCellValue(srcCell.RichStringCellValue);
                    break;

                case CellType.Unknown:
                    dstCell.SetCellValue(srcCell.StringCellValue);
                    break;
            }
        }

        public static void CopyAndOverwriteCell(IWorkbook workbook, ISheet srcSheet, ISheet dstSheet, int srcRowNum, int dstRowNum, int srcCellNum, int dstCellNum)
        {
            var srcRow = srcSheet.GetRow(srcRowNum);
            var srcCell = srcRow?.GetCell(srcCellNum);
            var dstRow = dstSheet.Row(dstRowNum);
            var dstCell = dstRow.Cell(dstCellNum);
            if (srcCell != null)
            {
                dstCell.CellStyle = srcCell.CellStyle;

                // If there is a cell comment, copy
                if (srcCell.CellComment != null)
                {
                    dstCell.CellComment = srcCell.CellComment;
                }

                // If there is a cell hyperlink, copy
                if (srcCell.Hyperlink != null)
                {
                    dstCell.Hyperlink = srcCell.Hyperlink;
                }

                // Set the cell data type
                dstCell.SetCellType(srcCell.CellType);

                // Set the cell data value
                CopyCellValue(srcCell, dstCell);
            }
            else
            {
                dstCell.SetCellValue((string)null);
            }
        }

        /// <summary>
        /// HSSFRow Copy Command
        ///
        /// Description:  Inserts a existing row into a new row, will automatically push down
        ///               any existing rows.  Copy is done cell by cell and supports, and the
        ///               command tries to copy all properties available (style, merged cells, values, etc...)
        /// </summary>
        public static void CopyRow(IWorkbook workbook, ISheet sheet, int srcRowNum, int dstRowNum)
        {
            CopyRow(workbook, sheet, sheet, srcRowNum, dstRowNum);
        }

        public static void CopyRow(IWorkbook workbook, ISheet srcSheet, ISheet dstSheet, int srcRowNum, int dstRowNum)
        {
            // Get the source / new row
            var srcRow = srcSheet.GetRow(srcRowNum);
            if (srcRow == null)
            {
                return;
            }
            var dstRow = dstSheet.GetRow(dstRowNum);

            // If the row exist in destination, push down all rows by 1 else create a new row
            if (dstRow != null)
            {
                dstSheet.ShiftRows(dstRowNum, dstSheet.LastRowNum, 1, true, false);
                dstRow = dstSheet.GetRow(dstRowNum);
            }

            if (dstRow == null)
            {
                dstRow = dstSheet.CreateRow(dstRowNum);
            }

            // Loop through source columns to add to new row
            for (int i = 0; i < srcRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                var srcCell = srcRow.GetCell(i);
                var dstCell = dstRow.CreateCell(i);

                // If the old cell is null jump to next cell
                if (srcCell == null)
                {
                    dstCell = null;
                    continue;
                }

                dstCell.CellStyle = srcCell.CellStyle;

                // If there is a cell comment, copy
                if (srcCell.CellComment != null)
                {
                    dstCell.CellComment = srcCell.CellComment;
                }

                // If there is a cell hyperlink, copy
                if (srcCell.Hyperlink != null)
                {
                    dstCell.Hyperlink = srcCell.Hyperlink;
                }

                // Set the cell data type
                dstCell.SetCellType(srcCell.CellType);

                // Set the cell data value
                CopyCellValue(srcCell, dstCell);
            }
            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < srcSheet.NumMergedRegions; i++)
            {
                var cellRangeAddress = srcSheet.GetMergedRegion(i);
                if (cellRangeAddress.FirstRow == srcRow.RowNum)
                {
                    var newCellRangeAddress = new CellRangeAddress(
                        dstRow.RowNum,
                        (dstRow.RowNum + (cellRangeAddress.FirstRow - cellRangeAddress.LastRow)),
                        cellRangeAddress.FirstColumn,
                        cellRangeAddress.LastColumn);
                    dstSheet.AddMergedRegion(newCellRangeAddress);
                }
            }
            if (srcRow.IsFormatted && srcRow.RowStyle != null)
            {
                dstRow.RowStyle = srcRow.RowStyle;
            }
            else if (dstRow.IsFormatted && dstRow.RowStyle != null)
            {
                dstRow.RowStyle = null;
            }
        }

        public static void CopyRows(IWorkbook workbook, ISheet sheet, int srcRowNum, int dstRowNum, int length)
        {
            for (int i = 0; i < length; i++)
            {
                CopyRow(workbook, sheet, srcRowNum + i, dstRowNum + i);
            }
        }

        public static void CopyRows(IWorkbook workbook, ISheet srcSheet, ISheet dstSheet, int srcRowNum, int dstRowNum, int length)
        {
            for (int i = 0; i < length; i++)
            {
                CopyRow(workbook, srcSheet, dstSheet, srcRowNum + i, dstRowNum + i);
            }
        }

        public static void DeleteRow(ISheet sheet, int rowNum)
        {
            if (rowNum >= 0 && rowNum <= sheet.LastRowNum)
            {
                if (rowNum == sheet.LastRowNum)
                {
                    var row = sheet.CreateRow(rowNum + 1);
                }
                sheet.ShiftRows(rowNum + 1, sheet.LastRowNum, -1, true, true);
            }
        }

        public static void DeleteRows(ISheet sheet, int rowNum, int length)
        {
            if (length <= 0)
            {
                return;
            }
            if (length == 1)
            {
                DeleteRow(sheet, rowNum);
                return;
            }
            if (rowNum >= 0 && rowNum + 2 * length < sheet.LastRowNum + 2)
            {
                sheet.ShiftRows(rowNum + length, sheet.LastRowNum, -length, true, true);
            }
            else if (length > 16)
            {
                int halflength = length / 2;
                DeleteRows(sheet, rowNum, halflength);
                DeleteRows(sheet, rowNum, length - halflength);
            }
            else
            {
                for (int i = 0; i < length; i++)
                {
                    DeleteRow(sheet, rowNum);
                }
            }
        }

        public static void CloneFont(IFont srcFont, IFont dstFont)
        {
            dstFont.Boldweight = srcFont.Boldweight;
            dstFont.Charset = srcFont.Charset;
            dstFont.Color = srcFont.Color;
            dstFont.FontHeight = srcFont.FontHeight;
            dstFont.FontHeightInPoints = srcFont.FontHeightInPoints;
            dstFont.FontName = srcFont.FontName;
            dstFont.IsItalic = srcFont.IsItalic;
            dstFont.IsStrikeout = srcFont.IsStrikeout;
            dstFont.TypeOffset = srcFont.TypeOffset;
            dstFont.Underline = srcFont.Underline;
        }

        public static int ExcelColumNameToColumnNumber(string excelColumName)
        {
            int columNumber = 0;
            if (string.IsNullOrEmpty(excelColumName))
            {
                throw new ArgumentOutOfRangeException("excelColumName");
            }
            excelColumName = excelColumName.ToUpper();
            for (int i = 0; i < excelColumName.Length; i++)
            {
                char c = excelColumName[i];
                if (c < 'A' || c > 'Z')
                {
                    throw new ArgumentOutOfRangeException("excelColumName");
                }
                columNumber = 26 * columNumber + (c - 'A' + 1);
            }
            return columNumber;
        }

        public static string ColumNumberToExcelColumName(int columNumber)
        {
            string excelColumName = string.Empty;
            if (columNumber <= 0)
            {
                throw new ArgumentOutOfRangeException("columNumber");
            }
            for (int n = columNumber; n > 0;)
            {
                int d = n % 26;
                int p = n / 26;
                if (d == 0)
                {
                    d = 26;
                    p--;
                }
                char c = (char)('A' + (d - 1));
                excelColumName = c + excelColumName;
                n = p;
            }
            return excelColumName;
        }

        public static void SetCellNumberValue(this ICell cell, object value)
        {
            if (value == null)
            {
                cell.SetCellValue((string)null);
            }
            else if (value is string)
            {
                cell.SetCellValue((string)value);
            }
            else if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
            }
            else
            {
                try
                {
                    cell.SetCellValue(Convert.ToDouble(value));
                }
                catch
                {
                    cell.SetCellValue((string)null);
                }
            }
        }

        public static IRow Row(this ISheet sheet, int index)
        {
            return sheet.GetRow(index) ?? sheet.CreateRow(index);
        }

        public static ICell Cell(this IRow row, int index)
        {
            return row.GetCell(index) ?? row.CreateCell(index);
        }

        public static IReadOnlyDictionary<string, IReadOnlyList<IReadOnlyList<string>>> ReadSheetStringCellValues(bool isOoxml, string path, IEnumerable<string> sheetNames)
        {
            var book = ReadWorkbook(isOoxml, path);
            return sheetNames.Distinct().ToDictionary(x => x, x =>
            {
                var sheet = book.GetSheet(x);
                return Enumerable.Range(sheet.FirstRowNum, sheet.LastRowNum - sheet.FirstRowNum + 1)
                    .Select(i => sheet.Row(i))
                    .Select(r => r.Cells.Select(c =>
                    {
                        switch (c.CellType)
                        {
                            case CellType.Unknown:
                                return c.StringCellValue;

                            case CellType.Numeric:
                                return c.NumericCellValue.ToString();

                            case CellType.String:
                                return c.StringCellValue;

                            case CellType.Formula:
                                {
                                    switch (c.CachedFormulaResultType)
                                    {
                                        case CellType.Unknown:
                                            return c.StringCellValue;

                                        case CellType.Numeric:
                                            return c.NumericCellValue.ToString();

                                        case CellType.String:
                                            return c.StringCellValue;

                                        case CellType.Formula:
                                            return c.CellFormula;

                                        case CellType.Blank:
                                            return c.StringCellValue;

                                        case CellType.Boolean:
                                            return c.BooleanCellValue.ToString();

                                        case CellType.Error:
                                            return c.ErrorCellValue.ToString();

                                        default:
                                            return c.StringCellValue;
                                    }
                                }

                            case CellType.Blank:
                                return c.StringCellValue;

                            case CellType.Boolean:
                                return c.BooleanCellValue.ToString();

                            case CellType.Error:
                                return c.ErrorCellValue.ToString();

                            default:
                                return c.StringCellValue;
                        }
                    }).ToList())
                    .ToList() as IReadOnlyList<IReadOnlyList<string>>;
            });
        }

        public static string GetCellEvaluatedValue(this ICell cell)
        {
            if (cell == null)
            {
                return string.Empty;
            }

            switch (cell.CellType)
            {
                case CellType.Blank:
                    return string.Empty;

                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();

                case CellType.Error:
                    return cell.ErrorCellValue.ToString();

                case CellType.Numeric:
                case CellType.Unknown:
                default:
                    return cell.ToString();//This is a trick to get the correct value of the cell. NumericCellValue will return a numeric value no matter the cell value is a date or a number

                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Formula:
                    try
                    {
                        IFormulaEvaluator e = new NPOI.XSSF.UserModel.XSSFFormulaEvaluator(cell.Sheet.Workbook);
                        e.EvaluateInCell(cell);
                        return cell.ToString();
                    }
                    catch
                    {
                        return cell.NumericCellValue.ToString();
                    }
            }
        }

        #region 自定义导出报表

        /// <summary>
        /// 导出 多sheet
        /// </summary>
        /// <param name="configModels"></param>
        /// <param name="isOoxml"></param>
        /// <returns></returns>
        public static MemoryStream ExportMultSheetWithCellFormat(Dictionary<string, ExportConfig> configModels, bool isOoxml)
        {
            IWorkbook workbook = CreateWorkbook(isOoxml);

            var cellStyles = configModels.SelectMany(x => x.Value.CellModelConfigs)
                .GroupBy(x => new
                {
                    x.CellStyle,
                    x.FontBackgroundColor,
                    x.BorderStyle,
                    x.IsNeedSetBorder
                })
                .Select(x =>
                {
                    var cellStyle = CreateCellStyle(workbook, new ExportCellStyleParameterModel
                    {
                        CellBackgroundColorEnums = x.Key.FontBackgroundColor,
                        CellStyleEnums = x.Key.CellStyle,
                        CellBorderEnums = x.Key.BorderStyle,
                        Font = x.FirstOrDefault().FontStyle
                    });

                    return new
                    {
                        CellStyleName = x.Key.CellStyle,
                        FontBackgroundColorName = x.Key.FontBackgroundColor,
                        Style = cellStyle,
                        x.Key.BorderStyle,
                        x.Key.IsNeedSetBorder
                    };
                }).ToList();

            configModels.ToList().ForEach(p =>
            {
                ISheet sheet = workbook.CreateSheet(p.Key);

                if (p.Value == null)
                {
                    return;
                }

                p.Value.CellModelConfigs.ForEach(x =>
                {
                    IRow row = sheet.GetRow(x.MergeRowIndexStart);
                    if (row == null)
                    {
                        row = sheet.CreateRow(x.MergeRowIndexStart);
                    }
                    ICell cell1 = row.CreateCell(x.MergeColIndexStart);

                    if (x.IsNeedNumberFormat)
                    {
                        double result;
                        double? doubleValue = double.TryParse((x.CellValue ?? "").ToString(), out result) ? (double?)result : null;
                        if (doubleValue.HasValue)
                        {
                            cell1.SetCellValue(doubleValue.Value);
                        }
                        else
                        {
                            cell1.SetCellValue(string.Empty);
                        }
                    }
                    else
                    {
                        cell1.SetCellValue(x.CellValue?.ToString());
                    }

                    cell1.CellStyle = cellStyles.Where(r => r.CellStyleName == x.CellStyle
                       && r.FontBackgroundColorName == x.FontBackgroundColor
                       && r.IsNeedSetBorder == x.IsNeedSetBorder
                       && r.BorderStyle == x.BorderStyle).FirstOrDefault().Style;

                    if (sheet.GetRow(x.MergeRowIndexEnd ?? x.MergeRowIndexStart) == null)
                    {
                        sheet.CreateRow(x.MergeRowIndexEnd ?? x.MergeRowIndexStart);
                    }
                    var merge = new CellRangeAddress(x.MergeRowIndexStart, x.MergeRowIndexEnd ?? x.MergeRowIndexStart, x.MergeColIndexStart, x.MergeColIndexEnd ?? x.MergeColIndexStart);

                    sheet.AddMergedRegion(merge);

                    if (x.IsNeedSetBorder)
                    {
                        // 需要设置边框 但是未定义边框 默认设置单线边框
                        if (x.BorderStyle == null)
                        {
                            RegionUtil.SetBorderBottom((int)BorderStyle.Thin, merge, sheet);
                            RegionUtil.SetBorderRight((int)BorderStyle.Thin, merge, sheet);
                            RegionUtil.SetBorderLeft((int)BorderStyle.Thin, merge, sheet);
                            RegionUtil.SetBorderTop((int)BorderStyle.Thin, merge, sheet);
                        }
                        else
                        {
                            RegionUtil.SetBorderBottom((int)cell1.CellStyle.BorderBottom, merge, sheet);
                            RegionUtil.SetBorderRight((int)cell1.CellStyle.BorderRight, merge, sheet);
                            RegionUtil.SetBorderLeft((int)cell1.CellStyle.BorderLeft, merge, sheet);
                            RegionUtil.SetBorderTop((int)cell1.CellStyle.BorderTop, merge, sheet);
                        }

                        //此处有时候会背景色会被替换掉 为了防止丢失 重新设置
                        if (x.FontBackgroundColor.HasValue)
                        {
                            SetFontbackgroundColorStyle(cell1.CellStyle, x.FontBackgroundColor.Value);
                        }
                    }
                });

                #region 单元格自适应宽度

                int colIndexWidth = p.Value.CellModelConfigs.Max(x => x.MergeColIndexEnd ?? x.MergeColIndexStart);

                for (int columnNum = 0; columnNum <= colIndexWidth; columnNum++)
                {
                    //sheet.AutoSizeColumn(columnNum);//先来个常规自适应

                    int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                    for (int rowNum = 0; rowNum < sheet.LastRowNum; rowNum++)
                    {
                        IRow currentRow;
                        //当前行未被使用过
                        if (sheet.GetRow(rowNum) == null)
                        {
                            currentRow = sheet.CreateRow(rowNum);
                        }
                        else
                        {
                            currentRow = sheet.GetRow(rowNum);
                        }

                        if (currentRow.GetCell(columnNum) != null)
                        {
                            ICell currentCell = currentRow.GetCell(columnNum);

                            int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                            if (columnWidth < length)
                            {
                                columnWidth = length;
                            }
                        }
                    }
                    sheet.SetColumnWidth(columnNum, columnWidth * 350);
                }

                #endregion 单元格自适应宽度

                sheet.ForceFormulaRecalculation = true;

                if (p.Value.IsNeedFreeZoneRowCol)
                {
                    if (p.Value.FreeZoneRowCol != null)
                    {
                        sheet.CreateFreezePane(p.Value.FreeZoneRowCol.Item1,
                            p.Value.FreeZoneRowCol.Item2,
                            p.Value.FreeZoneRowCol.Item3,
                            p.Value.FreeZoneRowCol.Item4);
                    }
                }
            });

            return WriteWorkbookToMemoryStream(workbook);
        }

        /// <summary>
        /// 导出 单sheet
        /// </summary>
        /// <param name="exportConfig">配置项</param>
        /// <param name="isOoxml"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static MemoryStream ExportWithCellFormat(ExportConfig exportConfig, bool isOoxml, string sheetName)
        {
            IWorkbook workbook = CreateWorkbook(isOoxml);
            ISheet sheet = workbook.CreateSheet(sheetName);

            if (exportConfig == null)
            {
                return WriteWorkbookToMemoryStream(workbook);
            }

            var cellStyles = exportConfig?.CellModelConfigs.GroupBy(x => new
            {
                x.CellStyle,
                x.FontBackgroundColor,
                x.BorderStyle,
                x.IsNeedSetBorder
            }).Select(x =>
            {
                var cellStyle = CreateCellStyle(workbook, new ExportCellStyleParameterModel
                {
                    CellBackgroundColorEnums = x.Key.FontBackgroundColor,
                    CellStyleEnums = x.Key.CellStyle,
                    CellBorderEnums = x.Key.BorderStyle,
                    Font = x.FirstOrDefault().FontStyle
                });

                return new
                {
                    CellStyleName = x.Key.CellStyle,
                    FontBackgroundColorName = x.Key.FontBackgroundColor,
                    Style = cellStyle,
                    x.Key.BorderStyle,
                    x.Key.IsNeedSetBorder
                };
            }).ToList();

            exportConfig?.CellModelConfigs.ForEach(x =>
            {
                IRow row = sheet.GetRow(x.MergeRowIndexStart);
                if (row == null)
                {
                    row = sheet.CreateRow(x.MergeRowIndexStart);
                }
                ICell cell1 = row.CreateCell(x.MergeColIndexStart);
                if (x.IsNeedNumberFormat)
                {
                    double result;
                    double? doubleValue = double.TryParse((x.CellValue ?? "").ToString(), out result) ? (double?)result : null;
                    if (doubleValue.HasValue)
                    {
                        cell1.SetCellValue(doubleValue.Value);
                    }
                    else
                    {
                        cell1.SetCellValue(string.Empty);
                    }
                }
                else
                {
                    cell1.SetCellValue(x.CellValue?.ToString());
                }

                cell1.CellStyle = cellStyles.Where(p => p.CellStyleName == x.CellStyle && p.FontBackgroundColorName == x.FontBackgroundColor && p.BorderStyle == x.BorderStyle && p.IsNeedSetBorder == x.IsNeedSetBorder)?.FirstOrDefault()?.Style;

                if (sheet.GetRow(x.MergeRowIndexEnd ?? x.MergeRowIndexStart) == null)
                {
                    sheet.CreateRow(x.MergeRowIndexEnd ?? x.MergeRowIndexStart);
                }

                var merge = new CellRangeAddress(x.MergeRowIndexStart, x.MergeRowIndexEnd ?? x.MergeRowIndexStart, x.MergeColIndexStart, x.MergeColIndexEnd ?? x.MergeColIndexStart);

  
                if (x.IsMerge)
                {
                    sheet.AddMergedRegion(merge);
                }

                if (x.IsNeedSetBorder)
                {
                    // 需要设置边框 但是未定义边框 默认设置单线边框
                    if (x.BorderStyle == null)
                    {
                        RegionUtil.SetBorderBottom((int)BorderStyle.Thin, merge, sheet);
                        RegionUtil.SetBorderRight((int)BorderStyle.Thin, merge, sheet);
                        RegionUtil.SetBorderLeft((int)BorderStyle.Thin, merge, sheet);
                        RegionUtil.SetBorderTop((int)BorderStyle.Thin, merge, sheet);
                    }
                    else
                    {
                        RegionUtil.SetBorderBottom((int)cell1.CellStyle.BorderBottom, merge, sheet);
                        RegionUtil.SetBorderRight((int)cell1.CellStyle.BorderRight, merge, sheet);
                        RegionUtil.SetBorderLeft((int)cell1.CellStyle.BorderLeft, merge, sheet);
                        RegionUtil.SetBorderTop((int)cell1.CellStyle.BorderTop, merge, sheet);
                    }

                    //此处有时候会背景色会被替换掉 为了防止丢失 重新设置
                    if (x.FontBackgroundColor.HasValue)
                    {
                        SetFontbackgroundColorStyle(cell1.CellStyle, x.FontBackgroundColor.Value);
                    }
                }
            });

            #region 单元格自适应宽度

            int colIndexWidth = exportConfig.CellModelConfigs.Select(x => x.MergeColIndexEnd ?? x.MergeColIndexStart).Max();

            for (int columnNum = 0; columnNum <= colIndexWidth; columnNum++)
            {
                //sheet.AutoSizeColumn(columnNum);//先来个常规自适应

                int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 0; rowNum < sheet.LastRowNum; rowNum++)
                {
                    IRow currentRow;
                    //当前行未被使用过
                    if (sheet.GetRow(rowNum) == null)
                    {
                        currentRow = sheet.CreateRow(rowNum);
                    }
                    else
                    {
                        currentRow = sheet.GetRow(rowNum);
                    }

                    if (currentRow.GetCell(columnNum) != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                        if (columnWidth < length)
                        {
                            columnWidth = length;
                        }
                    }
                }
                sheet.SetColumnWidth(columnNum, columnWidth * 350);
            }

            #endregion 单元格自适应宽度

            sheet.ForceFormulaRecalculation = true;

            //冻结设置
            if (exportConfig.IsNeedFreeZoneRowCol)
            {
                if (exportConfig?.FreeZoneRowCol != null)
                {
                    sheet.CreateFreezePane(exportConfig.FreeZoneRowCol.Item1,
                        exportConfig.FreeZoneRowCol.Item2,
                        exportConfig.FreeZoneRowCol.Item3,
                        exportConfig.FreeZoneRowCol.Item4);
                }
            }

            return WriteWorkbookToMemoryStream(workbook);
        }

        /// <summary>
        /// 创建样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cellStyleParameterModel">单元格样式配置参数</param>
        /// <returns></returns>
        private static ICellStyle CreateCellStyle(IWorkbook workbook, ExportCellStyleParameterModel cellStyleParameterModel)
        {
            if (cellStyleParameterModel == null)
            {
                return workbook.CreateCellStyle();
            }

            ExportCellFormatStyleHelper cellStyleHelper = new ExportCellFormatStyleHelper();

            ICellStyle cellStyle = cellStyleHelper.CreateCellStyle(workbook, cellStyleParameterModel);

            return cellStyle;
        }

        /// <summary>
        /// 设置背景色
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="color"></param>
        private static void SetFontbackgroundColorStyle(ICellStyle cellStyle, CellBackgroundColorEnums color)
        {
            cellStyle.FillPattern = FillPattern.SolidForeground;
            cellStyle.FillForegroundColor = 0;
            switch (color)
            {
                case CellBackgroundColorEnums.LightRed:
                    cellStyle.FillForegroundColor = HSSFColor.Red.Index;
                    break;

                case CellBackgroundColorEnums.LightGreen:
                    cellStyle.FillForegroundColor = HSSFColor.LightGreen.Index;
                    break;

                case CellBackgroundColorEnums.LightOrange:
                    ((XSSFColor)cellStyle.FillForegroundColorColor).SetRgb(new byte[] { 250, 191, 143 });
                    break;

                case CellBackgroundColorEnums.LightBrown:
                    ((XSSFColor)cellStyle.FillForegroundColorColor).SetRgb(new byte[] { 238, 236, 225 });
                    break;

                case CellBackgroundColorEnums.LightYellow:
                    ((XSSFColor)cellStyle.FillForegroundColorColor).SetRgb(new byte[] { 255, 255, 0 });
                    break;

                default:
                    break;
            }
        }

        #endregion 自定义导出报表
    }

    public class NpoiMemoryStream : MemoryStream
    {
        public NpoiMemoryStream()
        {
            // We always want to close streams by default to
            // force the developer to make the conscious decision
            // to disable it.  Then, they're more apt to remember
            // to re-enable it.  The last thing you want is to
            // enable memory leaks by default.  ;-)
            AllowClose = true;
        }

        public bool AllowClose { get; set; }

        public override void Close()
        {
            if (AllowClose)
            {
                base.Close();
            }
        }
    }
}
