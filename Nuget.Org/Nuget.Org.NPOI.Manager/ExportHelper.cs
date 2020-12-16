using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using Nuget.Org.Npoi.Manager.Model;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
namespace Nuget.Org.Npoi.Manager
{
    public static class ExportHelper
    {
        /// <summary>
        /// 创建工作薄
        /// </summary>
        /// <param name="isOoxml"></param>
        /// <returns></returns>
        private static IWorkbook CreateWorkbook(bool isOoxml)
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
        /// <summary>
        /// 将工作薄写入文件流
        /// </summary>
        /// <param name="workbook"></param>
        /// <returns></returns>
        private static MemoryStream WriteWorkbookToMemoryStream(IWorkbook workbook)
        {
            var ms = new NpoiMemoryStream();
            ms.AllowClose = false;
            WriteWorkbook(ms, workbook);
            ms.AllowClose = true;
            return ms;
        }

        private static void WriteWorkbook(Stream stream, IWorkbook workbook, bool seekBegin = true)
        {
            workbook.Write(stream);
            if (seekBegin)
            {
                stream.Seek(0, SeekOrigin.Begin);
            }
        }

        private static byte[] WriteWorkbook(IWorkbook workbook)
        {
            using (var stream = new MemoryStream())
            {
                workbook.Write(stream);
                return stream.ToArray();
            }
        }


        /// <summary>
        /// 导出 多sheet
        /// </summary>
        /// <param name="configModels"></param>
        /// <param name="isOoxml"></param>
        /// <returns></returns>
        public static MemoryStream ExportMultSheet(Dictionary<string, List<ExportConfigModel>> configModels, bool isOoxml)
        {
            IWorkbook workbook = CreateWorkbook(isOoxml);

            configModels.ToList().ForEach(p =>
            {
                ISheet sheet = workbook.CreateSheet(p.Key);
                p.Value.ForEach(x =>
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
                    cell1.CellStyle = MappingCellStyle(workbook, x.CellStyle);
                    if (x.IsMerge)
                    {
                        if (sheet.GetRow(x.MergeRowIndexEnd.Value) == null)
                        {
                            sheet.CreateRow(x.MergeRowIndexEnd.Value);
                        }
                        sheet.AddMergedRegion(new CellRangeAddress(x.MergeRowIndexStart, x.MergeRowIndexEnd.Value, x.MergeColIndexStart, x.MergeColIndexEnd.Value));
                    }
                });

                #region 单元格自适应宽度
                for (int col = 0; col <= p.Value.Select(x => x.MergeColIndexEnd ?? x.MergeColIndexStart).Max(); col++)
                {
                    sheet.AutoSizeColumn(col);//自适应宽度，但是其实还是比实际文本要宽
                    int columnWidth = sheet.GetColumnWidth(col) / 256;//获取当前列宽度
                    for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        ICell cell = row.GetCell(col);
                        int contextLength = cell == null ? columnWidth : Encoding.UTF8.GetBytes(cell.ToString()).Length;//获取当前单元格的内容宽度
                        columnWidth = columnWidth < contextLength ? contextLength : columnWidth;
                    }
                    sheet.SetColumnWidth(col, columnWidth * 256);
                }
                #endregion

                sheet.ForceFormulaRecalculation = true;
            });


            return WriteWorkbookToMemoryStream(workbook);

        }


        /// <summary>
        /// 导出 单sheet
        /// </summary>
        /// <param name="configModels"></param>
        /// <param name="isOoxml"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static MemoryStream Export(List<ExportConfigModel> configModels, bool isOoxml, string sheetName)
        {
            IWorkbook workbook = CreateWorkbook(isOoxml);
            ISheet sheet = workbook.CreateSheet(sheetName);

            configModels.ForEach(x =>
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
                cell1.CellStyle = MappingCellStyle(workbook, x.CellStyle);
                if (x.FontColor.HasValue)
                {
                    SetFontColorStyle(cell1.CellStyle, x.FontColor.Value);
                }
                if (x.IsMerge)
                {
                    if (sheet.GetRow(x.MergeRowIndexEnd ?? x.MergeColIndexStart) == null)
                    {
                        sheet.CreateRow(x.MergeRowIndexEnd ?? x.MergeColIndexStart);
                    }
                    sheet.AddMergedRegion(new CellRangeAddress(x.MergeRowIndexStart, x.MergeRowIndexEnd ?? x.MergeRowIndexStart, x.MergeColIndexStart, x.MergeColIndexEnd ?? x.MergeColIndexStart));
                }
            });

            #region 单元格自适应宽度
            for (int col = 0; configModels.Count > 0 && col <= configModels.Max(x => x.MergeColIndexEnd ?? x.MergeColIndexStart); col++)
            {
                sheet.AutoSizeColumn(col);//自适应宽度，但是其实还是比实际文本要宽
                int columnWidth = sheet.GetColumnWidth(col) / 256;//获取当前列宽度
                for (int rowIndex = 0; rowIndex <= sheet.LastRowNum; rowIndex++)
                {
                    IRow row = sheet.GetRow(rowIndex);
                    ICell cell = row?.GetCell(col);
                    int contextLength = cell == null ? columnWidth : Encoding.UTF8.GetBytes(cell.ToString()).Length;//获取当前单元格的内容宽度
                    columnWidth = columnWidth < contextLength ? contextLength : columnWidth;
                }
                sheet.SetColumnWidth(col, columnWidth * 256);
            }
            #endregion

            sheet.ForceFormulaRecalculation = true;
            return WriteWorkbookToMemoryStream(workbook);

        }

        /// <summary>
        /// 获取样式
        /// </summary>
        /// <param name="workbook"></param>
        /// <param name="cellStyleEnums"></param>
        /// <returns></returns>
        private static ICellStyle MappingCellStyle(IWorkbook workbook, CellStyleEnums cellStyleEnums)
        {
            ExportCellFormatStyle exportCellFormatStyle = new ExportCellFormatStyle(workbook);
            switch (cellStyleEnums)
            {
                case CellStyleEnums.StyleBoldCenter:
                    return exportCellFormatStyle.StyleBoldCenter;
                case CellStyleEnums.StyleNoBold:
                    return exportCellFormatStyle.StyleNoBold;
                case CellStyleEnums.StyleNoBoldLeft:
                    return exportCellFormatStyle.StyleNoBoldLeft;
                case CellStyleEnums.StyleNoBoldRight:
                    return exportCellFormatStyle.StyleNoBoldRight;
                case CellStyleEnums.StyleForDecimal4:
                    return exportCellFormatStyle.StyleForDecimal4;
                case CellStyleEnums.StyleForDecimal4Brackets:
                    return exportCellFormatStyle.StyleForDecimal4Brackets;
                case CellStyleEnums.StyleForDecimal2:
                    return exportCellFormatStyle.StyleForDecimal2;
                case CellStyleEnums.StyleForDecimal2Brackets:
                    return exportCellFormatStyle.StyleForDecimal2Brackets;
                case CellStyleEnums.StyleForDecimal0:
                    return exportCellFormatStyle.StyleForDecimal0;
                case CellStyleEnums.StyleForDecimal0Brackets:
                    return exportCellFormatStyle.StyleForDecimal0Brackets;
                default:
                    // 默认居中非加粗
                    return exportCellFormatStyle.StyleNoBold;
            }
        }

        private static void SetFontColorStyle(ICellStyle cellStyle, CellBackgroundColorEnums color)
        {
            cellStyle.FillPattern = FillPattern.SolidForeground;
            switch (color)
            {
                case CellBackgroundColorEnums.LightRed:
                    cellStyle.FillForegroundColor = HSSFColor.Red.Index;
                    break;
                case CellBackgroundColorEnums.LightGreen:
                    cellStyle.FillForegroundColor = HSSFColor.LightGreen.Index;
                    break;
                default:
                    break;
            }
        }

    }
}
