using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Nuget.Org.Npoi.ExcelExportHelper.Model
{
    public class ExportCellFormatStyleHelper
    {
        public ExportCellFormatStyleHelper()
        {
        }

        private IFont Font { get; set; }

        private IDataFormat DataFormat { get; set; }

        public ICellStyle CreateCellStyle(IWorkbook workBook, ExportCellStyleParameterModel cellStyleParameterModel)
        {
            ICellStyle _cellStyle = workBook.CreateCellStyle();

            if (DataFormat == null)
            {
                DataFormat = workBook.CreateDataFormat();
            }

            if (Font == null)
            {
                Font = workBook.CreateFont();
            }

            // 设置单元格字体
            SetCellFont(_cellStyle, cellStyleParameterModel?.Font);

            // 获取内置样式
            GetCellStyleByCellStyleEnums(_cellStyle, cellStyleParameterModel?.CellStyleEnums);

            //设置边框样式
            SetCellBorderStyle(_cellStyle, cellStyleParameterModel?.CellBorderEnums);

            //设置背景色
            SetFontbackgroundColorStyle(_cellStyle, cellStyleParameterModel?.CellBackgroundColorEnums);

            // 设置单元格自定义数据格式
            SetDataFormat(_cellStyle, cellStyleParameterModel?.DataFormatString);

            return _cellStyle;
        }

        private void GetCellStyleByCellStyleEnums(ICellStyle cellStyle, CellStyleEnums? cellStyleEnums)
        {
            if (!cellStyleEnums.HasValue)
            {
                return;
            }

            switch (cellStyleEnums)
            {
                case CellStyleEnums.StyleBoldCenter:
                    Font.IsBold = true;
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.SetFont(Font);
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    break;

                case CellStyleEnums.StyleNoBold:
                    cellStyle.Alignment = HorizontalAlignment.Center;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    break;

                case CellStyleEnums.StyleNoBoldLeft:
                    cellStyle.Alignment = HorizontalAlignment.Left;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    break;

                case CellStyleEnums.StyleNoBoldRight:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    break;

                case CellStyleEnums.StyleForDecimal4:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0.0000;[Red]-#,##0.0000");
                    break;

                case CellStyleEnums.StyleForDecimal4Brackets:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0.0000;[Red](#,##0.0000)");
                    break;

                case CellStyleEnums.StyleForDecimal2:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0.00;[Red]-#,##0.00");
                    break;

                case CellStyleEnums.StyleForDecimal2Brackets:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0.00;[Red](#,##0.00)");
                    break;

                case CellStyleEnums.StyleForDecimal0:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0;[Red]-#,##0");
                    break;

                case CellStyleEnums.StyleForDecimal0Brackets:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0;[Red](#,##0)");
                    break;

                case CellStyleEnums.StyleForDecimal0NoColor:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0;-#,##0");
                    break;

                case CellStyleEnums.StyleWithUSDCurrencySymbol:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("$#,##0.00;[Red]$-#,##0.00");
                    break;

                case CellStyleEnums.StyleForPercent:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("0.00%;[Red]-0.00%");
                    break;

                case CellStyleEnums.StyleBoldLeft:
                    Font.IsBold = true;
                    cellStyle.Alignment = HorizontalAlignment.Left;
                    cellStyle.SetFont(Font);
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.WrapText = true;
                    break;

                case CellStyleEnums.StyleNoRedForDecimal2:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0.00");
                    break;

                case CellStyleEnums.StyleNoRedForDecimal2Brackets:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("#,##0.00;(#,##0.00)");
                    break;

                case CellStyleEnums.StyleNoRedWithUSDCurrencySymbol:
                    cellStyle.Alignment = HorizontalAlignment.Right;
                    cellStyle.VerticalAlignment = VerticalAlignment.Center;
                    cellStyle.DataFormat = DataFormat.GetFormat("$#,##0.00");
                    break;

                default:
                    // 默认自动创建的单元格样式
                    break;
            }
        }

        /// <summary>
        /// 设置边框样式
        /// </summary>
        /// <param name="cellStyle">单元格样式</param>
        /// <param name="cellBorderEnums">单元格边框样式枚举值</param>
        private void SetCellBorderStyle(ICellStyle cellStyle, CellBorderEnums? cellBorderEnums)
        {
            if (!cellBorderEnums.HasValue)
            {
                return;
            }
            switch (cellBorderEnums)
            {
                case CellBorderEnums.AllDouble:
                    cellStyle.BorderBottom = BorderStyle.Double;
                    cellStyle.BorderLeft = BorderStyle.Double;
                    cellStyle.BorderRight = BorderStyle.Double;
                    cellStyle.BorderTop = BorderStyle.Double;
                    ; break;
                case CellBorderEnums.AllThin:
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    ; break;
                case CellBorderEnums.BDouble:
                    cellStyle.BorderBottom = BorderStyle.Double;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderRight = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    ; break;
                case CellBorderEnums.BRDouble:
                    cellStyle.BorderBottom = BorderStyle.Double;
                    cellStyle.BorderRight = BorderStyle.Double;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    ; break;
                case CellBorderEnums.RDouble:
                    cellStyle.BorderRight = BorderStyle.Double;
                    cellStyle.BorderBottom = BorderStyle.Thin;
                    cellStyle.BorderLeft = BorderStyle.Thin;
                    cellStyle.BorderTop = BorderStyle.Thin;
                    ; break;
                default:
                    cellStyle.BorderBottom = BorderStyle.None;
                    cellStyle.BorderLeft = BorderStyle.None;
                    cellStyle.BorderTop = BorderStyle.None;
                    cellStyle.BorderRight = BorderStyle.None;
                    break;
            }
        }

        /// <summary>
        /// 设置背景色
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="color"></param>
        private void SetFontbackgroundColorStyle(ICellStyle cellStyle, CellBackgroundColorEnums? color)
        {
            if (!color.HasValue)
            {
                return;
            }

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

        /// <summary>
        /// 设置数据格式
        /// </summary>
        /// <param name="cellStyle">单元格样式</param>
        /// <param name="dataFormatString">自定义数据格式</param>
        private void SetDataFormat(ICellStyle cellStyle, string dataFormatString)
        {
            if (dataFormatString == null || dataFormatString == "")
            {
                return;
            }
            cellStyle.DataFormat = DataFormat.GetFormat(dataFormatString);
        }

        /// <summary>
        /// 设置单元格字体样式
        /// </summary>
        /// <param name="cellStyle"></param>
        /// <param name="font"></param>
        private void SetCellFont(ICellStyle cellStyle, IFont font)
        {
            if (font == null)
            {
                return;
            }

            Font.FontName = font.FontName;

            cellStyle.SetFont(Font);
        }
    }

    /// <summary>
    /// 单元格边框样式
    /// </summary>
    public class Font : IFont
    {
        public string FontName { get; set; }
        public double FontHeight { get; set; }
        public short FontHeightInPoints { get; set; }
        public bool IsItalic { get; set; }
        public bool IsStrikeout { get; set; }
        public short Color { get; set; }
        public FontSuperScript TypeOffset { get; set; }
        public FontUnderlineType Underline { get; set; }
        public short Charset { get; set; }

        public short Index { get; set; }

        public short Boldweight { get; set; }
        public bool IsBold { get; set; }
        double IFont.FontHeightInPoints { get => throw new System.NotImplementedException(); set => throw new System.NotImplementedException(); }

        void IFont.CloneStyleFrom(IFont src)
        {
            throw new System.NotImplementedException();
        }
    }

    /// <summary>
    /// 导出Excel单元格样式参数
    /// </summary>
    public class ExportCellStyleParameterModel
    {
        /// <summary>
        /// 单元格内容显示样式类型
        /// </summary>
        public CellStyleEnums? CellStyleEnums { get; set; }

        /// <summary>
        /// 单元格边框样式类型
        /// </summary>
        public CellBorderEnums? CellBorderEnums { get; set; }

        /// <summary>
        /// 单元格字体样式设置
        /// </summary>
        public IFont Font { get; set; }

        /// <summary>
        /// 数据样式
        /// </summary>
        public string DataFormatString { get; set; }

        /// <summary>
        /// 单元格背景色
        /// </summary>
        public CellBackgroundColorEnums? CellBackgroundColorEnums { get; set; }
    }
}
