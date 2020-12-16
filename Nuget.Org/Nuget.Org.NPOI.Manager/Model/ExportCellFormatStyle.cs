using NPOI.SS.UserModel;

namespace Nuget.Org.Npoi.Manager.Model
{
    public class ExportCellFormatStyle
    {
        public ExportCellFormatStyle(IWorkbook workBook)
        {
            this.workBook = workBook;
            dataFormat = workBook.CreateDataFormat();
            cellStyle = workBook.CreateCellStyle();
        }
        private IWorkbook workBook;
        /// <summary>
        /// 数据格式
        /// </summary>
        private IDataFormat dataFormat;
        /// <summary>
        /// 单元格样式
        /// </summary>
        private ICellStyle cellStyle;
        /// <summary>
        /// 加粗 居中
        /// </summary>
        public ICellStyle StyleBoldCenter
        {
            get
            {
                IFont fontTitle = workBook.CreateFont();
                fontTitle.IsBold = true;
                cellStyle.Alignment = HorizontalAlignment.Center;
                cellStyle.SetFont(fontTitle);
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.WrapText = true;
                return cellStyle;
            }
        }

        /// <summary>
        /// 居中
        /// </summary>
        public ICellStyle StyleNoBold
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Center;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.WrapText = true;
                return cellStyle;
            }
        }

        /// <summary>
        /// 非加粗 左对齐
        /// </summary>
        public ICellStyle StyleNoBoldLeft
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Left;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.WrapText = true;
                return cellStyle;
            }
        }

        /// <summary>
        /// 非加粗 右对齐
        /// </summary>
        public ICellStyle StyleNoBoldRight
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.WrapText = true;
                return cellStyle;
            }
        }
        /// <summary>
        /// 数字型数据样式 保留四位 千分位  右对齐 负数变红
        /// </summary>
        public ICellStyle StyleForDecimal4
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0.0000;[Red]-#,##0.0000");
                return cellStyle;
            }
        }

        /// <summary>
        /// 数字型数据样式 保留四位 千分位  右对齐 负数变红(有括号无负号)
        /// </summary>
        public ICellStyle StyleForDecimal4Brackets
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0.0000;[Red](#,##0.0000)");
                return cellStyle;
            }
        }

        /// <summary>
        /// 数字型数据样式 保留2位 千分位  右对齐 负数变红(有括号无负号)
        /// </summary>
        public ICellStyle StyleForDecimal2Brackets
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0.00;[Red](#,##0.00)");
                return cellStyle;
            }
        }

        /// <summary>
        /// 数字型数据样式 保留2位 千分位  右对齐 负数变红（无括号，有负号）
        /// </summary>
        public ICellStyle StyleForDecimal2
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0.00;[Red]-#,##0.00");
                return cellStyle;
            }
        }

        /// <summary>
        /// 数字型数据样式 保留0位 千分位  右对齐 负数变红（无括号，有负号）
        /// </summary>
        public ICellStyle StyleForDecimal0
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0;[Red]-#,##0");
                return cellStyle;
            }
        }

        /// <summary>
        /// 数字型数据样式 保留0位 千分位  右对齐 负数变红（有括号，无负号）
        /// </summary>
        public ICellStyle StyleForDecimal0Brackets
        {
            get
            {
                cellStyle.Alignment = HorizontalAlignment.Right;
                cellStyle.VerticalAlignment = VerticalAlignment.Center;
                cellStyle.DataFormat = dataFormat.GetFormat("#,##0;[Red](#,##0)");
                return cellStyle;
            }
        }
    }

}
