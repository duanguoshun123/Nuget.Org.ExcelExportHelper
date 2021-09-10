using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;

namespace Nuget.Org.Npoi.ExcelExportHelper.Model
{
    public class ExportConfigModel
    {
        private bool isMerge;

        /// <summary>
        /// 是否合并
        /// </summary>
        public bool IsMerge { get { return isMerge; } set { isMerge = value; } }

        private int mergeRowIndexStart;

        /// <summary>
        /// 行开始索引（必填）
        /// </summary>
        public int MergeRowIndexStart { get { return mergeRowIndexStart; } set { mergeRowIndexStart = value; } }

        private int? mergeRowIndexEnd;

        /// <summary>
        /// 行结束索引（IsMerge 为true时 必须填，false可不填）
        /// </summary>
        public int? MergeRowIndexEnd { get { return mergeRowIndexEnd; } set { mergeRowIndexEnd = value; } }

        private int mergeColIndexStart;

        /// <summary>
        /// 列开始索引（必填）
        /// </summary>
        public int MergeColIndexStart { get { return mergeColIndexStart; } set { mergeColIndexStart = value; } }

        private int? mergeColIndexEnd;

        /// <summary>
        /// 列结束索引（IsMerge 为true时 必须填，false可不填）
        /// </summary>
        public int? MergeColIndexEnd { get { return mergeColIndexEnd; } set { mergeColIndexEnd = value; } }

        private CellStyleEnums cellStyle;

        /// <summary>
        /// 单元格格式
        /// </summary>
        public CellStyleEnums CellStyle { get { return cellStyle; } set { cellStyle = value; } }

        private CellBackgroundColorEnums? fontBackgroundColor;

        /// <summary>
        /// 单元格背景颜色样式
        /// </summary>
        public CellBackgroundColorEnums? FontBackgroundColor { get { return fontBackgroundColor; } set { fontBackgroundColor = value; } }

        private object cellValue;

        /// <summary>
        /// 单元格内容
        /// </summary>
        public object CellValue { get { return cellValue; } set { cellValue = value; } }

        private bool isNeedNumberFormat;

        /// <summary>
        /// 是否需要数字格式
        /// </summary>
        public bool IsNeedNumberFormat
        {
            get { return isNeedNumberFormat; }
            set { isNeedNumberFormat = value; }
        }

        private bool isNeedSetBorder { get; set; }

        /// <summary>
        /// 是否需要设置边框
        /// </summary>
        public bool IsNeedSetBorder
        {
            get { return isNeedSetBorder; }
            set { isNeedSetBorder = value; }
        }

        /// <summary>
        /// 边框样式
        /// </summary>
        public CellBorderEnums? BorderStyle { get; set; }

        /// <summary>
        /// 字体设置
        /// </summary>
        public IFont FontStyle { get; set; }

        /// <summary>
        /// 自定义数据格式
        /// </summary>
        public string DataFormatString { get; set; }
    }

    /// <summary>
    /// Excel导出配置
    /// </summary>
    public class ExportConfig
    {
        /// <summary>
        /// 单元格内容配置
        /// </summary>
        public List<ExportConfigModel> CellModelConfigs { get; set; }

        /// <summary>
        /// 是否需要冻结列行
        /// </summary>
        public bool IsNeedFreeZoneRowCol { get; set; }

        /// <summary>
        /// 冻结列行
        /// 冻结列数 冻结行数 可见列的列索引 可见行的行索引
        /// </summary>
        public Tuple<int, int, int, int> FreeZoneRowCol { get; set; }
    }
}
