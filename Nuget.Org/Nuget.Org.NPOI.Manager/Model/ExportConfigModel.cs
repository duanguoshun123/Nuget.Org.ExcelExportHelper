using NPOI.SS.UserModel;

namespace Nuget.Org.Npoi.Manager.Model
{
    /// <summary>
    /// 导出配置
    /// </summary>
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

        private CellBackgroundColorEnums? fontColor;
        /// <summary>
        /// 单元格颜色样式
        /// </summary>
        public CellBackgroundColorEnums? FontColor { get { return fontColor; } set { fontColor = value; } }

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

    }


}
