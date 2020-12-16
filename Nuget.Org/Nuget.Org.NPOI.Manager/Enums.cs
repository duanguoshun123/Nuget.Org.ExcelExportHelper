namespace Nuget.Org.Npoi.Manager
{
    /// <summary>
    /// 单元格样式枚举
    /// </summary>
    public enum CellStyleEnums
    {
        /// <summary>
        /// 加粗 居中
        /// </summary>
        StyleBoldCenter,
        /// <summary>
        /// 居中
        /// </summary>
        StyleNoBold,
        /// <summary>
        /// 非加粗 左对齐
        /// </summary>
        StyleNoBoldLeft,
        /// <summary>
        /// 非加粗 右对齐
        /// </summary>
        StyleNoBoldRight,
        /// <summary>
        /// 数字型数据样式 保留四位 千分位  右对齐 负数变红（无括号，有负号）
        /// </summary>
        StyleForDecimal4,
        /// <summary>
        /// 数字型数据样式 保留四位 千分位  右对齐 负数变红（有括号，无负号）
        /// </summary>
        StyleForDecimal4Brackets,
        /// <summary>
        /// 数字型数据样式 保留2位 千分位  右对齐 负数变红（无括号，有负号）
        /// </summary>
        StyleForDecimal2,
        /// <summary>
        /// 数字型数据样式 保留2位 千分位  右对齐 负数变红（有括号，无负号）
        /// </summary>
        StyleForDecimal2Brackets,
        /// <summary>
        ///  数字型数据样式 保留0位 千分位  右对齐 负数变红（无括号，有负号）
        /// </summary>
        StyleForDecimal0,
        /// <summary>
        ///  数字型数据样式 保留0位 千分位  右对齐 负数变红（有括号，无负号）
        /// </summary>
        StyleForDecimal0Brackets
    }

    /// <summary>
    /// 单元格背景色
    /// </summary>
    public enum CellBackgroundColorEnums : short
    {
        /// <summary>
        /// 浅红色
        /// </summary>
        LightRed,
        /// <summary>
        /// 浅绿色
        /// </summary>
        LightGreen,
    }
}
