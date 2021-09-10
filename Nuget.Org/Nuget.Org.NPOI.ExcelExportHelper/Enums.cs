namespace Nuget.Org.Npoi.ExcelExportHelper
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
        StyleForDecimal0Brackets,

        /// <summary>
        ///  数字型数据样式 保留0位 千分位  右对齐 （无括号，有负号）
        /// </summary>
        StyleForDecimal0NoColor,

        /// <summary>
        /// 数字型带货币符号$数据样式 保留2位 右对齐 （无括号，有负号）
        /// </summary>
        StyleWithUSDCurrencySymbol,

        /// <summary>
        /// 数字型带百分比数据样式 保留2位 右对齐 （无括号，有负号）
        /// </summary>
        StyleForPercent,

        /// <summary>
        /// 加粗 加粗左对齐
        /// </summary>
        StyleBoldLeft,

        /// <summary>
        /// 数字型数据样式 保留2位 千分位  右对齐 负数不变红（无括号，有负号）
        /// </summary>
        StyleNoRedForDecimal2,

        /// <summary>
        /// 数字型数据样式 保留2位 千分位  右对齐 负数不变红（有括号，无负号）
        /// </summary>
        StyleNoRedForDecimal2Brackets,

        /// <summary>
        /// 数字型带货币符号$数据样式 保留2位 右对齐负数不变红（无括号，有负号）
        /// </summary>
        StyleNoRedWithUSDCurrencySymbol,
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

        /// <summary>
        /// 浅橙色
        /// </summary>

        LightOrange,

        /// <summary>
        /// 浅褐色
        /// </summary>
        LightBrown,

        /// <summary>
        /// 浅黄色
        /// </summary>
        LightYellow
    }

    /// <summary>
    /// 边框设置
    /// </summary>
    public enum CellBorderEnums
    {
        /// <summary>
        /// 四面都是单线边框
        /// </summary>
        AllThin,

        /// <summary>
        ///  四面都是双线边框
        /// </summary>
        AllDouble,

        /// <summary>
        /// 右边框 双线
        /// </summary>
        RDouble,

        /// <summary>
        ///  底部边框 双线
        /// </summary>
        BDouble,

        /// <summary>
        ///  底部+右边边 双线
        /// </summary>
        BRDouble,
    }
}
