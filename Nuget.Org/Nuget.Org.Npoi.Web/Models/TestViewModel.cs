using System;
using System.Collections.Generic;

namespace Nuget.Org.Npoi.Web.Models
{
    public class TestViewModel
    {
        /// <summary>
        /// 品类
        /// </summary>
        public string CommodityTypeName { get; set; }
        /// <summary>
        /// 详情
        /// </summary>
        public List<DetailViewModel> Details { get; set; }

    }
    public class DetailViewModel
    {
        /// <summary>
        /// 日期
        /// </summary>
        public DateTime AccountingDate { get; set; }
        /// <summary>
        /// 持仓手数
        /// </summary>
        public decimal? Position { get; set; }
        /// <summary>
        /// 未结盈亏
        /// </summary>
        public decimal? PnL { get; set; }
    }
}