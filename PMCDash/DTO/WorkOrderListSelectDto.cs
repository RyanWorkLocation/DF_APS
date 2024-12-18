using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.DTO
{
    public class WorkOrderListSelectDto
    {
    }

    public class WorkOrderDetailDto
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderNo { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public int OPNo { get; set; }

        /// <summary>
        /// 製程名稱
        /// </summary>
        public string OPName { get; set; }

        /// <summary>
        /// 產品名稱
        /// </summary>
        public string ProductNo { get; set; }

        /// <summary>
        /// 分發機台
        /// </summary>
        public string Machine { get; set; }

        /// <summary>
        /// 工單數量
        /// </summary>
        public int RequireCount { get; set; }

        /// <summary>
        /// 當前數量
        /// </summary>
        public int CurrentCount { get; set; }

        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StarTime { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 預交日期
        /// </summary>
        public DateTime DueDate { get; set; }

        /// <summary>
        /// 延遲時數
        /// </summary>
        public DateTime DeleyTime { get; set; }

        /// <summary>
        /// 生管預交日
        /// </summary>
        public DateTime PlanningDueDate { get; set; }

        /// <summary>
        /// 生產進度
        /// </summary>
        public double ProgressBarRate { get; set; }

        /// <summary>
        /// 負責群組
        /// </summary>
        public string Group { get; set; }

        /// <summary>
        /// 負責人(操作人員)
        /// </summary>
        public string Operator { get; set; }

        /// <summary>
        /// 工單狀態
        /// </summary>
        public string WorkOrderStatus { get; set; }

        /// <summary>
        /// 設計圖製路徑
        /// </summary>
        public string CADFilePath { get; set; }

        /// <summary>
        /// 訂單交付日期
        /// </summary>
        public double OrderDueDate { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
    }


}
