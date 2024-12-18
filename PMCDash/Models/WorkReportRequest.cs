using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class WorkReportRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public float OPId { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public string Device { get; set; }
    }

    public class WorkReporStarttRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public float OPId { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 工單需求數量
        /// </summary>
        public int OrderQTY { get; set; }
    }

    public class WorkPassReportRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPId { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 回報生產數量
        /// </summary>
        public int QtyTol { get; set; }
        /// <summary>
        /// 不良品原因及數量
        /// </summary>
        public List<DefectiveReason> Reason { get; set; }
    }

    public class StartIdleReportRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderId { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string ReasonCode { get; set; }
    }

    public class EndIdleReportRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderId { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 設備編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string ReasonCode { get; set; }
    }

    /// <summary>
    /// 寫入量測資料
    /// </summary>
    public class WorkQCReportRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public float OPId { get; set; }
        /// <summary>
        /// 物料編號
        /// </summary>
        public string ProductId { get; set; }
        /// <summary>
        /// 檢驗點名稱
        /// </summary>
        public string QCPointName { get; set; }
        /// <summary>
        /// 檢驗工具編號
        /// </summary>
        public string QCToolID { get; set; }

        /// <summary>
        /// QC量測值
        /// </summary>
        public double QCValue { get; set; }

        /// <summary>
        /// 品檢模式:0=>自動 1=>手動
        /// </summary>
        public int QCMode { get; set; }

    }

    public class RequestReportInfo
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPId { get; set; }
        /// <summary>
        /// 報工回報數量
        /// </summary>
        public int QtyTol { get; set; }
        /// <summary>
        /// 不良品原因及數量
        /// </summary>
        public List<DefectiveReason> Reason { get; set; }
    }

    public class DefectiveReason
    {
        /// <summary>
        /// 不良原因
        /// </summary>
        public string QCReason { get; set; }
        /// <summary>
        /// 不量原因數量
        /// </summary>
        public int QtyNum { get; set; }
    }
}
