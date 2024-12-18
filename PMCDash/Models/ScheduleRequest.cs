using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    /// <summary>
    /// 交期優先排程法必要輸入
    /// </summary>
    public class ScheduleDateRangeRequest
    {

        /// <summary>
        /// 選擇交期開始時間
        /// </summary>
        public string StartTime { get; set; }
        /// <summary>
        /// 選擇交期結束時間
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// 權重模式【A:推薦參數、B:自定參數、C:推薦參數+自定參數】
        /// </summary>
        public string Mode { get; set; }

        /// <summary>
        /// 勾選工單製程項目
        /// </summary>
        public List<SelectOrder> SelectOrders { get; set; }

        /// <summary>
        /// 權重資料
        /// </summary>
        public Weight weights { get; set; }
    }

    /// <summary>
    /// 機台優先排程法必要輸入
    /// </summary>
    public class ScheduleOrdersRequest
    {
        /// <summary>
        /// 勾選工單製程項目
        /// </summary>
        public List<SelectOrder> SelectOrders { get; set; }
    }

    public class ScheduleOneOrderRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
    }

    public class ScheduleEditPredictDateRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public int OPID { get; set; }
        /// <summary>
        /// 生管欲交日期
        /// </summary>
        public string AssignDate_PM { get; set; }
    }


    /// <summary>
    /// 重要工單
    /// </summary>
    public class ScheduleEditImportantOrderRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 是否重要工單【True:重要、False:不重要】
        /// </summary>
        public bool Important { get; set; }
    }

    /// <summary>
    /// 派單
    /// </summary>
    public class ScheduleDispatchWorkOrderRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Machine { get; set; }
        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime SatrtTime { get; set; }

    }

    public class ManualScheduleRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Machine { get; set; }
        /// <summary>
        /// 開始時間
        /// </summary>
        public string StartTime { get; set; }

    }

    public class DispatchWorkScheduleRequest
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }

    }

    /// <summary>
    /// 取的甘特圖tab資訊必要輸入
    /// </summary>
    public class GanttChartTabRequest
    {
        /// <summary>
        /// 主要排程模式
        /// </summary>
        public string Mode { get; set; }
        ///// <summary>
        ///// 選擇的模式
        ///// </summary>
        //public int FocusTab { get; set; }
    }

    public class CancelScheduleRequest
    {
        /// <summary>
        /// 排程結果代號
        /// </summary>
        public string Mode { get; set; }
    }

    public class SelectOrder
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
    }

    public class ScheduleMethodSelectionRequest
    {
        /// <summary>
        /// 排程結果代號
        /// </summary>
        public string Mode { get; set; }
    }

    public class DashboardSeqResquest
    {
        public string[][] data { get; set; }
    }

    public class MachineBreakdownRequest
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Machine { get; set; }
        /// <summary>
        /// 排程結果代號
        /// </summary>
        public string Mode { get; set; }
        /// <summary>
        /// 機台報修開始時間
        /// </summary>
        public string StartTime { get; set; }
        /// <summary>
        /// 機台報修預計結束時間
        /// </summary>
        public string EndTime { get; set; }
    }

    public class DrapItem
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string orderid { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string opid { get; set; }
        /// <summary>
        /// 拖曳後的新開始時間
        /// </summary>
        public string st { get; set; }
        /// <summary>
        /// 拖曳厚後的新結束時間
        /// </summary>
        public string et { get; set; }
    }
}
