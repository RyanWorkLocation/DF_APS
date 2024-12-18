using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class ReporterInformation
    {
        /// <summary>
        /// 機台名稱
        /// </summary>
        public string DeviceName { get; set; }

        /// <summary>
        /// 操作人員
        /// </summary>
        public string Operator { get; set; }

        /// <summary>
        /// 工單號
        /// </summary>
        public string OrderID { get; set; }

        /// <summary>
        /// 料號
        /// </summary>
        public string Maktxt { get; set; }

        /// <summary>
        /// 批號數量
        /// </summary>
        public int OringinCount { get; set; }

        /// <summary>
        /// 報工數量
        /// </summary>
        public int RepotedConut { get; set; }

        /// <summary>
        /// 工單狀態
        /// </summary>
        public string OrderStatus { get; set; }
    }

    public class BadReportInfo
    {
        /// <summary>
        /// 不良原因
        /// </summary>
        public string BadReason { get; set; }
        /// <summary>
        /// 不良品數量
        /// </summary>
        public int BadCount { get; set; }
    }

    public class idlereason
    {
        /// <summary>
        /// 待機原因標號
        /// </summary>
        public string idleReasonId { get; set; }
        /// <summary>
        /// 待機原因種類【1:人員、2:機台】
        /// </summary>
        public string idleReasonType { get; set; }
        /// <summary>
        /// 待機原因名稱
        /// </summary>
        public string idleReasonTitle { get; set; }
    }

    public class idleList
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string orderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string opId { get; set; }
        /// <summary>
        /// 待機種類【1:人員、2:機台】
        /// </summary>
        public string idletype { get; set; }
        /// <summary>
        /// 待機設備名稱
        /// </summary>
        public string idleDevice { get; set; }
        /// <summary>
        /// 待機人員名稱
        /// </summary>
        public string idlePerson { get; set; }
        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string idlereson_id { get; set; }
        /// <summary>
        /// 待機原因名稱
        /// </summary>
        public string idlereson_name { get; set; }
        /// <summary>
        /// 待機開始時間
        /// </summary>
        public string idle_start_time { get; set; }
        /// <summary>
        /// 待機結束時間
        /// </summary>
        public string idle_end_time { get; set; }
        /// <summary>
        /// 待機時間
        /// </summary>
        public string idleDuration { get; set; }
    }

    public class idleList_1
    {
        /// <summary>
        /// 待機種類【1:人員、2:機台】
        /// </summary>
        public string idletype { get; set; }
        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string idlereson_id { get; set; }
        /// <summary>
        /// 待機原因名稱
        /// </summary>
        public string idlereson_name { get; set; }
        /// <summary>
        /// 待機開始時間
        /// </summary>
        public string idle_start_time { get; set; }
        /// <summary>
        /// 待機結束時間
        /// </summary>
        public string idle_end_time { get; set; }
        /// <summary>
        /// 待機結束時間
        /// </summary>
        public string idle_first_report_time { get; set; }
        /// <summary>
        /// 待機結束時間
        /// </summary>
        public string idle_scend_report_time { get; set; }

    }

}
