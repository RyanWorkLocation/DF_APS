using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class AwaitReason
    {
        public AwaitReason(string _AwaitReasonNo,string _AwaitReasonName,string _AwaitReasonType)
        {
            AwaitReasonNo = _AwaitReasonNo;
            AwaitReasonName = _AwaitReasonName;
            AwaitReasonType = _AwaitReasonType;
        }

        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string AwaitReasonNo { get; set; }

        /// <summary>
        /// 待機原因名稱
        /// </summary>
        public string AwaitReasonName { get; set; }

        /// <summary>
        /// 待機原因類別
        /// </summary>
        public string AwaitReasonType { get; set; }
    }

    public class AwaitReasonReport
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string WorkOrderNo { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public int OPNo { get; set; }
        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string AwaitReasonNo { get; set; }

        /// <summary>
        /// 待機原因名稱
        /// </summary>
        public string AwaitReasonName { get; set; }

        /// <summary>
        /// 待機原因類別
        /// </summary>
        public string AwaitReasonType { get; set; }

        /// <summary>
        /// 時間戳記
        /// </summary>
        public DateTime AwaitReasonTimeStamp { get; set; }
    }
}
