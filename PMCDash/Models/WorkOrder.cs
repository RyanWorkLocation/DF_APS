using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class WorkOrder
    {
        //public WorkOrder(string wONo, int progressValue, string materialNo, int requireCount, string dueDatePM, string dueDate, List<Process> processInfos)
        //{
        //    WONo = wONo;
        //    ProgressValue = progressValue;
        //    MaterialNo = materialNo;
        //    RequireCount = requireCount;
        //    DueDatePM = dueDatePM;
        //    DueDate = dueDate;
        //    ProcessInfos = processInfos;
        //}

        /// <summary>
        /// 訂單編號
        /// </summary>
        public string OrderNo { get; set; }

        /// <summary>
        /// 工單編號
        /// </summary>
        public string WONo { get; set; }

        /// <summary>
        /// 生產進度
        /// </summary>
        public int ProgressValue { get; set; }

        /// <summary>
        /// 物料編號
        /// </summary>
        public string MaterialNo { get; set; }

        /// <summary>
        /// 物料名稱
        /// </summary>
        public string PartName { get; set; }

        /// <summary>
        /// 客戶圖號
        /// </summary>
        public string DrawingNname { get; set; }
        
        /// <summary>
        /// 處理方法
        /// </summary>
        public string ProcessMethod { get; set; }

        /// <summary>
        /// 品項材質
        /// </summary>
        public string ProductMaterial { get; set; }

        /// <summary>
        /// 對應機型
        /// </summary>
        public string CusDevice { get; set; }

        /// <summary>
        /// 工單備註
        /// </summary>
        public string Noet { get; set; }

        /// <summary>
        /// 需求數量
        /// </summary>
        public int RequireCount { get; set; }

        /// <summary>
        /// 生管預交日
        /// </summary>
        public string DueDatePM { get; set; }

        /// <summary>
        /// 預交日
        /// </summary>
        public string DueDate { get; set; }

        /// <summary>
        /// 尚餘加工天數
        /// </summary>
        public string NeedDay { get; set; }

        /// <summary>
        /// 工單製程明細
        /// </summary>
        public List<Process> ProcessInfos { get; set; }

        /// <summary>
        /// 排程估計交期
        /// </summary>
        public string EstimatedDate { get; set; }


    }

    public class Process
    {
        //public Process(string no, string name, int progressValue, string cPKEvaluation, string remark)
        //{
        //    No = no;
        //    Name = name;
        //    ProgressValue = progressValue;
        //    CPKEvaluation = cPKEvaluation;
        //    Remark = remark;
        //}

        /// <summary>
        /// 製程編號
        /// </summary>
        public string No { get; set; }

        /// <summary>
        /// 製程順序
        /// </summary>
        public int Range { get; set; }

        /// <summary>
        /// 製程名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 生產進度
        /// </summary>
        public int ProgressValue { get; set; }

        /// <summary>
        /// CPK評價
        /// </summary>
        public string CPKEvaluation { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }

        /// <summary>
        /// 生產時間
        /// </summary>
        public string ExecutionTime { get; set; }

        /// <summary>
        /// 生產逾時
        /// </summary>
        public bool IsDelay { get; set; }

    }

    public class ordercount
    {
        public string orderid { get; set; }
        public string opid { get; set; }
        public int range { get; set; }
        public string DeviceGroup { get; set; }
        public string workgroup { get; set; }
        public string Operator { get; set; }
        public string wipevent { get; set; }
    }

    public class DataICompate : IEqualityComparer<ordercount>
    {
        public bool Equals(ordercount x, ordercount y)
        {
            return x.orderid == y.orderid 
                && x.opid ==y.opid
                && x.range == y.range;
        }

        public int GetHashCode(ordercount obj)
        {
            return obj.orderid.GetHashCode();
        }
    }

}
