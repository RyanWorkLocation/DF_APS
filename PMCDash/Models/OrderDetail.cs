using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class OrderDetail
    {
        /// <summary>
        /// 訂單編號
        /// </summary>
        public string OrderNum { get; set; }

        /// <summary>
        /// 生產進度
        /// </summary>
        public int ProgressValue { get; set; }

        /// <summary>
        /// 訂單預交日
        /// </summary>
        public string DueDate { get; set; }

        /// <summary>
        /// 生管預交日
        /// </summary>
        public string DueDatePM { get; set; }

        /// <summary>
        /// 客戶地區/名稱
        /// </summary>
        public string CustomerInfo { get; set; }


        /// <summary>
        /// 訂單狀態【最緊急:"SS"，緊急:"S"，急:"N"】
        /// 這邊好像少了一個不急的狀態，我先加上去【0:一般、1:急、2:緊急、3:最緊急】
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 製程進度【準時:"ONTIME"，延遲:"DELAY"】
        /// </summary>
        public string ProcessProgress { get; set; }

        private string ordernote;
        /// <summary>
        /// 訂單備註
        /// </summary>
        public string OrderNote {
            get { return ordernote; }
            set {
                if (value == null || value == "")
                    ordernote = "-";
                else
                    ordernote = value;
            } }

        /// <summary>
        /// 尚餘加工天數
        /// </summary>
        public string NeedDay { get; set; }

        /// <summary>
        /// 排程估計交期
        /// </summary>
        public string EstimatedDate { get; set; }

        /// <summary>
        /// 未完成品項 2024.12.16新增
        /// </summary>
        public string UndonePart { get; set; }
    }
}
