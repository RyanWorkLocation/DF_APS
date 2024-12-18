using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class StatusDefine
    {
        public StatusDefine(List<ProcessStatus> processDefine, List<OrderStatus> orderDefine, List<OrderListType> orderListType)
        {
            ProcessDefine = processDefine;
            OrderDefine = orderDefine;
            OrderListType = orderListType;
        }

        /// <summary>
        /// 製程進度狀態清單
        /// </summary>
        public List<ProcessStatus> ProcessDefine { get; set; }

        /// <summary>
        /// 訂單狀態清單
        /// </summary>
        public List<OrderStatus> OrderDefine { get; set; }

        /// <summary>
        /// 訂單狀態清單
        /// </summary>
        public List<OrderListType> OrderListType { get; set; }
    }

    public class ProcessStatus
    {
        public ProcessStatus(string status, string displayName)
        {
            Status = status;
            DisplayName = displayName;
        }

        /// <summary>
        /// 製程進度狀態 分為兩項 EX ONTIME、DELAY
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 當前狀態描述 EX 準時、延遲
        /// </summary>
        public string DisplayName { get; set; }
    }

    public class OrderStatus
    {
        public OrderStatus(string status, string displayName)
        {
            Status = status;
            DisplayName = displayName;
        }

        /// <summary>
        /// 訂單狀態 EX: 目前是 SS、S、N
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 訂單狀態描述 EX: 目前是 最緊急、緊急、急
        /// </summary>
        public string DisplayName { get; set; }
    }

    public class OrderListType
    {
        public OrderListType(string code, string displayName)
        {
            Code = code;
            DisplayName = displayName;
        }

        /// <summary>
        /// 訂單狀態 EX: 目前是 SS、S、N
        /// </summary>
        public string Code { get; set; }

        /// <summary>
        /// 訂單狀態描述 EX: 目前是 最緊急、緊急、急
        /// </summary>
        public string DisplayName { get; set; }
    }
}
