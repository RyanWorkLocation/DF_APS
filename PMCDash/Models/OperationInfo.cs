using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class OperationInfo
    {
        public OperationInfo(double utilizationRate, string status, double productionProgress, string customName, WorkOrderDetailInformation workorderInfo)
        {
            UtilizationRate = utilizationRate;
            Status = status;
            ProductionProgress = productionProgress;
            CustomName = customName;
            WorkOrderInfo = workorderInfo;
        }

        /// <summary>
        /// 稼動率
        /// </summary>
        public double UtilizationRate { get; set; }
        
        /// <summary>
        /// 當前機台運轉狀態
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 當前生產進度
        /// </summary>
        public double ProductionProgress { get; set; }

        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string CustomName { get; set; }

        public WorkOrderDetailInformation WorkOrderInfo { get;set;}
    }
}
