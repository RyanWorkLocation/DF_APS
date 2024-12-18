using System.Collections.Generic;
using System.Linq;

namespace PMCDash.Models
{

    public class CPKViewModel
    {
        public int Total { get; set; }
        
        public List<Quality> Qualities { get; set; }

    }

    public class Quality
    {
        public Quality(string name, string unit, decimal cPKValue, decimal cp, decimal ca, string total, int distance, decimal cL, decimal[] lSL, decimal[] uSL, decimal[] xChart, decimal[] xRMChart, decimal avg, decimal mse, List<WorkOrderInfo> workOrders, List<int> woInterval)
        {
            Name = name;
            Unit = unit;
            CPKValue = cPKValue;
            Cp = cp;
            Ca = ca;
            Total = total;
            Distance = distance;
            CL = cL;
            LSL = lSL;
            USL = uSL;
            XChart = xChart;
            XRMChart = xRMChart;
            Max = xChart.Max();
            Min = xChart.Min(); 
            if(CPKValue >= 1.33M)
            {
                Evaluation = "表現良好";
            }
            else if(CPKValue >= 1.0M)
            {
                Evaluation = "表現尚可";
            }
            else
            {
                Evaluation = "立即改善";
            }
            AVG = avg;
            MSE = mse;
            WorkOrders = workOrders;
            WOInterval = woInterval.ToArray();
        }

        /// <summary>
        /// CPK評價
        /// </summary>
        public string Evaluation { get; set; }

        /// <summary>
        /// 檢驗項目名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 管制圖單位 EX: nm、mm、cm、M
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// CPK值
        /// </summary>
        public decimal CPKValue { get; set; }

        /// <summary>
        /// Cp值
        /// </summary>
        public decimal Cp { get; set; }

        /// <summary>
        /// Ca值
        /// </summary>
        public decimal Ca { get; set; }
        
        /// <summary>
        /// 總件數
        /// </summary>
        public string Total { get; set; }

        /// <summary>
        /// 移動距離(n)
        /// </summary>
        public int Distance { get; set; }

        /// <summary>
        /// 中心值
        /// </summary>
        public decimal CL { get; set; }

        /// <summary>
        /// 規格/管制下限
        /// </summary>
        public decimal[] LSL { get; set; }

        /// <summary>
        /// 規格/管制上限
        /// </summary>
        public decimal[] USL { get; set; }

        /// <summary>
        /// X管制圖資料
        /// </summary>
        public decimal[] XChart { get; set; }

        /// <summary>
        /// XRM管制圖資料
        /// </summary>
        public decimal[] XRMChart { get; set; }

        /// <summary>
        /// 最大值
        /// </summary>
        public decimal Max { get; set; }

        /// <summary>
        /// 最小值
        /// </summary>
        public decimal Min { get; set; }

        /// <summary>
        /// 平均
        /// </summary>
        public decimal AVG { get; set; }

        /// <summary>
        /// 標準差
        /// </summary>
        public decimal MSE { get; set; }

        /// <summary>
        /// 工單區間
        /// </summary>
        public int[] WOInterval { get; set; }

        /// <summary>
        /// 每點(X管制圖)工單明細
        /// </summary>
        public List<WorkOrderInfo> WorkOrders { get; set; }
    }
    
    public class WorkOrderInfo
    {
        public WorkOrderInfo(string wONo, string custom, int require, string dueDate)
        {
            WONo = wONo;
            Custom = custom;
            Require = require;
            DueDate = dueDate;
        }

        /// <summary>
        /// 工單編號
        /// </summary>
        public string WONo { get; set; }

        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string Custom { get; set; }

        /// <summary>
        /// 工單數量
        /// </summary>
        public int Require { get; set; }

        /// <summary>
        /// 預交日期
        /// </summary>
        public string DueDate { get; set; }
    }
}
