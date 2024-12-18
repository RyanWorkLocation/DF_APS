using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class ActionRequest<T>
    {
        public string Action { get; set; }

        public T Data { get; set; }
    }

    public class RequestReportWorkOrder
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
        /// 報工模式
        /// </summary>
        public int ReportMode { get; set; }
    }

    public class RequestQCMeasureDeviceStatus
    {
        /// <summary>
        /// 檢驗設備ID
        /// </summary>
        public string DeviceID { get; set; }
        /// <summary>
        /// 最新量測資料時間
        /// </summary>
        public DateTime Top1MeasureDataTime { get; set; }
        /// <summary>
        /// 檢驗設備連線狀態
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// 量測項目
        /// </summary>
        public string MeasureItem { get; set; }
        /// <summary>
        /// 量測數值
        /// </summary>
        public string MeasureValue { get; set; }
        /// <summary>
        /// 量測時間
        /// </summary>
        public string MeasureTime { get; set; }
        /// <summary>
        /// 量測單位
        /// </summary>
        public string Unit { get; set; }
    }

    public class RequestQCMeasureData
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
        /// 檢驗人員
        /// </summary>
        public string QC_OperatorName { get; set; }
        /// <summary>
        /// 檢驗部位數值
        /// </summary>
        public List<QCMeasureData> QCPlaceItem { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
    }

    public class RequestWorkOrder
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
        /// 報工模式
        /// </summary>
        public int ReportMode { get; set; }
    }

    public class RequestFactory
    {
        /// <summary>
        /// 廠區名稱
        /// </summary>
        public string FactoryName { get; set; }

        /// <summary>
        /// 產線名稱
        /// </summary>
        public string ProductionName { get; set; }


        /// <summary>
        /// 機台名稱
        /// </summary>
        public string DeviceName { get; set; }

    }

    public class RequestProductionLine
    {
        /// <summary>
        /// 廠區名稱
        /// </summary>
        public string FactoryName { get; set; }

        /// <summary>
        /// 產線名稱
        /// </summary>
        public string ProductionName { get; set; }
    }

    public class RequestQualityControlWorkList
    {
        /// <summary>
        /// 廠區名稱
        /// </summary>
        public string FactoryName { get; set; }

        /// <summary>
        /// 產線名稱
        /// </summary>
        public string ProductionName { get; set; }


        /// <summary>
        /// 機台名稱
        /// </summary>
        public string DeviceName { get; set; }
    }

}
