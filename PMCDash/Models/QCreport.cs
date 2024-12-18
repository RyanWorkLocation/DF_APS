using PMCDash.Models.Part2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class QCreport
    {
        //public QCworkListInfomation(OrderInformation orderInfo, string wIPStatus, string stratTime, string endTime, int delayDays)
        //{
        //    OrderInfo = orderInfo;
        //    WIPStatus = wIPStatus;
        //    StratTime = stratTime;
        //    EndTime = endTime;
        //    DelayDays = delayDays;
        //}

        public QCreport()
        {
            OrderNo = "D212119";
            BOM_No = "DBOM076";
            BOM_Name = "料管組";
            Componentname = "料管組";
            MaterialNo = "TOPA-SRC";
            MaterialName = "雙合金A級";
            CurrentCount = 1; //當前數量
            Group = "安和"; //生產群組
            Machine = "CNC-3"; //生產機器
            OperatorName = "粗車外徑"; //生產負責人
            QC_OperatorName = "蘇文政";//檢驗負責人名稱
            QC_Mode = "全檢";//全/抽驗、已檢驗數量/需鑑驗數量
        }
        /// <summary>
        /// 工單基本資料
        /// </summary>
        public WorkOrderDetailInformation OrderInfo { get; set; }

        /// <summary>
        /// 訂單編號
        /// </summary>
        public string OrderNo { get; set; }

        /// <summary>
        /// BOM編號
        /// </summary>
        public string BOM_No { get; set; }

        /// <summary>
        /// BOM名稱(產品名稱)
        /// </summary>
        public string BOM_Name { get; set; }

        /// <summary>
        /// 物料名稱
        /// </summary>
        public string Componentname { get; set; }

        /// <summary>
        /// 素材編號
        /// </summary>
        public string MaterialNo { get; set; }

        /// <summary>
        /// 素材名稱
        /// </summary>
        public string MaterialName { get; set; }

        /// <summary>
        /// 當前數量
        /// </summary>
        public int CurrentCount { get; set; }

        /// <summary>
        /// 群組名稱
        /// </summary>
        public string Group { get; set; }

        /// <summary>
        /// 機器設備名稱
        /// </summary>
        public string Machine { get; set; }

        /// <summary>
        /// 生產負責人
        /// </summary>
        public string OperatorName { get; set; }

        /// <summary>
        /// 檢驗負責人名稱
        /// </summary>
        public string QC_OperatorName { get; set; }

        /// <summary>
        /// 檢驗模式
        /// </summary>
        public string QC_Mode { get; set; }




















        ///// <summary>
        ///// 工單編號
        ///// </summary>
        //public string WorkOrderNo { get; set; }

        ///// <summary>
        ///// 製程編號(工序編號)
        ///// </summary>
        //public int ProcessNo { get; set; }

        ///// <summary>
        ///// 製程名稱(工序名稱)
        ///// </summary>
        //public string ProcessName { get; set; }

        ///// <summary>
        ///// 物料編號(產品名稱)
        ///// </summary>
        //public string ComponentNo { get; set; }

        ///// <summary>
        ///// 工單數量
        ///// </summary>
        //public int RequireCount { get; set; }

        ///// <summary>
        ///// 檢驗期限(預交日期)
        ///// </summary>
        //public string QC_DueDate { get; set; }

    }

    public class QCWorkTask
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderId { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPId { get; set; }

        /// <summary>
        /// 製程名稱
        /// </summary>
        public string OPName { get; set; }

        /// <summary>
        /// 產品編號
        /// </summary>
        public string ProductId { get; set; }

        /// <summary>
        /// 產品名稱
        /// </summary>
        public string ProductName { get; set; }

        /// <summary>
        /// 預計開工
        /// </summary>
        public string StartTime { get; set; }

        /// <summary>
        /// 預計完工
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// 開工狀態【0:未開工、1:生產中、2:暫停中、3:已完成】
        /// </summary>
        public string OPStatus { get; set; }

        /// <summary>
        /// 產品交期
        /// </summary>
        public string AssignDate { get; set; }

        /// <summary>
        /// 延遲交期
        /// </summary>
        public string DeleyDays { get; set; }

        /// <summary>
        /// 機台編號
        /// </summary>
        public string Deviec { get; set; }

        /// <summary>
        /// 機台群組
        /// </summary>
        public string DeviceGroup { get; set; }

        /// <summary>
        /// 人員姓名
        /// </summary>
        public string OperatorName { get; set; }

        /// <summary>
        /// 進度
        /// </summary>
        public double Progress { get; set; }

        /// <summary>
        /// 需求數量
        /// </summary>
        public string RequireNum { get; set; }

        /// <summary>
        /// 已完成數量
        /// </summary>
        public string CompleteNum { get; set; }

        /// <summary>
        /// 不良品數量
        /// </summary>
        public string DefectiveNum { get; set; }

        /// <summary>
        /// 品質檢驗【N/A:不需檢驗、False:需檢驗、未檢驗完成、True:已經檢驗】
        /// </summary>
        public string IsQC { get; set; }

        /// <summary>
        /// 品檢人員
        /// </summary>
        public string QCman { get; set; }

    }

    public class QCWorkTaskDetail
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderId { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPId { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPName { get; set; }

        /// <summary>
        /// 產品編號
        /// </summary>
        public string ProductId { get; set; }

        /// <summary>
        /// 預計開工
        /// </summary>
        public string StartTime { get; set; }

        /// <summary>
        /// 預計完工
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// 產品交期
        /// </summary>
        public string AssignDate { get; set; }

        /// <summary>
        /// 延期天數
        /// </summary>
        public string DeleyDays { get; set; }

        /// <summary>
        /// 設備編號
        /// </summary>
        public string Deviec { get; set; }

        /// <summary>
        /// 機台群組
        /// </summary>
        public string DeviceGroup { get; set; }

        /// <summary>
        /// 人員姓名
        /// </summary>
        public string OperatorName { get; set; }

        /// <summary>
        /// 進度
        /// </summary>
        public string Progress { get; set; }

        /// <summary>
        /// 需求數量
        /// </summary>
        public string RequireNum { get; set; }

        /// <summary>
        /// 完成數量
        /// </summary>
        public string CompleteNum { get; set; }

        /// <summary>
        /// 剩餘數量
        /// </summary>
        public string RemainNum { get; set; }

        /// <summary>
        /// 品檢狀況
        /// </summary>
        public string IsQC { get; set; }

        /// <summary>
        /// 品檢人員
        /// </summary>
        public string QCman { get; set; }

        /// <summary>
        /// 開工類別
        /// </summary>
        public string type { get; set; }

        /// <summary>
        /// 良品數量
        /// </summary>
        public string QtyGood { get; set; }

        /// <summary>
        /// 不良品數量
        /// </summary>
        public string QtyBad { get; set; }

        /// <summary>
        /// 圖片位址
        /// </summary>
        public string ImgPath { get; set; }

        public List<QC> QCList { get; set; }

    }


    public class QCMeasureDataReport
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string WorkOrderNo { get; set; }
        /// <summary>
        /// 物料編號
        /// </summary>
        public string PartNo { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public int OPNo { get; set; }
        /// <summary>
        /// 檢驗人員
        /// </summary>
        public string QC_OperatorName { get; set; }
        /// <summary>
        /// 需求檢驗數量
        /// </summary>
        public int QC_Quantity { get; set; }
        /// <summary>
        /// 已驗數數量
        /// </summary>
        public int QC_CurrentCount { get; set; }
        /// <summary>
        /// 檢驗部位數值
        /// </summary>
        public List<QCMeasureData> QCPlaceItem { get; set; }
        /// <summary>
        /// 良品數量
        /// </summary>
        public int PASSofPartCount { get; set; }
        /// <summary>
        /// 不良品數量
        /// </summary>
        public int NotPASSofPartCount { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
    }

    public class QCMeasureData
    {
        /// <summary>
        /// 檢驗部位編號
        /// </summary>
        public string QCPlaceNo { get; set; }
        /// <summary>
        /// 檢驗部位名稱
        /// </summary>
        public string QCPlaceNanme { get; set; }
        /// <summary>
        /// 檢驗工具編號
        /// </summary>
        public string QCToolNo { get; set; }
        /// <summary>
        /// 檢驗部位單位
        /// </summary>
        public string QCPlaceUnit { get; set; }
        /// <summary>
        /// 檢驗實際數值
        /// </summary>
        public string QCRealValue { get; set; }
        /// <summary>
        /// 檢驗實際時間
        /// </summary>
        public DateTime QCRealDateTime { get; set; }
    }

    public class QCPlace
    {
        ///// <summary>
        ///// 檢驗部位資料
        ///// </summary>
        ///// <param name="_QCPlaceNo">檢驗部位編號</param>
        ///// <param name="_QCPlaceNanme">檢驗部位名稱</param>
        ///// <param name="_QCPlaceUnit">鑑驗單位</param>
        ///// <param name="_QCPlaceUCL">管制上限</param>
        ///// <param name="_QCPlaceLCL">管制下限</param>
        ///// <param name="_QCPlaceCL">管制中心值</param>
        ///// <param name="_QCPlaceUSL">規格上限值</param>
        ///// <param name="_QCPlaceLSL">規格下限值</param>
        ///// <param name="_Remark">備註</param>
        //public QCPlace(string _QCPlaceNo, string _QCPlaceNanme, string _QCPlaceUnit, string _QCPlaceUCL, string _QCPlaceLCL, string _QCPlaceCL, string _QCPlaceUSL, string _QCPlaceLSL, string _Remark)
        //{
        //    QCPlaceNo = _QCPlaceNo;
        //    QCPlaceNanme = _QCPlaceNanme;
        //    QCPlaceUnit = _QCPlaceUnit;
        //    QCPlaceUCL = _QCPlaceUCL;
        //    QCPlaceLCL = _QCPlaceLCL;
        //    QCPlaceCL = _QCPlaceCL;
        //    QCPlaceUSL = _QCPlaceUSL;
        //    QCPlaceLSL = _QCPlaceLSL;
        //    Remark = _Remark;
        //}
        /// <summary>
        /// 檢驗部位編號
        /// </summary>
        public string QCPlaceNo { get; set; }
        /// <summary>
        /// 檢驗部位名稱
        /// </summary>
        public string QCPlaceNanme { get; set; }
        /// <summary>
        /// 檢驗部位單位
        /// </summary>
        public string QCPlaceUnit { get; set; }
        /// <summary>
        /// 管制上限
        /// </summary>
        public string QCPlaceUCL { get; set; }
        /// <summary>
        /// 管制下限
        /// </summary>
        public string QCPlaceLCL { get; set; }
        /// <summary>
        /// 管制中心值
        /// </summary>
        public string QCPlaceCL { get; set; }
        /// <summary>
        /// 規格上限
        /// </summary>
        public string QCPlaceUSL { get; set; }
        /// <summary>
        /// 規格下限
        /// </summary>
        public string QCPlaceLSL { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }


    }
}
