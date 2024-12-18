using PMCDash.Models.Part2;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    /// <summary>
    /// 取回資料庫供單用
    /// </summary>
    public class Assignment
    {
        public string SeriesID { get; set; }
        public string OrderID { get; set; }
        public string ERPOrderID { get; set; }
        public string OPID { get; set; }
        public string OPLTXA1 { get; set; }
        public string MachOpTime { get; set; }
        public string HumanOpTime { get; set; }
        public string StartTime { get; set; }
        public string EndTime { get; set; }
        public string WorkGroup { get; set; }
        public string Operator { get; set; }
        public string AssignDate { get; set; }
        public string OrderQTY { get; set; }
        public string Scheduled { get; set; }
        public string AssignDate_PM { get; set; }
        public string MAKTX { get; set; }
        public string PRIORITY { get; set; }
        public string ImgPath { get; set; }
        public string Note { get; set; }
        public string Important { get; set; }
        public string CPK { get; set; }
    }

    public class OrderGroup
    {
        public string OrderID { get; set; }

        public double OrderProgress { get; set; }

        public List<WorkOrderOverview> WorkOrder { get; set; }
    }

    public class WorkOrderOverview
    {
        public string WorkOrderID { get; set; }

        public double WorkOrderProgress { get; set; }

        public List<OPOverview> OP { get; set; }
    }

    /// <summary>
    /// BOM物料結構
    /// </summary>
    public class BOMIteem
    {
        //public BOMIteem(string _BOMNo, string _PartNo, string _PartName, int _QuantityOfBOM, string _Material, string _ProcessMethod, string _Price, string _SubTotal,string _Remark)
        //{
        //    BOMNo = _BOMNo;
        //    PartNo = _PartNo;
        //    PartName = _PartName;
        //    QuantityOfBOM = _QuantityOfBOM;
        //    Material = _Material;
        //    ProcessMethod = _ProcessMethod;
        //    Price = _Price;
        //    SubTotal = _SubTotal;
        //    Remark = _Remark;
        //}
        /// <summary>
        /// BOM編號
        /// </summary>
        public string BOMNo { get; set; }
        /// <summary>
        /// 物料編號
        /// </summary>
        public string PartNo { get; set; }
        /// <summary>
        /// 物料名稱
        /// </summary>
        public string PartName { get; set; }
        /// <summary>
        /// BOM需求數量
        /// </summary>
        public int QuantityOfBOM { get; set; }
        /// <summary>
        /// 材質
        /// </summary>
        public string Material { get; set; }
        /// <summary>
        /// 處理方式
        /// </summary>
        public string ProcessMethod { get; set; }
        /// <summary>
        /// 單價
        /// </summary>
        public string Price { get; set; }
        /// <summary>
        /// 金額
        /// </summary>
        public string SubTotal { get; set; }
        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
    }

    /// <summary>
    /// 客戶資料
    /// </summary>
    public class CustomerInformation
    {
        public CustomerInformation(string _CustomerNo, string _CustomerName, string _CustomerBuyer, string _TEL, string _FAX, string _Payment, string _BusinessTax, string _Currency, string _InternalOrderNo, string _OrderDate)
        {
            CustomerNo = _CustomerNo;
            CustomerName = _CustomerName;
            CustomerBuyer = _CustomerBuyer;
            TEL = _TEL;
            FAX = _FAX;
            Payment = _Payment;
            BusinessTax = _BusinessTax;
            Currency = _Currency;
            InternalOrderNo = _InternalOrderNo;
            OrderDate = _OrderDate;
        }

        /// <summary>
        /// 客戶編號
        /// </summary>
        public string CustomerNo { get; set; }
        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string CustomerName { get; set; }
        /// <summary>
        /// 採購人員
        /// </summary>
        public string CustomerBuyer { get; set; }
        /// <summary>
        /// 電話
        /// </summary>
        public string TEL { get; set; }
        /// <summary>
        /// 傳真
        /// </summary>
        public string FAX { get; set; }
        /// <summary>
        /// 付款說明
        /// </summary>
        public string Payment { get; set; }
        /// <summary>
        /// 營業稅
        /// </summary>
        public string BusinessTax { get; set; }
        /// <summary>
        /// 幣別
        /// </summary>
        public string Currency { get; set; }
        /// <summary>
        /// 內部單號
        /// </summary>
        public string InternalOrderNo { get; set; }
        /// <summary>
        /// 接單日期
        /// </summary>
        public string OrderDate { get; set; }
    }

    /// <summary>
    /// 工單資訊
    /// </summary>
    public class WorkOrderInformation
    {
        //public WorkOrderInformation(string orderNo, int oPNo, string opName, string productNo, int requireCount, int currentCount, string _StartTime, string _EndTime, string _Machine,string _Group,string _Operator,string _WorkOrderStatus)
        //{
        //    OrderNo = orderNo;
        //    OPNo = oPNo;
        //    OPName = opName;
        //    RequireCount = requireCount;
        //    CurrentCount = currentCount;
        //    StarTime = _StartTime;
        //    EndTime = _EndTime;
        //    Machine = _Machine;
        //    Group = _Group;
        //    Operator = _Operator;
        //    WorkOrderStatus = _WorkOrderStatus;
        //}

        /// <summary>
        /// 工單編號
        /// </summary>
        /// [Display(Name = "機台編號")]
        public string OrderNo { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public int OPNo { get; set; }

        /// <summary>
        /// 製程名稱
        /// </summary>
        public string OPName { get; set; }

        /// <summary>
        /// 分發機台
        /// </summary>
        public string Machine { get; set; }

        /// <summary>
        /// 工單數量
        /// </summary>
        public int RequireCount { get; set; }

        /// <summary>
        /// 當前數量
        /// </summary>
        public int CurrentCount { get; set; }

        /// <summary>
        /// 開始時間
        /// </summary>
        public string StarTime { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// 負責群組
        /// </summary>
        public string Group { get; set; }

        /// <summary>
        /// 負責人
        /// </summary>
        public string OperatorName { get; set; }

        /// <summary>
        /// 工單狀態
        /// </summary>
        public string WorkOrderStatus { get; set; }
    }

    /// <summary>
    /// 工單詳細資訊
    /// </summary>
    public class WorkOrderDetailInformation
    {
        public WorkOrderDetailInformation(string orderNo, int oPNo, string opName, string productNo, int requireCount, int currentCount, DateTime dueDate)
        {
            OrderNo = orderNo;
            OPNo = oPNo;
            OPName = opName;
            ProductNo = productNo;
            RequireCount = requireCount;
            CurrentCount = currentCount;
            DueDate = dueDate;
        }

        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderNo { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public int OPNo { get; set; }

        /// <summary>
        /// 製程名稱
        /// </summary>
        public string OPName { get; set; }

        /// <summary>
        /// 產品名稱
        /// </summary>
        public string ProductNo { get; set; }

        /// <summary>
        /// 分發機台
        /// </summary>
        public string Machine { get; set; }

        /// <summary>
        /// 工單數量
        /// </summary>
        public int RequireCount { get; set; }

        /// <summary>
        /// 當前數量
        /// </summary>
        public int CurrentCount { get; set; }

        /// <summary>
        /// 開始時間
        /// </summary>
        public DateTime StarTime { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        public DateTime EndTime { get; set; }

        /// <summary>
        /// 預交日期
        /// </summary>
        public DateTime DueDate { get; set; }

        /// <summary>
        /// 延遲時數
        /// </summary>
        public DateTime DeleyTime { get; set; }

        /// <summary>
        /// 生管預交日
        /// </summary>
        public DateTime PlanningDueDate { get; set; }

        /// <summary>
        /// 生產進度
        /// </summary>
        public double ProgressBarRate { get; set; }

        /// <summary>
        /// 負責群組
        /// </summary>
        public string Group { get; set; }

        /// <summary>
        /// 負責人(操作人員)
        /// </summary>
        public string Operator { get; set; }

        /// <summary>
        /// 工單狀態
        /// </summary>
        public string WorkOrderStatus { get; set; }

        /// <summary>
        /// 設計圖製路徑
        /// </summary>
        public string CADFilePath { get; set; }

        /// <summary>
        /// 訂單交付日期
        /// </summary>
        public double OrderDueDate { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }
    }

    /// <summary>
    /// 綁定工單資訊
    /// </summary>
    public class WorkOrderBindInfo
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string DeviceId { get; set; }

        /// <summary>
        /// 機台名稱
        /// </summary>
        public string MachineName { get; set; }

        /// <summary>
        /// 工單編號
        /// </summary>
        public string WorkOrderID { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }

        /// <summary>
        /// 作業人員編號
        /// </summary>
        public string OperatorID { get; set; }

        /// <summary>
        /// 作業人員名稱
        /// </summary>
        public string OperatorName { get; set; }

        /// <summary>
        /// 建立時間
        /// </summary>
        public string CreateTime { get; set; }

        /// <summary>
        /// 更新時間
        /// </summary>
        public string LastUpdateTime { get; set; }

    }

    public class OrderInfo : ICloneable
    {
        [Display(Name = "工單SeriesID")]
        [Required]
        public string SeriesID { get; set; }
        [Display(Name = "工單編號")]
        [Required]
        public string OrderID { get; set; }

        [Display(Name = "工序編號")]
        public double OPID { get; set; }

        [Display(Name = "製程順序")]
        public int Range { get; set; }

        [Display(Name = "分發數量")]
        public int OrderQTY { get; set; }

        [Display(Name = "開始時間")]
        public string StartTime { get; set; }

        [Display(Name = "結束時間")]
        public string EndTime { get; set; }

        [Display(Name = "預交日期")]
        public string AssignDate { get; set; }

        [Display(Name = "延遲天數")]
        public int DelayDays { get; set; }

        [Display(Name = "分發機台")]
        public string WorkGroup { get; set; }

        [Display(Name = "人工時")]
        public double HumanOpTime { get; set; }

        [Display(Name = "機工時")]
        public double MachOpTime { get; set; }

        [Display(Name = "工單內文")]
        public string OPLTXA1 { get; set; }

        [Display(Name = "生管預交日期")]
        public string AssignDate_PM { get; set; }

        [Display(Name = "工件號碼")]
        public string Maktx { get; set; }

        public int Progress { get; set; }

        public int PRIORITY { get; set; }

        public string Note { get; set; }

        public bool Important { get; set; }

        public bool Assign { get; set; }

        public string ProcessStatus { get; set; }
        public int CanSync { get; set; }

        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string CustomerName { get; set; }

        /// <summary>
        /// 客戶地點
        /// </summary>
        public string CustomerLocation { get; set; }

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }
}
