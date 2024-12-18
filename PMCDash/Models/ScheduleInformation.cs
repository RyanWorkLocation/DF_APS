using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class ScheduleInformation
    {
        public ScheduleInformation(WorkOrderDetailInformation orderInfo, string wIPStatus, string stratTime, string endTime, int delayDays)
        {
            OrderInfo = orderInfo;
            WIPStatus = wIPStatus;
            StratTime = stratTime;
            EndTime = endTime;
            DelayDays = delayDays;
        }

        public WorkOrderDetailInformation OrderInfo { get; set; }

        public string WIPStatus { get; set; }

        public string StratTime { get; set; }

        public string EndTime { get; set; }

        public int DelayDays { get; set; }
    }

    public class Schedulelist
    {
        /// <summary>
        /// 排程模式編號【1:交期優先-推薦權重、2:交期優先-自訂權重】
        /// </summary>
        public string mode { get; set; }
        /// <summary>
        /// 排程產出資料
        /// </summary>
        public List<Schedule> schedules { get; set; }
    }

    public class Schedule
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPName { get; set; }

        /// <summary>
        /// 是否為重要工單
        /// </summary>
        public bool Important { get; set; }

        ///////////////可編排順序////////////////////

        /// <summary>
        /// 物料編號
        /// </summary>
        public string MAKTX { get; set; }

        /// <summary>
        /// 物料名稱
        /// </summary>
        public string PartName { get; set; }

        /// <summary>
        /// 分發機台
        /// </summary>
        public string WorkGroup { get; set; }

        /// <summary>
        /// 需求數量
        /// </summary>
        public int OrderQTY { get; set; }

        /// <summary>
        /// 開始時間
        /// </summary>
        public string StartTime { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// 預交日期
        /// </summary>
        public string AssignDate { get; set; }

        /// <summary>
        /// 延遲天數
        /// </summary>
        public int DelayDays { get; set; }

        /// <summary>
        /// 生管預交日期
        /// </summary>
        public string AssignDate_PM { get; set; }

        /// <summary>
        /// 生產進度
        /// </summary>
        public int Progress { get; set; }

        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string CustomerName { get; set; }

        /// <summary>
        /// 客戶地點
        /// </summary>
        public string CustomerLocation { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string Note { get; set; }

        /// <summary>
        /// 人工時
        /// </summary>
        public string HunmanOpTime { get; set; }
        /// <summary>
        /// 機工時
        /// </summary>
        public string MachOpTime { get; set; }

        ///////////////////////////////////////

        ///// <summary>
        ///// CPK數值(CPK>1.67為合格, 1.33<=CPK<=1.67為不合格, CPK<1.33為立即檢討)
        ///// </summary>
        //public double CPK { get; set; }

        ///// <summary>
        ///// CPK評價(根據CPK數值給定判斷標準)
        ///// </summary>
        //public string CPKLevel { get; set; }

        /// <summary>
        /// 工單狀態
        /// </summary>
        public int WIPEvent { get; set; }



        /// <summary>
        /// 是否為已派工單
        /// </summary>
        public bool Assign { get; set; }

        /// <summary>
        /// 製程狀態
        /// DELAY:已經延遲，WILLDELAY:即將延遲，NORMAL:進度正常
        /// </summary>
        public string ProcessStatus { get; set; }

        ///// <summary>
        ///// 是否即將延遲
        ///// </summary>
        //public bool AssignDelayIsComing { get; set; }

        ///// <summary>
        ///// 是否已經延遲
        ///// </summary>
        //public bool OrderHasDelayed { get; set; }

        ///// <summary>
        ///// 剩餘需生產數量
        ///// </summary>
        //public string QtyRemain { get; set; }0.


        ///// <summary>
        ///// 訂單交付
        ///// </summary>
        //public string OrderDeadline { get; set; }

        ///// <summary>
        ///// wt數值
        ///// </summary>
        //public string WT { get; set; }

        ///// <summary>
        ///// ct數值
        ///// </summary>
        //public string CT { get; set; }

        ///// <summary>
        ///// 製程名稱
        ///// </summary>
        //public string OPLTXA1 { get; set; }
    }

    /// <summary>
    /// 機台故障專用
    /// </summary>
    public class DelayScheduleOP
    {
        /// <summary>
        /// 報修機台名稱
        /// </summary>
        public string MachineName { get; set; }
        /// <summary>
        /// 開始時間
        /// </summary>
        public string BreakdownST { get; set; }
        /// <summary>
        /// 結束時間
        /// </summary>
        public string BreakdownET { get; set; }
    }

    public class Schedulelist_MB
    {
        /// <summary>
        /// 排程模式編號【1:交期優先-推薦權重、2:交期優先-自訂權重】
        /// </summary>
        public string mode { get; set; }
        /// <summary>
        /// 設備故障資訊
        /// </summary>
        public MachinebreakdowInfo breakdownInfo { get; set; }
        /// <summary>
        /// 排程產出資料
        /// </summary>
        public List<Schedule> schedules { get; set; }
    }

    public class MachinebreakdowInfo
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Machine { get; set; }
        /// <summary>
        /// 機台維修開始時間
        /// </summary>
        public string StartTime { get; set; }
        /// <summary>
        /// 機台維修結束時間
        /// </summary>
        public string EndTime { get; set; }
    }

    public class Dashboard
    {
        ///// <summary>
        ///// 欄位ID
        ///// </summary>
        //public int ID { get; set; }
        /// <summary>
        /// 欄位名稱(中文)
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 欄位順序
        /// </summary>
        public int Sequence { get; set; }
        /// <summary>
        /// 欄位名稱(英文)，可對照DashBoard API中輸出欄位名稱
        /// </summary>
        public string Key { get; set; }
        ///// <summary>
        ///// 是否可以前端排序
        ///// </summary>
        //public bool CanBeSort { get; set; }
        /// <summary>
        /// 是否固定為左側欄位
        /// </summary>
        public bool Freeze { get; set; }
        ///// <summary>
        ///// 資料表欄寬
        ///// </summary>
        //public string TableWidth { get; set; }
        /// <summary>
        /// 資料屬性
        /// </summary>
        public string DateType { get; set; }
        /// <summary>
        /// 是否顯示
        /// </summary>
        public bool IsShow { get; set; }
        /// <summary>
        /// 是否進階搜尋
        /// </summary>
        public bool IsAdvancedSearch { get; set; }

    }

    public class DashboardUpdate
    {
        /// <summary>
        /// 欄位名稱(中文)
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 欄位順序
        /// </summary>
        public int Sequence { get; set; }
        /// <summary>
        /// 表單種類2:定單列表、3:工單列表
        /// </summary>
        public int Page { get; set; }
    }

    public class SearchParameter
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        //public List<string> OrderID { get; set; }
        public string OrderID { get; set; }

        /// <summary>
        /// 物料編號
        /// </summary>
        public string MAKTX { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        //public List<string> OPID { get; set; }
        public string OPID { get; set; }

        /// <summary>
        /// 分發機台
        /// </summary>
        public string WorkGroup { get; set; }

        /// <summary>
        /// 交付日期
        /// </summary>
        public string AssignDate { get; set; }

        /// <summary>
        /// 生管預交日期
        /// </summary>
        public string AssignDate_PM { get; set; }

        /// <summary>
        /// 分發數量
        /// </summary>
        public string OrderQty { get; set; }

        /// <summary>
        /// 客戶名稱
        /// </summary>
        public string CustomerName { get; set; }

        /// <summary>
        /// 客戶地區
        /// </summary>
        public string CustomerLocation { get; set; }

        /// <summary>
        /// 產品名稱
        /// </summary>
        public string ProductName { get; set; }

        /// <summary>
        /// 篩選是否為重要工單
        /// </summary>
        public string Important { get; set; }

        /// <summary>
        /// 篩選是否為未派工單
        /// </summary>
        public string Scheduled { get; set; }

        /// <summary>
        /// 篩選是否為即將延遲
        /// </summary>
        public string WillDelay { get; set; }

        /// <summary>
        /// 篩選是否為已經延遲
        /// </summary>
        public string AlreadyDelay { get; set; }

        ///// <summary>
        ///// 篩選CPK評價(1: 合格, 2: 不合格, 3:立即檢討)
        ///// </summary>
        //public string CPK { get; set; }

        /// <summary>
        /// 開始時間
        /// </summary>
        public string StartTime { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        public string EndTime { get; set; }

        ///// <summary>
        ///// 工作人員編號
        ///// </summary>
        //public string Operator { get; set; } 

        public SearchParameter()
        {
            //this.OrderID = new List<string>();
            //this.OPID = new List<string>();
            this.OrderID = "";
            this.OPID = "";
            this.Important = @"'%%'";
            this.Scheduled = @"'%%'";
            this.WillDelay = "";
            this.AlreadyDelay = "";
            //this.CPK = "";
            this.StartTime = "";
            this.EndTime = "";
            this.MAKTX = "";
            this.AssignDate_PM = "";
            this.AssignDate = "";
            //this.Operator = "";
            this.CustomerName = "";
            this.CustomerLocation = "";
            this.WorkGroup = "";
            this.OrderQty = "";
            this.ProductName = "";
        }

    }

    public class AssignmentInfo
    {
        public List<AssignmentID> AssignmentID { get; set; }

        public List<int> OPCount { get; set; }

        public List<List<double>> Code { get; set; }
    }

    public class AssignmentID
    {
        public string OrderID { get; set; }

        public string OPID { get; set; }

        public double NeedTime { get; set; }
    }

    public class ScheduleWeight
    {
        /// <summary>
        /// 權重名稱
        /// </summary>
        public string DisplayName { get; set; }
        /// <summary>
        /// 權重關鍵字
        /// </summary>
        public string Key { get; set; }
        /// <summary>
        /// 權重值
        /// </summary>
        public int WeightDefaultValue { get; set; }
        /// <summary>
        /// 權重最小值
        /// </summary>
        public int MiniWeight { get; set; }
        /// <summary>
        /// 權重最大值
        /// </summary>
        public int MaxWeight { get; set; }
    }

    public class Weight
    {
        ///// <summary>
        ///// 權重編號
        ///// </summary>
        //public int WeightCode { get; set; }
        ///// <summary>
        ///// 權重值
        ///// </summary>
        //public int WeightValue { get; set; }

        /// <summary>
        /// 準交率
        /// </summary>
        public string OrderFillRate { get; set; }
        /// <summary>
        /// 稼動率
        /// </summary>
        public string UtilizationRate { get; set; }
        /// <summary>
        /// 搬移距離
        /// </summary>
        public string MovingDistance { get; set; }
        /// <summary>
        /// 負載平衡
        /// </summary>
        public string LoadBalance { get; set; }
    }

    public class ScheduleDataComparison
    {
        /// <summary>
        /// 比較項目
        /// </summary>
        public string Item { get; set; }

        /// <summary>
        /// WT值
        /// </summary>
        public string WTValue { get; set; }

        /// <summary>
        /// CT值
        /// </summary>
        public string CTValue { get; set; }

        /// <summary>
        /// 總延遲時間
        /// </summary>
        public string TotalDelay { get; set; }

        /// <summary>
        /// 搬移次數
        /// </summary>
        public string MovingTime { get; set; }

        /// <summary>
        /// 稼動率
        /// </summary>
        public string UtilizationRate { get; set; }

        /// <summary>
        /// 準交率
        /// </summary>
        public string OrderFillRate { get; set; }
    }

    public class ScheduleMethodSelection
    {
        /// <summary>
        /// 排程方法的編號
        /// </summary>
        public string MethodId { get; set; }
        /// <summary>
        /// 排程方法的名稱
        /// </summary>
        public string MethodName { get; set; }
    }


    public class chartTab
    {
        /// <summary>
        /// 排程代號
        /// </summary>
        public string ModeCode { get; set; }
        /// <summary>
        /// 標籤名稱
        /// </summary>
        public string DisplayTitle { get; set; }
        ///// <summary>
        ///// 是否被選擇
        ///// </summary>
        //public bool IsFocus { get; set; }

    }

    public class GanttChartData1
    {
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string WorkGroup { get; set; }
        public string MachineState { get; set; }
        public double op { get; set; }
        public string orderNum { get; set; }
        public GanttChartStatus ganttChartStatus { get; set; }
        public GanttChartPRInterval ganttChartPRInterval { get; set; }
    }

    public class GanttChartStatus
    {
        public bool AssignDelayIsComing { get; set; }
        public bool OrderHasDelayed { get; set; }
        public bool IsImportant { get; set; }
        public int OrderState { get; set; }
        public bool IsUnconfirmed { get; set; }
        public bool IsAffected { get; set; }
    }

    public class GanttChartPRInterval
    {
        public int AssignPMState { get; set; }
        public string PMIntervalST { get; set; }
        public string PMIntervalET { get; set; }
    }

    #region Part of Service/PMCSchedule

    public class Chromsome
    {
        public string SeriesID { get; set; }
        public string OrderID { get; set; }
        public double OPID { get; set; }
        public int Range { get; set; }
        public string WorkGroup { get; set; }
        public int Delay { get; set; }
        public TimeSpan Duration { get; set; }
        public DateTime AssignDate { get; set; }
        public DateTime StartTime { get; set; }
        public DateTime EndTime { get; set; }
        public string Maktx { get; set; }
        public int PartCount { get; set; }
        public int EachMachineSeq { get; set; }

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }

    public class GaSchedule
    {
        public string SeriesID { get; set; }
        public string OrderID { get; set; }
        public double OPID { get; set; }
        public int Range { get; set; }
        public TimeSpan Duration { get; set; }
        public DateTime Assigndate { get; set; }
        public string Maktx { get; set; }
        public int PartCount { get; set; }
        public int Scheduled { get; set; }
    }

    public class LocalMachineSeq
    {
        public string SeriesID { get; set; }
        public string OrderID { get; set; }
        public double OPID { get; set; }
        public int Range { get; set; }
        public TimeSpan Duration { get; set; }
        public DateTime PredictTime { get; set; }
        public int PartCount { get; set; }
        public string Maktx { get; set; }
        public string WorkGroup { get; set; }
        public int EachMachineSeq { get; set; }
    }

    public class Evafitnessvalue
    {
        public int Idx { get; set; }
        public double Fitness { get; set; }

        public Evafitnessvalue(int Idx, double Fitness)
        {
            this.Idx = Idx;
            this.Fitness = Fitness;
        }
    }

    public class GanttChartData
    {
        /// <summary>
        /// 起始時間
        /// </summary>
        public string startDate { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        public string endDate { get; set; }

        /// <summary>
        /// 機台名稱
        /// </summary>
        public string taskName { get; set; }

        /// <summary>
        /// 製程狀態【0:未開工、1:加工中、2:暫停中、3:已完成】
        /// </summary>
        public string status { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public double op { get; set; }

        /// <summary>
        /// 工單編號
        /// </summary>
        public string orderNum { get; set; }
    }

    public class EditOrderModels
    {
        public string SeriesID { get; set; }

        public string OrderID { get; set; }

        public double OPID { get; set; }

        public string StartTime { get; set; }

        public string WorkGroup { get; set; }
    }

    public class ProcessCR
    {
        public string OrderID { get; set; }
        public double CR { get; set; }
    }

    #endregion

    #region Rayn's calss

    /// <summary>
    /// OPdetail API資料格式
    /// </summary>
    public class DetailData
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 總數量
        /// </summary>
        public int Total { get; set; }
        /// <summary>
        /// 預計開始時間
        /// </summary>
        public string PredictST { get; set; }
        /// <summary>
        /// 預計結束時間
        /// </summary>
        public string PredictET { get; set; }
        /// <summary>
        /// 生管預交日期
        /// </summary>
        public string AssignDate_PM { get; set; }
        /// <summary>
        /// 預交日期
        /// </summary>
        public string AssignDate { get; set; }
        ///// <summary>
        ///// 訂單交付日期
        ///// </summary>
        //public string OrderDelivDay { get; set; }
        /// <summary>
        /// 生產完成度
        /// </summary>
        public string Progress { get; set; }
        ///// <summary>
        ///// 實際執行時間
        ///// </summary>
        //public string ExeTime { get; set; }
        /// <summary>
        /// 延遲天數
        /// </summary>
        public int DelayDay { get; set; }
        /// <summary>
        /// WT值
        /// </summary>
        public double WT { get; set; }
        /// <summary>
        /// CT值
        /// </summary>
        public double CT { get; set; }
    }

    /// <summary>
    /// 製程狀態資料格式(List)
    /// </summary>
    public class ListWorkConditions
    {
        /// <summary>
        /// 製程結束時間
        /// </summary>
        public string Endtime { get; set; }
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 提醒訊息
        /// </summary>
        public string Message { get; set; }
        /// <summary>
        /// 是否延遲(用於顏色辨認)
        /// </summary>
        public bool IsOPDeley { get; set; }
    }

    /// <summary>
    /// 提醒訊息格式
    /// </summary>
    public class WorkInfos
    {
        public string OrderID { get; set; }
        public string OPID { get; set; }
    }


    /// <summary>
    /// 製程狀態資料格式(Chart-基本資訊)
    /// </summary>
    public class BasicChartWork
    {
        /// <summary>
        /// 預計開始時間
        /// </summary>
        public string AssignST { get; set; }
        /// <summary>
        /// 預計結束時間
        /// </summary>
        public string AssignET { get; set; }
        /// <summary>
        /// 製程狀態(未開工、進行中、已暫停、已完成)
        /// </summary>
        public string OPState { get; set; }
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 製程順序
        /// </summary>
        public int Range { get; set; }
        /// <summary>
        /// 機台編號
        /// </summary>
        public string WorkGroup { get; set; }
    }

    /// <summary>
    /// 製程狀態資料格式(Chart-基本資訊)
    /// </summary>
    public class BasicChartWorkConditions
    {
        /// <summary>
        /// 群組編號(委外機台)
        /// </summary>
        public string GroupID { get; set; }
        /// <summary>
        /// 預計開始時間
        /// </summary>
        public string AssignST { get; set; }
        /// <summary>
        /// 預計結束時間
        /// </summary>
        public string AssignET { get; set; }
        /// <summary>
        /// 製程狀態(UNSTARTED:未開工、PROCESSING:進行中、FINISH:已完成)
        /// </summary>
        public string OPSTATE { get; set; }
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 機台編號
        /// </summary>
        public string WorkGroup { get; set; }
        /// <summary>
        /// 機檯群組
        /// </summary>
        public string GroupNumber { get; set; }
        ///// <summary>
        ///// 機台狀態
        ///// </summary>
        //public string MachineState { get; set; }

        /// <summary>
        /// 目前執行進度
        /// </summary>
        public int Progress { get; set; }

        /// <summary>
        /// 預計交期
        /// </summary>
        public string DueDate { get; set; }
        /// <summary>
        /// OP狀態屬性
        /// </summary>
        public BasicOPstates OPICON { get; set; }
        /// <summary>
        /// OP實際預計的屬性
        /// </summary>
        public ActExpCompare actExpCompare { get; set; }
        /// <summary>
        /// 是否要鎖定此區塊
        /// </summary>
        public bool isLock { get; set; }
    }

    /// <summary>
    /// OP的狀態
    /// </summary>
    public class BasicOPstates
    {
        ///// <summary>
        ///// 即將延遲icon
        ///// </summary>
        //public bool AssignDelayIsComing { get; set; }

        ///// <summary>
        ///// 已經延遲icon
        ///// </summary>
        //public bool OrderHasDelayed { get; set; }

        /// <summary>
        /// 製程狀態【DELAY:已延遲、WILLDELAY:即將延遲、NORMAL:進度正常】
        /// </summary>
        public string ProcessStatus { get; set; }

        /// <summary>
        /// 重要工單icon
        /// </summary>
        public bool IsImportant { get; set; }

        /// <summary>
        /// 工單緊急程度(0:非急工單, 1:急, 2:緊急, 3:最緊急)
        /// </summary>
        public int OrderPriority { get; set; }

        /// <summary>
        /// 待確認製程樣式
        /// </summary>
        public bool IsUnconfirmed { get; set; }

        /// <summary>
        /// 受影響製程樣式
        /// </summary>
        public bool IsAffected { get; set; }
    }

    /// <summary>
    /// 生管預交狀態，顯示實際報工與計畫排程相比較之差值
    /// </summary>
    public class ActExpCompare
    {
        /// <summary>
        /// 實際/預計延遲狀態(提早完成:0、準時:1、延遲完成:2)
        /// </summary>
        public int AssignPMState { get; set; }
        /// <summary>
        /// 實際報工提早/延遲開始時間
        /// </summary>
        public string PMIntervalST { get; set; }
        /// <summary>
        /// 實際報工提早/延遲結束時間
        /// </summary>
        public string PMIntervalET { get; set; }
    }


    public class BasicChartOPData
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }

        /// <summary>
        /// 分發機台
        /// </summary>
        public string Workgroup { get; set; }

        /// <summary>
        /// 機檯群組
        /// </summary>
        public string GroupNumber { get; set; }

        /// <summary>
        /// 機台狀態
        /// </summary>
        public string MachineState { get; set; }

        /// <summary>
        /// 排程開始時間
        /// </summary>
        public DateTime AssignST { get; set; }

        /// <summary>
        /// 排程結束時間
        /// </summary>
        public DateTime AssignET { get; set; }

        /// <summary>
        ///實際開始時間 
        /// </summary>
        public string Real_ST { get; set; }

        /// <summary>
        /// 
        /// </summary>
        public string Real_ET { get; set; }

        /// <summary>
        /// 生管預交日
        /// </summary>
        public string PMDate { get; set; }

        /// <summary>
        /// 預交日期
        /// </summary>
        public string ASDate { get; set; }

        /// <summary>
        /// 工單狀態
        /// </summary>
        public string WIPEvent { get; set; }

        /// <summary>
        /// 緊急程度
        /// </summary>
        public string Priority { get; set; }

        /// <summary>
        /// 重要工單
        /// </summary>
        public string Important { get; set; }

        /// <summary>
        /// 執行進度
        /// </summary>
        public int Progress { get; set; }


        /// <summary>
        /// 預交日期
        /// </summary>
        public string DueDate { get; set; }
    }




    public class OP_Data
    {
        public string OrderID { get; set; }
        public string OPID { get; set; }
        public string Assign_ST { get; set; }
        public string Assign_ET { get; set; }
        public string Real_ST { get; set; }
        public string Real_ET { get; set; }
        public string Workgroup { get; set; }
    }

    /// <summary>
    /// 製程狀態資料格式(Chart-進度資訊)
    /// </summary>
    public class ProgressChartWorkConditions
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string WorkGroup { get; set; }
        /// <summary>
        /// 製程進度資訊
        /// </summary>
        public List<ProgressOPstates> OPConditions { get; set; }
    }

    public class ProgressOPstates
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }
        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }
        /// <summary>
        /// 進度狀態(1:進度正常、2:進度延遲)
        /// </summary>
        public int ProgressState { get; set; }
        /// <summary>
        /// 進度開始時間
        /// </summary>
        public string ProgressST { get; set; }
        /// <summary>
        /// 進度結束時間
        /// </summary>
        public string ProgressET { get; set; }
        /// <summary>
        /// 是否要有選擇執行方式ICON
        /// </summary>
        public bool ChooseMethod { get; set; }
    }

    public class ProgressChartOPData
    {
        public string OrderID { get; set; }
        public string OPID { get; set; }
        public string Workgroup { get; set; }
        public string AssignST { get; set; }
        public string AssignET { get; set; }
        public string Real_ST { get; set; }
        public string Real_ET { get; set; }
        public string WIPEvent { get; set; }
    }

    public class AllMachine
    {
        /// <summary>
        /// 機台列表
        /// </summary>
        public List<string> Machine { get; set; }
    }

    /// <summary>
    /// 生管預交修改資訊
    /// </summary>
    public class EditPredictDateData
    {
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderID { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPID { get; set; }

        /// <summary>
        /// 預交日期
        /// </summary>
        public string AssignDate { get; set; }

        /// <summary>
        /// 分發數量
        /// </summary>
        public string OrderQTY { get; set; }

        /// <summary>
        /// 人工時
        /// </summary>
        public string HumanOpTime { get; set; }

        /// <summary>
        /// 機工時
        /// </summary>
        public string MachOpTime { get; set; }
    }

    #endregion
}
