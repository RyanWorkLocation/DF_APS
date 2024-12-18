using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
    public class OP
    {
        /// <summary>
        /// 製程詳細資料
        /// </summary>
        /// <param name="_Id">製程編號</param>
        /// <param name="_Name">製程名稱</param>
        /// <param name="_SettingTime">準備工時</param>
        /// <param name="_StandardTime">標準工時</param>
        /// <param name="_MainDeviec">主要設備</param>
        /// <param name="_SubDevice">替代設備</param>
        public OP(string _Id, string _Name, string _SettingTime, string _StandardTime, string _MainDeviec, List<SubDevice> _SubDevice)
        {
            Id = _Id;
            Name = _Name;
            SettingTime = _SettingTime;
            StandardTime = _StandardTime;
            MainDeviec = _MainDeviec;
            SubDevices = _SubDevice;
        }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// 製程名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 準備工時
        /// </summary>
        public string SettingTime { get; set; }
        /// <summary>
        /// 標準工時
        /// </summary>
        public string StandardTime { get; set; }
        /// <summary>
        /// 主要設備
        /// </summary>
        public string MainDeviec { get; set; }
        /// <summary>
        /// 指定主程式
        /// </summary>
        public string MainNCProgram { get; set; }
        /// <summary>
        /// 替代設備
        /// </summary>
        public List<SubDevice> SubDevices { get; set; }

        /// <summary>
        /// 製作費用
        /// </summary>
        public string cost { get; set; }

        /// <summary>
        /// 製程檢驗項目
        /// </summary>
        public List<QCRule> QCRules { get; set; }

    }

    public class SubDevice
    {
        /// <summary>
        /// 替代設備
        /// </summary>
        /// <param name="_Id">設備編號</param>
        /// <param name="_Name">設備名稱</param>
        /// <param name="_Type">設備類別</param>
        /// <param name="_Group">歸屬群組</param>
        /// <param name="_Operator">負責人員</param>
        public SubDevice(string _Id, string _Name, string _Type, string _Group, string _Operator)
        {
            Id = _Id;
            Name = _Name;
            Type = _Type;
            Group = _Group;
            Operator = _Operator;
        }
        /// <summary>
        /// 設備編號
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// 設備名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 設備類別
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// 歸屬群組
        /// </summary>
        public string Group { get; set; }
        /// <summary>
        /// 負責人員
        /// </summary>
        public string Operator { get; set; }

    }

    public class QCRule
    {
        /// <summary>
        /// 中心值
        /// </summary>
        public string CL { get; set; }
        /// <summary>
        /// 管制上限
        /// </summary>
        public string UCL { get; set; }
        /// <summary>
        /// 管制下限
        /// </summary>
        public string LCL { get; set; }
        /// <summary>
        /// 規格上限
        /// </summary>
        public string USL { get; set; }
        /// <summary>
        /// 規格下限
        /// </summary>
        public string LSL { get; set; }
    }

    /// <summary>
    /// 製程概要
    /// </summary>
    public class OPOverview
    {
        //製程ID
        public string OPID { get; set; }
        //製程進度
        public double Progress { get; set; }
        //物料編號
        public string MAKTX { get; set; }
        //預繳日期
        public string AssignDate { get; set; }
        //生管預繳日期
        public string AssignDate_PM { get; set; }
        //備註
        public string Note { get; set; }
        //訂單原檔
        public string ImgPath { get; set; }
    }

    public class OPList
    {
        /// <summary>
        /// 訂單編號
        /// </summary>
        public string ERPorderID { get; set; }
        /// <summary>
        /// 工單編號
        /// </summary>
        public string OrderId { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string OPId { get; set; }

        /// <summary>
        /// 製程順序
        /// </summary>
        public int Range { get; set; }

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
        /// 實際開工
        /// </summary>
        public string wipStartTime { get; set; }

        /// <summary>
        /// 實際完工
        /// </summary>
        public string wipEndTime { get; set; }

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
        /// 延期天數，排序用
        /// </summary>
        public int Days { get; set; }

        /// <summary>
        /// 是否延遲
        /// </summary>
        public bool IsDeley { get; set; }

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
        /// 品質檢驗【N/A:不需檢驗、False:需檢驗、True:已經檢驗】
        /// </summary>
        public string IsQC { get; set; }

        /// <summary>
        /// 品檢人員
        /// </summary>
        public string QCman { get; set; }

    }

    public class OPDetail
    {
        /// <summary>
        /// 訂單編號
        /// </summary>
        public string ERPorderID { get; set; }
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
        /// 開工狀態【0:未開工、1:生產中、2:暫停中、3:已完成】
        /// </summary>
        public string OPStatus { get; set; }

        /// <summary>
        /// 預計開工
        /// </summary>
        public string StartTime { get; set; }

        /// <summary>
        /// 預計完工
        /// </summary>
        public string EndTime { get; set; }

        /// <summary>
        /// 實際開工
        /// </summary>
        public string wipStartTime { get; set; }


        /// <summary>
        /// 實際完工
        /// </summary>
        public string wipEndTime { get; set; }


        /// <summary>
        /// 產品交期
        /// </summary>
        public string AssignDate { get; set; }

        /// <summary>
        /// 延期交期
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
        public int Progress { get; set; }

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

        /// <summary>
        /// 不良品數量
        /// </summary>
        public string QtyBad { get; set; }

        /// <summary>
        /// 圖紙路徑
        /// </summary>
        public string ImgPath { get; set; }

        /// <summary>
        /// 檢驗項目列表
        /// </summary>
        public List<QC> QCList { get; set; }

    }

    public class QC
    {
        /// <summary>
        /// 量測項目
        /// </summary>
        public string QCPointName { get; set; }

        /// <summary>
        /// 量測數值
        /// </summary>
        public string QCPointValue { get; set; }

        /// <summary>
        /// 量測點狀態【合格、不合格】
        /// </summary>
        public string QCResult { get; set; }

        /// <summary>
        /// 量測點單位
        /// </summary>
        public string QCUnit { get; set; }

        /// <summary>
        /// 量測設備編號
        /// </summary>
        public string QCToolID { get; set; }

        /// <summary>
        /// 量測時間
        /// </summary>
        public string CreateTime { get; set; }

        /// <summary>
        /// 最後量測時間
        /// </summary>
        public string LastUpdateTime { get; set; }

        /// <summary>
        /// 量具誤差值
        /// </summary>
        public List<string> toolDiff { get; set; }

        /// <summary>
        /// 是否開放量測【Ture:可以量測、False:不可以量測】
        /// </summary>
        public bool Enable { get; set; }
    }
}
