using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class Device
    {
        [Display(Name = "機台編號")]
        public int ID { set; get; }

        [Display(Name = "機台名稱")]
        public string MachineName { set; get; }

        [Display(Name = "Remark")]
        public string Remark { set; get; }

        [Display(Name = "群組名稱")]
        public string GroupName { set; get; }

        [Display(Name = "作業員編號")]
        public string OperatorNO { set; get; }

        [Display(Name = "作業員姓名")]
        public string OperatorName { set; get; }
    }

    public class DeviceInfo
    {
        [Display(Name = "機台編號")]
        public int ID { set; get; }

        [Display(Name = "機台名稱")]
        public string MachineName { set; get; }

        [Display(Name = "Remark")]
        public string Remark { set; get; }

        [Display(Name = "群組名稱")]
        public string GroupName { set; get; }
    }

    public class Device_e
    {
        public Device_e(string value, string text)
        {
            Value = value;
            Text = text;
        }

        public string Value { get; set; }
        public string Text { get; set; }
    }

	public class Machine
	{

		/// <summary>
		/// 機台名稱
		/// </summary>
		public string MachineName { get; set; }

		/// <summary>
		/// 機台所屬群組
		/// </summary>
		public string MachingGroup { get; set; }

		/// <summary>
		/// 機台圖片
		/// </summary>
		public string MachineImg { get; set; }

		/// <summary>
		/// 操作人員名稱
		/// </summary>
		public string OperatorName { get; set; }

		/// <summary>
		/// 未開工製程數量
		/// </summary>
		public int NotStartNum { get; set; }

		/// <summary>
		/// 生產中製程數量
		/// </summary>
		public int RunningNum { get; set; }

		/// <summary>
		/// 暫停中製程數量
		/// </summary>
		public int SuspendNum { get; set; }

		/// <summary>
		/// 完成製程數量
		/// </summary>
		public int CompleteNum { get; set; }

		/// <summary>
		/// 進行中工單/最新完成工單
		/// </summary>
		public List<LastUpdateOP> LastUpdateOP { get; set; }
	}

	public class LastUpdateOP
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
		/// 工單報工狀態【0:未開工(灰色)、1:生產中(綠色)、2:暫停中(黃色)、3:已完成(灰色)】
		/// </summary>
		public string WIPEvent { get; set; }
		/// <summary>
		/// 製程進度
		/// </summary>
		public int Progress { get; set; }

		public LastUpdateOP()
		{
			this.OrderID = "";
			this.OPID = "";
			this.WIPEvent = "";
			this.Progress = 0;
		}
	}
}
