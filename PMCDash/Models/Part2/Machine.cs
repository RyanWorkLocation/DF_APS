using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
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
		public string OrderID { get; set; }

		public string OPID { get; set; }

		public string WIPEvent { get; set; }

		public string Progress { get; set; }

		public LastUpdateOP()
		{
			this.OrderID = "";
			this.OPID = "";
			this.WIPEvent = "";
			this.Progress = "";
		}
	}
}

