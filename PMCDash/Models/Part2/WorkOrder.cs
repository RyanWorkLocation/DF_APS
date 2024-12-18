using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
    //public class WorkOrder
    //{
    //    /// <summary>
    //    /// 工單資料
    //    /// </summary>
    //    /// <param name="_Id">工單編號</param>
    //    /// <param name="_Name">工單名稱</param>
    //    /// <param name="_RouteData">工單途程</param>
    //    /// <param name="_Remark">工單備註</param>
    //    public WorkOrder(string _Id, string _Name, Routing _RouteData, string _Remark)
    //    {
    //        Id = _Id;
    //        Name = _Name;
    //        RouteData = _RouteData;
    //        Remark = _Remark;
    //    }

    //    /// <summary>
    //    /// 工單編號
    //    /// </summary>
    //    public string Id { get; set; }

    //    /// <summary>
    //    /// 工單名稱
    //    /// </summary>
    //    public string Name { get; set; }

    //    /// <summary>
    //    /// 工單途程
    //    /// </summary>
    //    public Routing RouteData { get; set; }

    //    /// <summary>
    //    /// 工單備註
    //    /// </summary>
    //    public string Remark { get; set; }
    //}

    public class WorkOrderTaskDetail
    {
        public string MachineId { get; set; }
        public string WorkOrderId { get; set; }
        public string OPId { get; set; }
        public string SatrtTime { get; set; }
        public string EndTime { get; set; }
        public string Qty { get; set; }
        public string ActCount { get; set; }
        public string GroupName { get; set; }
        public string OperatorName { get; set; }
        public string CadFilePath { get; set; }
        public string Rrmark { get; set; }
    }

    public class WorkOrderOverview
    {
        public string WorkOrderID { get; set; }

        public double WorkOrderProgress { get; set; }

        public List<OPOverview> OP { get; set; }
    }



    public class OrderGroup
    {
        public string OrderID { get; set; }

        public double OrderProgress { get; set; }

        public List<WorkOrderOverview> WorkOrder { get; set; }
    }


}
