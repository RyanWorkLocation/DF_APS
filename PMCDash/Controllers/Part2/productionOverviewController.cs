using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NSwag.Annotations;
using PMCDash.Models;
using PMCDash.Models.Part2;


namespace PMCDash.Controllers.Part2
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductionOverviewController : BaseApiController
    {
        ConnectStr _ConnectStr = new ConnectStr();

        //資料庫連線
        //private readonly string _ConnectStr.Local = @"Data Source = 127.0.0.1; Initial Catalog = DPI; User ID = MES2014; Password = PMCMES;";
        //時間格式
        private readonly string _timeFormat = "yyyy-MM-dd HH:mm:ss";

        /// <summary>
        /// 取得所有訂單之Overview
        /// </summary>
        /// <returns></returns>
        [HttpPost("OrderOverview")]
        public List<OrderGroup> OrderOverview(string CPK = "", string WillDelay = "", string AlreadyDelay = "")
        {
            var result = new List<OrderGroup>();
            
            var SqlStr = @"SELECT * FROM AssignmentGroup as g ORDER BY g.ID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                string OrderID = SqlData["OrderID"].ToString();
                                var WorkOrder = WorkOrderOverview(OrderID, CPK, WillDelay, AlreadyDelay);
                                result.Add(new OrderGroup
                                {
                                    OrderID = OrderID,
                                    OrderProgress = WorkOrder.Item1,
                                    WorkOrder = WorkOrder.Item2
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 以訂單ID搜尋所有工單
        /// </summary>
        /// <param name="OrderID"></param>
        /// <returns></returns>
        private Tuple<double, List<WorkOrderOverview>> WorkOrderOverview(string OrderID, string CPK = "", string WillDelay = "", string AlreadyDelay = "")
        {
            var result = new List<WorkOrderOverview>();
            var SqlStr = @"SELECT g.OrderID, a.OrderID as WorkOrderID
                        FROM Assignment as a 
                        JOIN AssignmentGroup as g ON g.WorkOrderID = a.OrderID WHERE g.OrderID = @OrderID ORDER BY g.ID";
            double TotalProgress = 0;
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                string WorkOrderID = SqlData["WorkOrderID"].ToString();
                                var OP = OPOverview(WorkOrderID, CPK, WillDelay, AlreadyDelay);
                                result.Add(new WorkOrderOverview
                                {
                                    WorkOrderID = WorkOrderID,
                                    WorkOrderProgress = OP.Item1,
                                    OP = OP.Item2
                                });
                                TotalProgress += OP.Item1;
                            }
                        }
                    }
                }
            }
            TotalProgress /= result.Count;
            return Tuple.Create(TotalProgress, result);
        }

        /// <summary>
        /// 以工單ID搜尋所有製程
        /// </summary>
        /// <param name="WorkOrderID"></param>
        /// <param name="CPK"></param>
        /// <param name="WillDelay"></param>
        /// <param name="AlreadyDelay"></param>
        /// <returns></returns>
        private Tuple<double,List<OPOverview>> OPOverview(string WorkOrderID, string CPK = "", string WillDelay = "", string AlreadyDelay = "")
        {
            var result = new List<OPOverview>();
            var SqlStr = @"SELECT a.OrderID as WorkOrderID, a.OPID,
                        Progress = cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as int),
                        a.MAKTX, a.AssignDate, a.AssignDate_PM, a.ImgPath, a.Note
                        FROM Assignment as a 
                        JOIN WIP as w ON (a.OrderID = w.OrderID AND a.OPID = w.OPID)
                        WHERE a.OrderID = @WorkOrderID
                        ";
            double TotalProgress = 0;
            //選出 CPK
            if (CPK != "")
            {
                if (CPK != "1" && CPK != "2" && CPK != "3" && CPK != "4")
                {
                    CPK = "0";
                }
                else if (CPK == "1")
                {
                    SqlStr += " AND CPK >= 2";
                }
                else if (CPK == "2")
                {
                    SqlStr += " AND CPK < 2 AND CPK >=1.33";
                }
                else if (CPK == "3")
                {
                    SqlStr += " AND CPK < 1.33 ";
                }
                else if (CPK == "4")
                {
                    SqlStr += "AND CPK < 1.33";
                }
            }
            //選出即將延遲或已經延遲
            if (((WillDelay != "" && AlreadyDelay != "") && (WillDelay == "1" && AlreadyDelay == "1")) || (AlreadyDelay != "" && AlreadyDelay == "1"))
            {
                SqlStr += " AND a.EndTime > a.AssignDate";
            }
            else if (WillDelay != "" && WillDelay == "1")
            {
                SqlStr += " AND a.EndTime > a.AssignDate_PM";
            }
            SqlStr += " ORDER BY a.OPID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@WorkOrderID"), SqlDbType.NVarChar).Value = WorkOrderID;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new OPOverview
                                {
                                    OPID = SqlData["OPID"].ToString(),
                                    Progress = Convert.ToDouble(SqlData["Progress"]),
                                    MAKTX = SqlData["MAKTX"].ToString(),
                                    AssignDate = SqlData["AssignDate"].ToString(),
                                    AssignDate_PM = SqlData["AssignDate_PM"].ToString(),
                                    Note = SqlData["Note"].ToString(),
                                    ImgPath = SqlData["ImgPath"].ToString()
                                });
                                TotalProgress += Convert.ToDouble(SqlData["Progress"]);
                            }
                        }
                    }
                }
            }
            TotalProgress /= result.Count;
            return Tuple.Create(TotalProgress, result);
        }


        /// <summary>
        /// 取得訂單進度列表
        /// </summary>
        [ApiExplorerSettings(IgnoreApi = true)]
        public void Get_OrderProcessRateList()
        {
            //工單編號
            //製程編號
            //檢視製程能力分析
            //物料編號
            //訂單生產進度條
            //工單生產進度條
            //製程生產進度條
            //訂單預交日
            //生管預交日
            //訂單狀態燈號
            //訂單原始檔檢視
            //備註欄

        }

        /// <summary>
        /// 取得單一物料製程能力分分析資料
        /// </summary>
        [ApiExplorerSettings(IgnoreApi = true)]
        public void Get_PartOfOP_SPCAnalysisDetail()
        {
            //物料編號
            //製程編號
            //量測點
            //CPK評價
            //CPK折線圖
            //CPK相關數據
            //詳細規格資料

        }

        /// <summary>
        /// 取得單一物料製程能力CPK歷史資料
        /// </summary>
        [ApiExplorerSettings(IgnoreApi = true)]
        public void Get_PartOfOP_CPKHistory()
        {
            //X軸批次/件數
            //Y軸規格數據
            //預設最新30筆數據
        }

        /// <summary>
        /// 取得最新CPK評價
        /// </summary>
        [ApiExplorerSettings(IgnoreApi = true)]
        public void Get_PartOfOP_CPKLeavel()
        {
            //CPK係數
            //CPK評價訊息
            //CPK燈號

        }
    }

}