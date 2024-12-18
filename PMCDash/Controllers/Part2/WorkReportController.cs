using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using PMCDash.Models;
using PMCDash.Models.Part2;
using PMCDash.Services;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using static PMCDash.Controllers.Part2.QCWorkReportController;
using static PMCDash.Services.AccountService;

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class WorkReportController : BaseApiController
    //public class ScheduleController : ControllerBase
    {
        ConnectStr _ConnectStr = new ConnectStr();
        //資料庫連線
        //private readonly string _ConnectStr.Local = @"Data Source = 127.0.0.1; Initial Catalog = DPI; User ID = MES2014; Password = PMCMES;";
        //private readonly string _ConnectStr.Local = @"Data Source = 192.168.0.156; Initial Catalog = DPI;User ID = MES2014; Password = PMCMES;";

        //時間格式
        private readonly string _timeFormat = "yyyy-MM-dd HH:mm:ss";
        private readonly string _dateFormat = "yyyy-MM-dd";

        /// <summary>
        /// 列出所有機台
        /// </summary>
        /// <param></param>
        /// <returns></returns>
        [HttpPost("MachineList_v2")]
        public List<Machine> MachineList_v2()
        {
            //var UserPower = checkLoginPermission();//檢查登入權限

            var UserName = User.Identity.Name;
            string baseUrl = "http://" + Request.Host.Value + "/DeviceImages/";
            //string baseUrl = @"http://localhost:8088/DeviceImages/";
            var result = new List<Machine>();
            var LastUpdateOP = new List<LastUpdateOP>();
            var sqlStr = "";

            sqlStr = @$"SELECT 
                             a.*
                            ,d.remark,d.GroupName, d.img as MachineImg,u.[user_name]
                            ,w.*
                            ,uu.[user_name] as worker
                            FROM {_ConnectStr.APSDB}.[dbo].Assignment as a
                            LEFT JOIN　{_ConnectStr.APSDB}.[dbo].[Device] as d ON d.remark = a.WorkGroup
                            inner join {_ConnectStr.AccuntDB}.dbo.GroupDevice as j on d.ID=j.DeviceId
                            LEFT JOIN {_ConnectStr.AccuntDB}.dbo.[User] as u ON u.usergroup_id=j.GroupSeq
                            LEFT JOIN {_ConnectStr.APSDB}.[dbo].[WIP] as w ON (a.OrderID = w.OrderID and a.OPID = w.OPID)
                            left join {_ConnectStr.AccuntDB}.dbo.[User] as uu on a.Operator=uu.[user_id]
                            WHERE u.[user_account] = @user_id
                            ORDER BY a.WorkGroup";
            //sqlStr = $@"select d.*,c.remark,c.GroupName, c.img as MachineImg 
            //            ,a.[user_name],e.WIPEvent
            //            from {_ConnectStr.AccuntDB}.dbo.[User] as a
            //            inner join {_ConnectStr.AccuntDB}.dbo.GroupDevice as b on a.usergroup_id=b.GroupSeq
            //            inner join {_ConnectStr.APSDB}.[dbo].Device as c on b.DeviceId=c.ID
            //            inner join {_ConnectStr.APSDB}.[dbo].Assignment as d on c.remark=d.WorkGroup
            //            inner join {_ConnectStr.APSDB}.[dbo].WIP as e on d.OrderID=e.OrderID and d.OPID=e.OPID
            //            where a.user_account=@user_id
            //            ORDER BY d.WorkGroup";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@user_id", SqlDbType.NVarChar).Value = UserName;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            string tmp = "";
                            bool first = false;
                            string MachineNameTmp = "";
                            string MachineGroupTmp = "";
                            string MachineImgTmp = "";
                            string OperatorName = "";
                            int NotStartCount = 0;
                            int RunningCount = 0;
                            int SuspendCount = 0;
                            int CompleteCount = 0;
                            while (SqlData.Read())
                            {
                                if (!SqlData["WorkGroup"].Equals(tmp) && first)
                                {
                                    LastUpdateOP = RunningOPStatus(MachineNameTmp);
                                    //NotStartCount = CanWorkOrderCount(MachineNameTmp);//判斷是否有單可以執行
                                    //if (NotStartCount != 0 && RunningCount != 0 && SuspendCount != 0 && CompleteCount != 0)//若都沒有供單可執行就不顯示機台，節省效能
                                    //{
                                    result.Add(new Machine
                                    {
                                        MachineName = MachineNameTmp,
                                        MachingGroup = MachineGroupTmp,
                                        MachineImg = checkImgExit(baseUrl + MachineImgTmp),
                                        OperatorName = OperatorName,
                                        NotStartNum = NotStartCount,
                                        RunningNum = RunningCount,
                                        SuspendNum = SuspendCount,
                                        CompleteNum = CompleteCount,
                                        LastUpdateOP = LastUpdateOP
                                    });
                                    //}
                                    tmp = SqlData["WorkGroup"].ToString();
                                    NotStartCount = 0; RunningCount = 0; SuspendCount = 0; CompleteCount = 0;

                                }
                                else
                                {
                                    first = true;
                                    tmp = SqlData["WorkGroup"].ToString();
                                }
                                MachineNameTmp = SqlData["remark"].ToString();
                                MachineGroupTmp = SqlData["GroupName"].ToString();
                                MachineImgTmp = SqlData["MachineImg"].ToString();
                                OperatorName = string.IsNullOrEmpty(SqlData["worker"].ToString()) ? "-" : SqlData["worker"].ToString();
                                int Event = Convert.ToInt16(SqlData["WIPEvent"]);
                                if (Event == 0)
                                {
                                    NotStartCount++;
                                }
                                else if (Event == 1)
                                {
                                    RunningCount++;
                                }
                                else if (Event == 2)
                                {
                                    SuspendCount++;
                                }
                                else if (Event == 3)
                                {
                                    CompleteCount++;
                                }
                            }

                            LastUpdateOP = RunningOPStatus(MachineNameTmp);
                            result.Add(new Machine
                            {
                                MachineName = MachineNameTmp,
                                MachingGroup = MachineGroupTmp,
                                MachineImg = checkImgExit(baseUrl + MachineImgTmp),
                                OperatorName = OperatorName,
                                NotStartNum = NotStartCount,//CanWorkOrderCount(MachineNameTmp),
                                RunningNum = RunningCount,
                                SuspendNum = SuspendCount,
                                CompleteNum = CompleteCount,
                                LastUpdateOP = LastUpdateOP
                            });
                        }
                    }
                }
            }

            foreach (var item in result)
            {
                if (item.LastUpdateOP.Exists(x => x.WIPEvent == "3" || x.WIPEvent == "0" || x.WIPEvent == ""))
                {
                    item.LastUpdateOP.Find(x => x.WIPEvent == "3" || x.WIPEvent == "0" || x.WIPEvent == "").OrderID = "請點選開工";
                    item.LastUpdateOP.Find(x => x.WIPEvent == "3" || x.WIPEvent == "0" || x.WIPEvent == "").OPID = "N/A";
                    item.OperatorName = "-";
                }
                //if (item.NotStartNum == 0 && item.RunningNum == 0 && item.SuspendNum == 0)
                //    result.Remove(item);//先刪掉沒有任何工作的機台
            }

            return result;
        }

        [HttpPost("MachineList")]
        public List<Machine> MachineList()
        {
            var result = new List<Machine>();
            UserData userdata = UserInfo();
            string baseUrl = "http://" + Request.Host.Value + "/DeviceImages/";
            //string baseUrl = "http://" + "127.0.0.1:8088" + "/DeviceImages/";//測試用

            //List<string> devices = getcanusedevice(userdata.GroupId);

            var temp = new List<ordercount>();
            //取得工單資料
            var sqlStr = @$";WITH cte AS
                        (
                           SELECT *,
                                 ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY [Range]) AS rn
                           FROM 
                           (select 
	                        a.OrderID,a.OPID,a.[Range],a.OPLTXA1,a.MAKTX,c.[Name],a.StartTime,a.EndTime,a.AssignDate,a.WorkGroup ,d.GroupName
	                        ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=a.Operator) as [user_name]
	                        ,Progress = cast( (cast(b.QtyGood as float) + cast(b.QtyBad as float)) / cast(a.OrderQTY as float) * 100 as int)
	                        ,RemainingCount = (a.OrderQTY - b.QtyTol)
	                        ,a.OrderQTY,b.QtyGood,b.QtyBad,b.WIPEvent
                            ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=f.QCman) as QCman
                            ,e.GroupSeq	         
							,d.ID as DeviceID
                            from {_ConnectStr.APSDB}.[dbo].Assignment as a
	                        inner join {_ConnectStr.APSDB}.[dbo].WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
	                        inner join {_ConnectStr.MRPDB}.dbo.Part as c on a.MAKTX=c.Number
	                        inner join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
	                        left join {_ConnectStr.AccuntDB}.dbo.GroupDevice as e on d.ID=e.DeviceId
                            left join {_ConnectStr.APSDB}.[dbo].QCAssignment as f on a.OrderID=f.WorkOrderID and a.OPID=f.OPID
	                        where GroupSeq=2 and a.WorkGroup is not null) as a
                        )
                        SELECT *
                        FROM cte as aa
						left join {_ConnectStr.AccuntDB}.dbo.GroupDevice as bb on aa.DeviceID=bb.DeviceId
                        where bb.GroupSeq=@GroupName
                        order by OrderID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    //comm.Parameters.Add("@user_id", SqlDbType.NVarChar).Value = User.Identity.Name;
                    comm.Parameters.Add("@GroupName", SqlDbType.Int).Value = userdata.GroupId;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                temp.Add(new ordercount
                                {
                                    orderid = SqlData["OrderID"].ToString().Trim(),
                                    opid = SqlData["OPID"].ToString().Trim(),
                                    range = int.Parse(SqlData["Range"].ToString().Trim()),
                                    DeviceGroup = SqlData["GroupName"].ToString().Trim(),
                                    workgroup = SqlData["WorkGroup"].ToString().Trim(),
                                    Operator = string.IsNullOrEmpty(SqlData["user_name"].ToString().Trim()) ? "N/A" : SqlData["user_name"].ToString().Trim(),
                                    wipevent = SqlData["WIPEvent"].ToString().Trim()
                                });
                            };
                        }
                    }
                }
            }

            List<string> devices = temp.Select(x => x.workgroup).Distinct().ToList();
            foreach (var item in devices)
            {
                result.Add(new Machine
                {
                    MachineName = item,
                    MachingGroup = temp.Where(x => x.workgroup == item).First().DeviceGroup,
                    MachineImg = baseUrl + item + ".jpg",
                    //MachineImg = baseUrl + "DEMO_MC.jpg",
                    OperatorName = OperatorName(item),
                    NotStartNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "0").Count,
                    RunningNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "1").Count,
                    SuspendNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "2").Count,
                    CompleteNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "3").Count,
                    LastUpdateOP = LastUpdateOPStatus(item)
                });

                //if (!(temp.FindAll(x => x.workgroup == item && x.wipevent == "0").Count == 0 &&
                //    temp.FindAll(x => x.workgroup == item && x.wipevent == "1").Count == 0 &&
                //    temp.FindAll(x => x.workgroup == item && x.wipevent == "2").Count == 0 &&
                //    temp.FindAll(x => x.workgroup == item && x.wipevent == "3").Count == 0))
                //    result.Add(new Machine
                //    {
                //        MachineName = item,
                //        MachingGroup = "",
                //        MachineImg = baseUrl + item + ".jpg",
                //        OperatorName = OperatorName(item),
                //        NotStartNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "0").Count,
                //        RunningNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "1").Count,
                //        SuspendNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "2").Count,
                //        CompleteNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "3").Count,
                //        LastUpdateOP = LastUpdateOPStatus(item)
                //    });
                //else
                //    result.Add(new Machine
                //    {
                //        MachineName = item,
                //        MachingGroup = "",
                //        MachineImg = baseUrl + item + ".jpg",
                //        OperatorName = OperatorName(item),
                //        NotStartNum = 0,//temp.FindAll(x => x.workgroup == item && x.wipevent == "0").Count,
                //        RunningNum = 0,//temp.FindAll(x => x.workgroup == item && x.wipevent == "1").Count,
                //        SuspendNum = 0,//temp.FindAll(x => x.workgroup == item && x.wipevent == "2").Count,
                //        CompleteNum = 0,//temp.FindAll(x => x.workgroup == item && x.wipevent == "3").Count,
                //        LastUpdateOP = LastUpdateOPStatus(item)
                //    });
            }

            //temp.Clear();

            //sqlStr = @$";WITH cte AS
            //            (
            //               SELECT *,
            //                     ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY [Range]) AS rn
            //               FROM 
            //               (select 
            //             a.OrderID,a.OPID,a.[Range],a.OPLTXA1,a.MAKTX,c.[Name],a.StartTime,a.EndTime,a.AssignDate,a.WorkGroup ,d.GroupName
            //             ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=a.Operator) as [user_name]
            //             ,Progress = cast( (cast(b.QtyGood as float) + cast(b.QtyBad as float)) / cast(a.OrderQTY as float) * 100 as int)
            //             ,RemainingCount = (a.OrderQTY - b.QtyTol)
            //             ,a.OrderQTY,b.QtyGood,b.QtyBad,b.WIPEvent
            //                ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=f.QCman) as QCman
            //             from Assignment as a
            //             inner join WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
            //             inner join {_ConnectStr.MRPDB}.dbo.Part as c on a.MAKTX=c.Number
            //             inner join Device as d on a.WorkGroup=d.remark
            //             inner join {_ConnectStr.AccuntDB}.dbo.GroupDevice as e on d.ID=e.DeviceId
            //                left join QCAssignment as f on a.OrderID=f.WorkOrderID and a.OPID=f.OPID
            //             WHERE b.WIPEvent!=0 and e.GroupSeq=@GroupName ) as a
            //            )
            //            SELECT *
            //            FROM cte
            //            order by OrderID";
            //using (var conn = new SqlConnection(_ConnectStr.Local))
            //{
            //    using (var comm = new SqlCommand(sqlStr, conn))
            //    {
            //        if (conn.State != ConnectionState.Open)
            //            conn.Open();
            //        //comm.Parameters.Add("@user_id", SqlDbType.NVarChar).Value = User.Identity.Name;
            //        comm.Parameters.Add("@GroupName", SqlDbType.NVarChar).Value = userdata.GroupId;
            //        using (SqlDataReader SqlData = comm.ExecuteReader())
            //        {
            //            if (SqlData.HasRows)
            //            {
            //                while (SqlData.Read())
            //                {
            //                    temp.Add(new ordercount
            //                    {
            //                        orderid = SqlData["OrderID"].ToString().Trim(),
            //                        opid = SqlData["OPID"].ToString().Trim(),
            //                        range = int.Parse(SqlData["Range"].ToString().Trim()),
            //                        DeviceGroup = SqlData["GroupName"].ToString().Trim(),
            //                        workgroup = SqlData["WorkGroup"].ToString().Trim(),
            //                        Operator = string.IsNullOrEmpty(SqlData["user_name"].ToString().Trim()) ? "N/A" : SqlData["user_name"].ToString().Trim(),
            //                        wipevent = SqlData["WIPEvent"].ToString().Trim()
            //                    });
            //                };
            //            }
            //        }
            //    }
            //}

            //devices = temp.Select(x => x.workgroup).Distinct().ToList();
            //foreach (var item in devices)
            //{
            //    if (result.Exists(x => x.MachineName == item))
            //    {
            //        result.Find(x => x.MachineName == item).RunningNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "1").Count;
            //        result.Find(x => x.MachineName == item).SuspendNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "2").Count;
            //        result.Find(x => x.MachineName == item).CompleteNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "3").Count;
            //    }
            //    else
            //        result.Add(new Machine
            //        {
            //            MachineName = item,
            //            MachingGroup = "",
            //            MachineImg = baseUrl + item + ".jpg",
            //            OperatorName = OperatorName(item),
            //            RunningNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "1").Count,
            //            SuspendNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "2").Count,
            //            CompleteNum = temp.FindAll(x => x.workgroup == item && x.wipevent == "3").Count,
            //            LastUpdateOP = LastUpdateOPStatus(item)
            //        });
            //}


            //foreach (var item in result)
            //{
            //    if (item.LastUpdateOP.Exists(x => x.WIPEvent == "3" || x.WIPEvent == "0" || x.WIPEvent == ""))
            //    {
            //        item.LastUpdateOP.Find(x => x.WIPEvent == "3" || x.WIPEvent == "0" || x.WIPEvent == "").OrderID = "請點選開工";
            //        item.LastUpdateOP.Find(x => x.WIPEvent == "3" || x.WIPEvent == "0" || x.WIPEvent == "").OPID = "N/A";
            //        item.OperatorName = "-";
            //    }
            //}

            return result.OrderBy(x => x.MachineName).ToList();
        }

        private List<string> getcanusedevice(long groupId)
        {
            List<string> result = new List<string>();
            string SqlStr = $@"SELECT b.remark
                              FROM {_ConnectStr.AccuntDB}.[dbo].[GroupDevice] as a
                              inner join {_ConnectStr.APSDB}.dbo.Device as b on a.DeviceId = b.ID
                              where a.GroupSeq = @GroupName";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@GroupName"), SqlDbType.NVarChar).Value = groupId;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(SqlData["remark"].ToString().Trim());
                            }
                        }
                    }
                }
            }

            return result;
        }

        //取得尚未開工的工單數量
        private int noStartNum(List<ordercount> canworklist)
        {
            int result = 0;
            foreach (var item in canworklist)
            {
                if (iscanwork(item.orderid, item.range))
                    result++;
            }
            return result;
        }

        //判斷該工單的"前"製程是否已經開工，若前製程非未開工則代表目前製程可以開工
        private bool iscanwork(string orderid, int range)
        {
            if (range == 0)
                return true;

            //該工單的前一到製程，也就是Range-1的製程
            int Range = 0;
            if ((range - 1) <= 0)
                Range = 0;
            else
                Range = range - 1;

            bool result = false;

            string sqlStr = @$"SELECT a.OrderID,a.OPID,a.[Range], a.WorkGroup,b.WIPEvent FROM {_ConnectStr.APSDB}.[dbo].[Assignment] as a
                                inner join {_ConnectStr.APSDB}.[dbo].WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                                where a.OrderID=@OrderID and a.[Range]=@Range and b.WIPEvent!=0
                                order by [Range]";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = orderid;
                    comm.Parameters.Add("@Range", SqlDbType.Float).Value = Range;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            result = true;
                        }
                    }
                }
            }

            return result;
        }

        //取得該機台最近一次操作人員的名字
        private string OperatorName(string DeviceID)
        {
            string result = "N/A";

            string sqlStr = @$"select * from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=(
                                SELECT top(1) StaffID
                                  FROM {_ConnectStr.APSDB}.[dbo].[WIPLog]
                                  where DeviceID=@DeviceID
                                  order by CreateTime desc
                                  )";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@DeviceID", SqlDbType.NVarChar).Value = DeviceID;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result = SqlData["user_name"].ToString().Trim();
                            }
                        }
                    }
                }
            }

            return result;
        }

        private long checkLoginPermission()
        {
            long groupid = 0;
            var result = UserInfo();
            if (!result.FunctionAccess.Exists(x => x.FunctionName == "機台報工"))
                throw new UnauthorizedAccessException($"Some message");
            groupid = result.GroupId;
            return groupid;
        }

        private string checkImgExit(string url)
        {
            string result = "http://" + Request.Host.Value + "/DeviceImages/default.jpg";
            //string result = "http://" + "127.0.0.1:8088" + "/DeviceImages/default.jpg";//測試用
            try
            {
                WebRequest request = WebRequest.Create(url);
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();

                if (response.StatusCode == HttpStatusCode.OK)
                    result = url;
            }
            catch (Exception)
            {
                return result;

            }
            return result;
        }

        private List<LastUpdateOP> RunningOPStatus(string MachineID)
        {
            var result = new List<LastUpdateOP>();

            //string SqlStr = @$"SELECT TOP (1) w.OrderID,w.OPID, w.WIPEvent, w.UpdateTime,
            //                Progress = cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as int), RemainingCount = (w.OrderQTY - w.QtyTol)
            //                FROM {_ConnectStr.APSDB}.[dbo].WIP as w
            //                WHERE w.WorkGroup = @MachineID AND w.WIPEvent = 1
            //                ORDER BY w.UpdateTime DESC";
            string SqlStr = @$"SELECT TOP(1) a.OrderID,a.OPID,a.QtyGood,a.QtyBad,a.DeviceID,c.WIPEvent,a.CreateTime,a.StaffID,
                              Progress = cast( (cast(a.QtyGood as float) + cast(a.QtyBad as float)) / cast(c.OrderQTY as float) * 100 as int), RemainingCount = (c.OrderQTY - c.QtyTol)
                              FROM {_ConnectStr.APSDB}.[dbo].WIPLog as a
                              left join {_ConnectStr.APSDB}.[dbo].Assignment as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                              left join {_ConnectStr.APSDB}.[dbo].WIP as c on b.OrderID=c.OrderID and b.OPID=c.OPID
                              where a.DeviceID  =@MachineID
                              order by a.CreateTime desc";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new LastUpdateOP
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString().Trim(),
                                    WIPEvent = SqlData["WIPEvent"].ToString().Trim(),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString())
                                });
                            }
                        }
                        else
                        {
                            result = LastUpdateOPStatus(MachineID);
                        }
                    }
                }
            }
            return result;
        }

        private List<LastUpdateOP> LastUpdateOPStatus(string MachineID)
        {
            var result = new List<LastUpdateOP>();
            string SqlStr = @$"SELECT top(1) w.OrderID,w.OPID, w.WIPEvent, w.UpdateTime
                            , Progress = cast( (cast(w.QtyGood as float) + cast(w.QtyBad as float)) / cast(w.OrderQTY as float) * 100 as int)
                            , RemainingCount = (w.OrderQTY - w.QtyTol)
                            ,a.Range
                            FROM {_ConnectStr.APSDB}.[dbo].WIP as w
                            inner join {_ConnectStr.APSDB}.[dbo].Assignment as a on w.OrderID=a.OrderID and w.OPID=a.OPID
                            WHERE w.WorkGroup = @MachineID
                            ORDER BY w.UpdateTime DESC , w.WIPEvent";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@MachineID", SqlDbType.NVarChar).Value = MachineID;
                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                //if (SqlData["Range"].ToString() != "0")
                                result.Add(new LastUpdateOP
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString().Trim(),
                                    WIPEvent = SqlData["WIPEvent"].ToString().Trim(),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim())
                                });
                                //else
                                //    result.Add(new LastUpdateOP
                                //    {
                                //        OrderID = "-",
                                //        OPID = "-",
                                //        WIPEvent = "0",
                                //        Progress = 0
                                //    });

                            }
                        }
                    }
                }
            }

            if (result.Count == 0)
            {
                result.Add(new LastUpdateOP
                {
                    OrderID = "-",
                    OPID = "-",
                    WIPEvent = "0",
                    Progress = 0
                });
            }

            return result;
        }

        private int CanWorkOrderCount(string MachineID)
        {
            int result = 0;
            var SqlStr = @$"SELECT a.*, d.remark,d.GroupName, p.user_name, wip.WIPEvent, wip.OrderQTY, wip.QtyTol, wip.QtyGood, wip.QtyBad ,q.IsQC, k.user_name as QCman,
                        Progress = cast( (cast(wip.QtyGood as float) + cast(wip.QtyBad as float)) / cast(a.OrderQTY as float) * 100 as int), RemainingCount = (a.OrderQTY - wip.QtyTol)
                        FROM {_ConnectStr.APSDB}.[dbo].[Assignment] as a
                        LEFT JOIN {_ConnectStr.APSDB}.[dbo].[Device] as d ON a.[WorkGroup] = d.remark
                        LEFT JOIN {_ConnectStr.AccuntDB}.[dbo].[User] as p ON a.Operator = p.[user_id]
                        LEFT JOIN {_ConnectStr.APSDB}.[dbo].[WIP] as wip ON (wip.OrderID=a.OrderID and wip.OPID = a.OPID)
                        LEFT JOIN {_ConnectStr.APSDB}.[dbo].[QCAssignment] as q ON (q.WorkOrderID = a.OrderID and q.OPID = a.OPID)
                        left join {_ConnectStr.AccuntDB}.[dbo].[User] as k on q.QCman=k.[user_id]
                        WHERE d.remark= @MachineName and wip.WIPEvent = 0";
            var temp = new List<string>();
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@MachineName"), SqlDbType.NVarChar).Value = MachineID;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                if (canWork(SqlData["OrderId"].ToString().Trim(), Convert.ToInt32(SqlData["OPID"].ToString().Trim())))
                                {
                                    temp.Add(SqlData["OrderId"].ToString().Trim());
                                    //temp.Add(new OPList
                                    //{
                                    //    OrderId = checkNoword(SqlData["OrderId"].ToString().Trim()),
                                    //    OPId = checkNoword(SqlData["OPID"].ToString().Trim()),
                                    //    ProductId = checkNoword(SqlData["MAKTX"].ToString().Trim()),
                                    //    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    //    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    //    OPStatus = checkNoword(SqlData["WIPEvent"].ToString().Trim()),
                                    //    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["AssignDate"]).ToString(_dateFormat),
                                    //    DeleyDays = deleyday(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString()),
                                    //    Deviec = checkNoword(SqlData["remark"].ToString().Trim()),
                                    //    DeviceGroup = checkNoword(SqlData["GroupName"].ToString().Trim()),
                                    //    OperatorName = checkNoword(SqlData["user_name"].ToString().Trim()),
                                    //    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim()),
                                    //    RequireNum = checkNoword(SqlData["OrderQTY"].ToString().Trim()),
                                    //    CompleteNum = checkNoword(SqlData["QtyGood"].ToString().Trim()),
                                    //    DefectiveNum = checkNoword(SqlData["QtyBad"].ToString().Trim()),
                                    //    IsQC = QCStatus(IsQCDone(SqlData["OrderId"].ToString().Trim(), SqlData["OPID"].ToString().Trim()).ToString()),
                                    //    QCman = checkNoword(SqlData["QCman"].ToString().Trim())
                                    //});
                                }
                            }
                        }
                    }
                }
            }
            result = temp.Count();
            return result;
        }

        /// <summary>
        /// 取得機台所有工單製程
        /// </summary>
        /// <param name="MachineName"></param>
        /// <param name="Event"></param>
        /// <returns></returns>
        [HttpPost("OrderList")]
        public List<OPList> OrderList(string MachineName, string Event = "")
        {
            //var UserPower = checkLoginPermission();//檢查登入權限

            UserData userdata = UserInfo();
            var result = new List<OPList>();

            if (!ValidateMachine(MachineName))
            {
                return result;
            }

            var SqlStr = "";
            var cmd1 = "";
            var cmd2 = "";
            if (userdata.GroupId.ToString().Trim() == "2")//給生館跳製程權限
            {
                if (Event != "3")
                {
                    //cmd1 = $"where b.WIPEvent!=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                }
                else
                {
                    //cmd1 = $"where b.WIPEvent=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                    cmd1 = $"where b.WIPEvent=3";
                    cmd2 = $"where bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                }
            }
            else
            {
                if (Event != "3")
                {
                    //cmd1 = $"where b.WIPEvent!=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName and aa.remark=@MachineName";

                }
                else
                {
                    //cmd1 = $"where b.WIPEvent=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName and aa.remark=@MachineName";
                }
            }
            SqlStr = @$";WITH cte AS
                        (
                           SELECT
                                 ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY [Range]) AS rn,*
                           FROM 
                           (select 
	                        ord.OrderID as ERPOrderID, a.OrderID,a.OPID,a.[Range],a.OPLTXA1,a.MAKTX,c.[Name],a.StartTime,a.EndTime,a.AssignDate,a.WorkGroup as remark,d.GroupName
	                        ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=a.Operator) as [user_name]
	                        ,Progress = cast( (cast(b.QtyGood as float) + cast(b.QtyBad as float)) / cast(a.OrderQTY as float) * 100 as int)
	                        ,RemainingCount = (a.OrderQTY - b.QtyTol)
	                        ,a.OrderQTY,b.QtyGood,b.QtyBad,b.WIPEvent,b.StartTime as wipStartTime,b.EndTime as wipEndTime
                            ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=f.QCman) as QCman
                            ,g.CustomerInfo
							,d.ID as DeviceID
	                        from {_ConnectStr.APSDB}.[dbo].Assignment as a
	                        left join {_ConnectStr.APSDB}.[dbo].WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                            left join {_ConnectStr.APSDB}.[dbo].[OrderOverview] as ord on a.ERPOrderID=ord.OrderID
	                        left join {_ConnectStr.MRPDB}.dbo.Part as c on a.MAKTX=c.Number
	                        left join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                            left join {_ConnectStr.APSDB}.[dbo].QCAssignment as f on a.OrderID=f.WorkOrderID and a.OPID=f.OPID
                            left join {_ConnectStr.APSDB}.[dbo].OrderOverview as g on a.ERPOrderID=g.OrderID
	                        {cmd1} and a.WorkGroup is not null) as a
                        )
                        SELECT TOP(5)*
                        FROM cte as aa
						left join {_ConnectStr.AccuntDB}.dbo.GroupDevice as bb on aa.DeviceID=bb.DeviceId
						{cmd2}
                        order by OrderID";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@MachineName"), SqlDbType.NVarChar).Value = MachineName;
                    comm.Parameters.Add(("@GroupName"), SqlDbType.NVarChar).Value = userdata.GroupId.ToString().Trim();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                string delaydaydisplay = "";
                                int dd = deleyday(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString());
                                if (dd < 0)
                                    delaydaydisplay = "延遲 " + Math.Abs(dd) + " 天";
                                else
                                    delaydaydisplay = "還有 " + Math.Abs(dd) + " 天";

                                result.Add(new OPList
                                {
                                    ERPorderID = string.IsNullOrEmpty(SqlData["ERPOrderID"].ToString().Trim()) ? "N/A" : SqlData["ERPOrderID"].ToString().Trim(),
                                    OrderId = checkNoword(SqlData["OrderId"].ToString().Trim()),
                                    OPId = checkNoword(SqlData["OPID"].ToString().Trim()),
                                    Range = string.IsNullOrEmpty(SqlData["Range"].ToString()) ? new Random().Next(999, 1024) : int.Parse(SqlData["Range"].ToString()),
                                    OPName = checkNoword(SqlData["OPLTXA1"].ToString().Trim()),
                                    ProductId = checkNoword(SqlData["MAKTX"].ToString().Trim()),
                                    //ProductName = "【" + (string.IsNullOrEmpty(SqlData["CustomerInfo"].ToString().Trim()) ? "NA/NA/NA" : SqlData["CustomerInfo"].ToString().Trim()).Split("/")[1] + "】" + (checkNoword(SqlData["Name"].ToString().Trim())),
                                    ProductName = (checkNoword(SqlData["Name"].ToString().Trim())),
                                    //暫時先改給【預計開始時間】(目前前端讀取)
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    //暫時先改給【預計結束時間】(目前前端讀取)
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    OPStatus = checkNoword(SqlData["WIPEvent"].ToString().Trim()),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["AssignDate"]).ToString(_dateFormat),
                                    DeleyDays = delaydaydisplay,
                                    Days = dd,
                                    IsDeley = Isdeley(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString()),
                                    Deviec = checkNoword(SqlData["remark"].ToString().Trim()),
                                    DeviceGroup = checkNoword(SqlData["GroupName"].ToString().Trim()),
                                    OperatorName = checkNoword(SqlData["user_name"].ToString().Trim()),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim()),
                                    RequireNum = checkNoword(SqlData["OrderQTY"].ToString().Trim()),
                                    CompleteNum = checkNoword(SqlData["QtyGood"].ToString().Trim()),
                                    DefectiveNum = checkNoword(SqlData["QtyBad"].ToString().Trim()),
                                    //IsQC = "N/A",//2023.06.17 Ryan修改 皆無須檢驗
                                    IsQC = QCStatus(IsQCDone(SqlData["OrderId"].ToString().Trim(), SqlData["OPID"].ToString().Trim()).ToString()),
                                    QCman = checkNoword(SqlData["QCman"].ToString().Trim()),
                                    //實際開工時間
                                    wipStartTime = string.IsNullOrEmpty(SqlData["wipStartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["wipStartTime"]).ToString(_timeFormat),
                                    //實際完工時間
                                    wipEndTime = string.IsNullOrEmpty(SqlData["wipEndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["wipEndTime"]).ToString(_timeFormat),
                                });
                            }
                        }
                    }
                }
            }

            if(result.Count>0)
            {
                foreach (var item in result)
                {
                    //判斷機台是否有委外廠商名稱
                    bool hasBrackets = item.Deviec.Contains("(") && item.Deviec.Contains(")");
                    if (hasBrackets)
                    {
                        int start = item.Deviec.IndexOf("(");
                        int end = item.Deviec.IndexOf(")");
                        if (start >= 0 && end > start)
                        {
                            string factory = item.Deviec.Substring(start + 1, end - start - 1);
                            item.ProductName += $"({factory})";
                        }

                    }
                }
                return result.OrderBy(x => x.Days)
                    .ThenBy(x => x.AssignDate)
                    .ThenBy(x => x.OrderId)
                    .ThenBy(x => x.Range)
                    .ToList();
            }
            else
            {
                return result;
            }
            
        }

        /// <summary>
        /// 檢視全部機台工單
        /// </summary>
        /// <returns></returns>
        [HttpPost("OrderListByALL")]
        public List<OPList> OrderListByALL(string Event)
        {
            //var UserPower = checkLoginPermission();//檢查登入權限

            UserData userdata = UserInfo();
            var result = new List<OPList>();
            var SqlStr = "";

            var cmd1 = "";
            var cmd2 = "";
            if (userdata.GroupId.ToString().Trim() == "2")//給生館跳製程權限
            {
                if (Event != "3")
                {
                    //cmd1 = $"where b.WIPEvent!=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName";

                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName";


                }
                else
                {
                    //cmd1 = $"where b.WIPEvent=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName";

                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName";
                }
            }
            else
            {
                if (Event != "3")
                {
                    //cmd1 = $"where b.WIPEvent!=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName";

                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName";
                }
                else
                {
                    //cmd1 = $"where b.WIPEvent=3";
                    //cmd2 = $"where WIPEvent={Event} and bb.GroupSeq=@GroupName";

                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where bb.GroupSeq=@GroupName";
                }
            }
            SqlStr = @$"WITH cte AS
                        (
                           SELECT
                                 ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY [Range]) AS rn,*
                           FROM 
                           (select 
	                        ord.OrderID as ERPOrderID,a.OrderID,a.OPID,a.[Range],a.OPLTXA1,a.MAKTX,c.[Name],a.StartTime,a.EndTime,a.AssignDate,a.WorkGroup as remark,d.GroupName
	                        ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=a.Operator) as [user_name]
	                        ,Progress = cast( (cast(b.QtyGood as float) + cast(b.QtyBad as float)) / cast(a.OrderQTY as float) * 100 as int)
	                        ,RemainingCount = (a.OrderQTY - b.QtyTol)
	                        ,a.OrderQTY,b.QtyGood,b.QtyBad,b.WIPEvent,b.StartTime as wipStartTime,b.EndTime as wipEndTime
                            ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=f.QCman) as QCman
                            ,g.CustomerInfo
							,d.ID as DeviceID
                            ,(SELECT COUNT(*) FROM {_ConnectStr.MRPDB}.dbo.QCrule as qcr WHERE qcr.id=a.OPID) AS QC_count
                            ,(select COUNT(*) FROM {_ConnectStr.APSDB}.dbo.QCPointValue as qcpv where qcpv.WorkOrderID=a.OrderID and qcpv.OPID=a.OPID and qcpv.MAKTX=a.MAKTX) AS QCv_count
	                        from {_ConnectStr.APSDB}.[dbo].Assignment as a
	                        inner join {_ConnectStr.APSDB}.[dbo].WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                            left join {_ConnectStr.APSDB}.[dbo].[OrderOverview] as ord on a.ERPOrderID=ord.OrderID
	                        left join {_ConnectStr.MRPDB}.dbo.Part as c on a.MAKTX=c.Number
	                        left join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                            left join {_ConnectStr.APSDB}.[dbo].QCAssignment as f on a.OrderID=f.WorkOrderID and a.OPID=f.OPID
                            left join {_ConnectStr.APSDB}.[dbo].OrderOverview as g on a.ERPOrderID=g.OrderID
	                        {cmd1}  and a.WorkGroup is not null) as a
                        )
                        SELECT  TOP(5)*
                        FROM cte as aa
						left join {_ConnectStr.AccuntDB}.dbo.GroupDevice as bb on aa.DeviceID=bb.DeviceId
						{cmd2}
                        order by OrderID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@GroupName"), SqlDbType.NVarChar).Value = userdata.GroupId.ToString().Trim();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                string delaydaydisplay = "";
                                int dd = deleyday(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString());
                                if (dd < 0)
                                    delaydaydisplay = "延遲 " + Math.Abs(dd) + " 天";
                                else
                                    delaydaydisplay = "還有 " + Math.Abs(dd) + " 天";

                                result.Add(new OPList
                                {
                                    ERPorderID = string.IsNullOrEmpty(SqlData["ERPOrderID"].ToString().Trim()) ? "N/A" : SqlData["ERPOrderID"].ToString().Trim(),
                                    OrderId = string.IsNullOrEmpty(SqlData["OrderId"].ToString().Trim()) ? "N/A" : SqlData["OrderId"].ToString().Trim(),
                                    OPId = string.IsNullOrEmpty(SqlData["OPID"].ToString().Trim()) ? "N/A" : SqlData["OPID"].ToString().Trim(),
                                    Range = string.IsNullOrEmpty(SqlData["Range"].ToString().Trim()) ? new Random().Next(999, 1024) : int.Parse(SqlData["Range"].ToString().Trim()),
                                    OPName = string.IsNullOrEmpty(SqlData["OPLTXA1"].ToString().Trim()) ? "N/A" : SqlData["OPLTXA1"].ToString().Trim(),
                                    ProductId = string.IsNullOrEmpty(SqlData["MAKTX"].ToString().Trim()) ? "N/A" : SqlData["MAKTX"].ToString().Trim(),
                                    //ProductName = "【" + (string.IsNullOrEmpty(SqlData["CustomerInfo"].ToString().Trim()) ? "NA/NA/NA" : SqlData["CustomerInfo"].ToString().Trim()).Split("/")[1] + "】" + (string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? "N/A" : SqlData["Name"].ToString().Trim()),
                                    ProductName = (string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? "N/A" : SqlData["Name"].ToString().Trim()),
                                    //暫時先改給【預計開始時間】(目前前端讀取)
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),//checkNoword(SqlData["StartTime"].ToString().Trim()),
                                    //暫時先改給【預計結束時間】(目前前端讀取)
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),//checkNoword(SqlData["EndTime"].ToString().Trim()),
                                    OPStatus = string.IsNullOrEmpty(SqlData["WIPEvent"].ToString().Trim()) ? "N/A" : SqlData["WIPEvent"].ToString().Trim(),
                                    //OPStatus = "1",
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["AssignDate"]).ToString(_dateFormat),
                                    DeleyDays = delaydaydisplay,
                                    Days = dd,
                                    IsDeley = Isdeley(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString()),
                                    Deviec = string.IsNullOrEmpty(SqlData["remark"].ToString().Trim()) ? "N/A" : SqlData["remark"].ToString().Trim(),
                                    DeviceGroup = string.IsNullOrEmpty(SqlData["GroupName"].ToString().Trim()) ? "N/A" : SqlData["GroupName"].ToString().Trim(),
                                    OperatorName = string.IsNullOrEmpty(SqlData["user_name"].ToString().Trim()) ? "N/A" : SqlData["user_name"].ToString().Trim(),
                                    //實際開工時間
                                    wipStartTime = string.IsNullOrEmpty(SqlData["wipStartTime"].ToString().Trim()) ? "N/A" : Convert.ToDateTime(SqlData["wipStartTime"]).ToString(_timeFormat).Trim(),
                                    //實際完工時間
                                    wipEndTime = string.IsNullOrEmpty(SqlData["wipEndTime"].ToString().Trim()) ? "N/A" : Convert.ToDateTime(SqlData["wipEndTime"]).ToString(_timeFormat).Trim(),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim()),
                                    RequireNum = string.IsNullOrEmpty(SqlData["OrderQTY"].ToString().Trim()) ? "N/A" : SqlData["OrderQTY"].ToString().Trim(),
                                    CompleteNum = string.IsNullOrEmpty(SqlData["QtyGood"].ToString().Trim()) ? "N/A" : SqlData["QtyGood"].ToString().Trim(),
                                    DefectiveNum = string.IsNullOrEmpty(SqlData["QtyBad"].ToString().Trim()) ? "N/A" : SqlData["QtyBad"].ToString().Trim(),
                                    //IsQC = "N/A",//2023.06.17 Ryan修改 皆無須檢驗
                                    IsQC = QCStatus(IsQCDone_2(SqlData["QC_count"].ToString(), SqlData["QCv_count"].ToString())),
                                    QCman = string.IsNullOrEmpty(SqlData["QCMan"].ToString()) ? "N/A" : SqlData["QCMan"].ToString()
                                });
                            }
                        }
                    }
                }
            }

            if(result.Count>0)
            {
                foreach (var item in result)
                {
                    //判斷機台是否有委外廠商名稱
                    bool hasBrackets = item.Deviec.Contains("(") && item.Deviec.Contains(")");
                    if (hasBrackets)
                    {
                        int start = item.Deviec.IndexOf("(");
                        int end = item.Deviec.IndexOf(")");
                        if (start >= 0 && end > start)
                        {
                            string factory = item.Deviec.Substring(start + 1, end - start - 1);
                            item.ProductName += $"({factory})";
                        }

                    }
                }

                return result.OrderBy(x => x.Days)
                .ThenBy(x => x.AssignDate)
                .ThenBy(x => x.OrderId)
                .ThenBy(x => x.Range)
                .ToList();
            }
            else
            {
                return result;
            }

            
            
            
        }

        /// <summary>
        /// 搜尋功能
        /// </summary>
        /// <param name="Event"></param>
        /// <param name="OrderId"></param>
        /// <returns></returns>
        [HttpPost("OrderListByALLSearch")]
        public List<OPList> OrderListByALLSearch(string Event, string OrderId)
        {
            UserData userdata = UserInfo();
            var result = new List<OPList>();
            var SqlStr = "";
            var cmd1 = "";
            var cmd2 = "";
            if (userdata.GroupId.ToString().Trim() == "2")//給生館跳製程權限
            {
                if (Event != "3")
                {
                    //cmd1 = $"where b.WIPEvent!=3";
                    //cmd2 = $"where WIPEvent={Event} and aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";

                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";
                }
                else
                {
                    //cmd1 = $"where b.WIPEvent=3";
                    //cmd2 = $"where WIPEvent={Event} and aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";

                    cmd1 = $"where b.WIPEvent={Event}";
                    cmd2 = $"where aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";
                }
            }
            else
            {
                if (Event != "3")
                {
                    cmd1 = $"where b.WIPEvent!=3";
                    cmd2 = $"where WIPEvent={Event} and aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";
                    //cmd2 = $"where rn=1 and WIPEvent={Event} and aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";
                }
                else
                {
                    cmd1 = $"where b.WIPEvent=3";
                    cmd2 = $"where WIPEvent={Event} and aa.OrderID='{OrderId}' and bb.GroupSeq=@GroupName";
                }
            }

            SqlStr = @$";WITH cte AS
                        (
                            SELECT
                                ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY [Range]) AS rn,
                                *
                            FROM 
                            (
                                SELECT 
                                    ord.OrderID as ERPOrderID,
                                    a.OrderID,
                                    a.OPID,
                                    a.[Range],
                                    a.OPLTXA1,
                                    a.MAKTX,
                                    c.[Name],
                                    a.StartTime,
                                    a.EndTime,
                                    a.AssignDate,
                                    a.WorkGroup as remark,
                                    d.GroupName,
                                    (SELECT TOP(1) [user_name] FROM {_ConnectStr.AccuntDB}.dbo.[User] WHERE [user_id]=a.Operator) as [user_name],
                                    Progress = CAST((CAST(b.QtyGood as float) + CAST(b.QtyBad as float)) / CAST(a.OrderQTY as float) * 100 as int),
                                    RemainingCount = (a.OrderQTY - b.QtyTol),
                                    a.OrderQTY,
                                    b.QtyGood,
                                    b.QtyBad,
                                    b.WIPEvent,
                                    b.StartTime as wipStartTime,
                                    b.EndTime as wipEndTime,
                                    (SELECT TOP(1) [user_name] FROM {_ConnectStr.AccuntDB}.dbo.[User] WHERE [user_id]=f.QCman) as QCman,
                                    g.CustomerInfo,
                                    d.ID as DeviceID,
                                    (SELECT COUNT(*) FROM {_ConnectStr.MRPDB}.dbo.QCrule as qcr WHERE qcr.id=a.OPID) AS QC_count,
                                    (SELECT COUNT(*) FROM {_ConnectStr.APSDB}.dbo.QCPointValue as qcpv WHERE qcpv.WorkOrderID=a.OrderID and qcpv.OPID=a.OPID and qcpv.MAKTX=a.MAKTX) AS QCv_count
                                FROM {_ConnectStr.APSDB}.[dbo].Assignment as a
                                INNER JOIN {_ConnectStr.APSDB}.[dbo].WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                                LEFT JOIN {_ConnectStr.APSDB}.[dbo].[OrderOverview] as ord on a.ERPOrderID=ord.OrderID
                                LEFT JOIN {_ConnectStr.MRPDB}.[dbo].[Part] as c on a.MAKTX=c.Number
                                LEFT JOIN {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                                LEFT JOIN {_ConnectStr.APSDB}.[dbo].QCAssignment as f on a.OrderID=f.WorkOrderID and a.OPID=f.OPID
                                LEFT JOIN {_ConnectStr.APSDB}.[dbo].OrderOverview as g on a.ERPOrderID=g.OrderID
                                {cmd1} and a.Workgroup is not null
                            ) as a
                        )
                        SELECT *
                        FROM cte as aa
                        LEFT JOIN {_ConnectStr.AccuntDB}.dbo.GroupDevice as bb on aa.DeviceID=bb.DeviceId
                        {cmd2}
                        ORDER BY OrderID, [Range]";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@GroupName"), SqlDbType.NVarChar).Value = userdata.GroupId.ToString().Trim();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                string delaydaydisplay = "";
                                int dd = deleyday(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString());
                                if (dd < 0)
                                    delaydaydisplay = "延遲 " + Math.Abs(dd) + " 天";
                                else
                                    delaydaydisplay = "還有 " + Math.Abs(dd) + " 天";

                                result.Add(new OPList
                                {
                                    OrderId = string.IsNullOrEmpty(SqlData["OrderId"].ToString().Trim()) ? "N/A" : SqlData["OrderId"].ToString().Trim(),
                                    OPId = string.IsNullOrEmpty(SqlData["OPID"].ToString().Trim()) ? "N/A" : SqlData["OPID"].ToString().Trim(),
                                    Range = string.IsNullOrEmpty(SqlData["Range"].ToString().Trim()) ? new Random().Next(999, 1024) : int.Parse(SqlData["Range"].ToString().Trim()),
                                    OPName = string.IsNullOrEmpty(SqlData["OPLTXA1"].ToString().Trim()) ? "N/A" : SqlData["OPLTXA1"].ToString().Trim(),
                                    ProductId = string.IsNullOrEmpty(SqlData["MAKTX"].ToString().Trim()) ? "N/A" : SqlData["MAKTX"].ToString().Trim(),
                                    ProductName = "【" + (string.IsNullOrEmpty(SqlData["CustomerInfo"].ToString().Trim()) ? "NA/NA/NA" : SqlData["CustomerInfo"].ToString().Trim()).Split("/")[1] + "】" + (string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? "N/A" : SqlData["Name"].ToString().Trim()),
                                    //暫時先改給【預計開始時間】
                                    StartTime = string.IsNullOrEmpty(SqlData["wipStartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["wipStartTime"]).ToString(_timeFormat),//checkNoword(SqlData["StartTime"].ToString().Trim()),
                                    //暫時先改給【預計結束時間】
                                    EndTime = string.IsNullOrEmpty(SqlData["wipEndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["wipEndTime"]).ToString(_timeFormat),//checkNoword(SqlData["EndTime"].ToString().Trim()),
                                    OPStatus = string.IsNullOrEmpty(SqlData["WIPEvent"].ToString().Trim()) ? "N/A" : SqlData["WIPEvent"].ToString().Trim(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["AssignDate"]).ToString(_dateFormat),
                                    DeleyDays = delaydaydisplay,
                                    Days = dd,
                                    IsDeley = Isdeley(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString()),
                                    Deviec = string.IsNullOrEmpty(SqlData["remark"].ToString().Trim()) ? "N/A" : SqlData["remark"].ToString().Trim(),
                                    DeviceGroup = string.IsNullOrEmpty(SqlData["GroupName"].ToString().Trim()) ? "N/A" : SqlData["GroupName"].ToString().Trim(),
                                    OperatorName = string.IsNullOrEmpty(SqlData["user_name"].ToString().Trim()) ? "N/A" : SqlData["user_name"].ToString().Trim(),
                                    wipStartTime = string.IsNullOrEmpty(SqlData["wipStartTime"].ToString().Trim()) ? "N/A" : Convert.ToDateTime(SqlData["wipStartTime"]).ToString(_timeFormat).Trim(),
                                    wipEndTime = string.IsNullOrEmpty(SqlData["wipEndTime"].ToString().Trim()) ? "N/A" : Convert.ToDateTime(SqlData["wipEndTime"]).ToString(_timeFormat).Trim(),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim()),
                                    RequireNum = string.IsNullOrEmpty(SqlData["OrderQTY"].ToString().Trim()) ? "N/A" : SqlData["OrderQTY"].ToString().Trim(),
                                    CompleteNum = string.IsNullOrEmpty(SqlData["QtyGood"].ToString().Trim()) ? "N/A" : SqlData["QtyGood"].ToString().Trim(),
                                    DefectiveNum = string.IsNullOrEmpty(SqlData["QtyBad"].ToString().Trim()) ? "N/A" : SqlData["QtyBad"].ToString().Trim(),
                                    IsQC = QCStatus(IsQCDone_2(SqlData["QC_count"].ToString(), SqlData["QCv_count"].ToString())),//QCStatus(IsQCDone(SqlData["QC_count"].ToString().Trim(), SqlData["OPID"].ToString().Trim()).ToString()),
                                    QCman = string.IsNullOrEmpty(SqlData["QCMan"].ToString()) ? "N/A" : SqlData["QCMan"].ToString()
                                });
                            }
                        }
                    }
                }
            }

            return result.OrderBy(x => x.Days)
                .ThenBy(x => x.AssignDate)
                .ThenBy(x => x.OrderId)
                .ThenBy(x => x.Range)
                .ToList();
        }

        private bool checklastopdone(OPList item)
        {
            bool result = false;
            string SqlStr = @"select * from Assignment as a 
                            inner join WIP as b on a.OrderID =b.OrderID and a.OPID=b.OPID
                            where a.OrderID=@OrderID and a.Range=@Range and b.WIPEvent!=0
                            order by a.Range";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = item.OrderId;
                    comm.Parameters.Add(("@Range"), SqlDbType.Int).Value = item.Range - 1;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            //如果查有資料代表前製成已經開工，報工系統可以顯示
                            result = true;
                        }
                    }
                }
            }
            return result;
        }

        //未完成
        private void getallorderlist(string Event)
        {
            //變數
            var result = new List<OPList>();
            string SqlStr = "";
            UserData userdata = UserInfo();

            //取得WIP列表
            SqlStr = @"select * from WIP where WIPEvent=0 order by SeriesID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    //comm.Parameters.Add(("@GroupName"), SqlDbType.NVarChar).Value = userdata.GroupId.ToString().Trim();
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new OPList
                                {
                                    OrderId = string.IsNullOrEmpty(SqlData["OrderId"].ToString().Trim()) ? "N/A" : SqlData["OrderId"].ToString().Trim(),
                                    OPId = string.IsNullOrEmpty(SqlData["OPID"].ToString().Trim()) ? "N/A" : SqlData["OPID"].ToString().Trim(),
                                    OPName = string.IsNullOrEmpty(SqlData["OPLTXA1"].ToString().Trim()) ? "N/A" : SqlData["OPLTXA1"].ToString().Trim(),
                                    //ProductId = string.IsNullOrEmpty(SqlData["MAKTX"].ToString().Trim()) ? "N/A" : SqlData["MAKTX"].ToString().Trim(),
                                    //ProductName = string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? "N/A" : SqlData["Name"].ToString().Trim(),
                                    //StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),//checkNoword(SqlData["StartTime"].ToString().Trim()),
                                    //EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),//checkNoword(SqlData["EndTime"].ToString().Trim()),
                                    //OPStatus = string.IsNullOrEmpty(SqlData["WIPEvent"].ToString().Trim()) ? "N/A" : SqlData["WIPEvent"].ToString().Trim(),
                                    //AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["AssignDate"]).ToString(_dateFormat),
                                    //DeleyDays = deleyday(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString()),
                                    //Deviec = string.IsNullOrEmpty(SqlData["remark"].ToString().Trim()) ? "N/A" : SqlData["remark"].ToString().Trim(),
                                    //DeviceGroup = string.IsNullOrEmpty(SqlData["GroupName"].ToString().Trim()) ? "N/A" : SqlData["GroupName"].ToString().Trim(),
                                    //OperatorName = string.IsNullOrEmpty(SqlData["user_name"].ToString().Trim()) ? "N/A" : SqlData["user_name"].ToString().Trim(),
                                    //Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim()),
                                    //RequireNum = string.IsNullOrEmpty(SqlData["OrderQTY"].ToString().Trim()) ? "N/A" : SqlData["OrderQTY"].ToString().Trim(),
                                    //CompleteNum = string.IsNullOrEmpty(SqlData["QtyGood"].ToString().Trim()) ? "N/A" : SqlData["QtyGood"].ToString().Trim(),
                                    //DefectiveNum = string.IsNullOrEmpty(SqlData["QtyBad"].ToString().Trim()) ? "N/A" : SqlData["QtyBad"].ToString().Trim(),
                                    //IsQC = QCStatus(IsQCDone(SqlData["OrderId"].ToString().Trim(), SqlData["OPID"].ToString().Trim()).ToString()),
                                    //QCman = string.IsNullOrEmpty(SqlData["QCMan"].ToString()) ? "N/A" : SqlData["QCMan"].ToString()
                                });
                            }
                        }
                    }
                }
            }

            //取得我個人負責機台
            var canUseDevice = new List<string>();
            SqlStr = $@"select * from {_ConnectStr.AccuntDB}.dbo.GroupDevice as a
                        inner join Device as b on a.DeviceId=b.ID
                        where a.GroupSeq=6";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    //comm.Parameters.Add(("@GroupName"), SqlDbType.NVarChar).Value = userdata.GroupId.ToString().Trim();
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                canUseDevice.Add(SqlData["remark"].ToString());
                            }
                        }
                    }
                }
            }

            //取得供單種類
            var orderids = result.Select(x => x.OrderId).Distinct().ToList();
            foreach (var i in orderids)
            {
                //result.Where(x=>x.OrderId==i)..First(x=>x.OPStatus=="0")
                int currentOPindex = result.IndexOf(result.Where(x => x.OrderId == i).First(x => x.OPStatus == "0"));

            }

        }

        private int deleyday(string DeadLine_t, string WillDone_t)
        {
            string result = "";
            int res = 0;
            DateTime dtDate;
            if (_ConnectStr.Customer == 0)
            {
                if (DateTime.TryParse(DeadLine_t, out dtDate))
                {
                    DateTime start = Convert.ToDateTime(DeadLine_t);

                    DateTime end = Convert.ToDateTime(WillDone_t);

                    TimeSpan ts = end.Subtract(start); //兩時間天數相減

                    if (ts.Days < 0)
                    {
                        result = "0 天";
                    }
                    else
                    {
                        result = ts.Days.ToString() + " 天"; //相距天數
                    }
                }
                else
                {
                    result = "N/A";
                }
            }
            else
            {
                if (DateTime.TryParse(DeadLine_t, out dtDate))
                {
                    DateTime start = Convert.ToDateTime(DeadLine_t);

                    DateTime end = DateTime.Now;

                    TimeSpan ts = start.Subtract(end); //兩時間天數相減

                    res = ts.Days;
                }
                else
                {
                    res = 0;
                }
            }
            return res;
        }

        private bool Isdeley(string DeadLine_t, string WillDone_t)
        {
            bool result = false;
            DateTime dtDate;
            if (DateTime.TryParse(DeadLine_t, out dtDate))
            {
                DateTime start = Convert.ToDateTime(DeadLine_t);

                DateTime end = DateTime.Now;

                TimeSpan ts = end.Subtract(start); //兩時間天數相減

                if (ts.Days < 0)
                {
                    result = false;
                }
                else
                {
                    result = true; //相距天數
                }
            }
            else
            {
                result = false;
            }
            return result;
        }

        private string QCStatus(string isqc)
        {
            if (isqc == "" || isqc == "N/A")
            {
                return "N/A";//無須檢驗
            }
            else
            {
                return isqc;//False需要檢驗、Ture檢驗完成
            }
        }

        private UserData UserInfo()
        {
            UserData user = new UserData();

            string sqlStr = @$"SELECT aa.user_id,aa.user_account,aa.user_name,aa.usergroup_id,bb.usergroup_name,bb.DeviceGroup FROM (
                                SELECT user_id, user_account, user_name, usergroup_id
                                FROM {_ConnectStr.AccuntDB}.[dbo].[User])as aa
                                LEFT JOIN {_ConnectStr.AccuntDB}.[dbo].[Units] as bb
                                ON aa.usergroup_id=bb.usergroup_id
                                WHERE aa.user_account=@account";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@account", User.Identity.Name);

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            sqldata.Read();
                            user.User_Id = Convert.ToInt64(sqldata["user_id"]);
                            user.EmpolyeeAccount = sqldata["user_account"].ToString().Trim();
                            user.EmpolyeeName = sqldata["user_name"].ToString().Trim();
                            user.GroupId = Convert.ToInt64(sqldata["usergroup_id"]);
                            user.GroupName = sqldata["usergroup_name"].ToString().Trim();
                            user.DeviceGroupId = sqldata["DeviceGroup"].ToString().Trim();
                        }
                    }
                }
                sqlStr = @$"SELECT * FROM {_ConnectStr.APSDB}.[dbo].[Device] as a
                            left join　{_ConnectStr.AccuntDB}.[dbo].[Units] as b on a.GroupName = b.DeviceGroup
                            left join {_ConnectStr.AccuntDB}.[dbo].[User] as c on b.usergroup_id=c.usergroup_id
                            where c.user_id = @userid";
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@userid", user.User_Id);

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            List<GroupDevice> GroupDevices = new List<GroupDevice>();
                            while (sqldata.Read())
                            {
                                GroupDevices.Add(new GroupDevice
                                {
                                    ID = Convert.ToInt64(sqldata["ID"]),
                                    MachineNmae = sqldata["MachineName"].ToString().Replace("\n", string.Empty),
                                    remark = sqldata["remark"].ToString(),
                                    GroupName = sqldata["GroupName"].ToString(),
                                    CommonName = sqldata["CommonName"].ToString()
                                });
                            }
                            user.GroupDevices = GroupDevices;
                        }
                    }
                }

                sqlStr = @$"select a.*, b.*,b.FuncName,  b.DetailSet, a.ViewRight, a.AuditRight, a.ModifyRight, a.DeleteRight , 
                            Used = case when (a.FuncName = b.FuncName) then 'True' else 'False' end from  {_ConnectStr.AccuntDB}.[dbo].Functions b  left join {_ConnectStr.AccuntDB}.[dbo].GroupRights as a on a.GroupSeq =
                            (select TOP(1) gm.GroupSeq FROM {_ConnectStr.AccuntDB}.[dbo].GroupMembers as gm Where gm.MemberSeqNo = @userid)
                            and b.Status = '啟用' and a.FuncName = b.FuncName
                            order by Belong";
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@userid", user.User_Id);

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            List<FunctionAccess> FunctionAccess = new List<FunctionAccess>();
                            while (sqldata.Read())
                            {
                                bool used = Convert.ToBoolean(sqldata["Used"]);
                                if (used)
                                {
                                    FunctionAccess.Add(new FunctionAccess
                                    {
                                        FunctionName = sqldata["FuncName"].ToString(),
                                        ViewRight = sqldata["ViewRight"].ToString(),
                                        ModifyRight = sqldata["ModifyRight"].ToString(),
                                        DeleteRight = sqldata["DeleteRight"].ToString(),
                                        AuditRight = sqldata["AuditRight"].ToString(),
                                        Used = sqldata["Used"].ToString()
                                    });
                                }

                            }
                            user.FunctionAccess = FunctionAccess;
                            return user;
                        }
                    }
                }

                return user;
            }
            //UserData user = new UserData();

            //string sqlStr = @$"SELECT aa.user_id,aa.user_account,aa.user_name,aa.usergroup_id,bb.usergroup_name,bb.DeviceGroup FROM (
            //                    SELECT user_id, user_account, user_name, usergroup_id
            //                    FROM {_ConnectStr.AccuntDB}.[dbo].[User])as aa
            //                    LEFT JOIN  {_ConnectStr.AccuntDB}.[dbo].[Units] as bb
            //                    ON aa.usergroup_id=bb.usergroup_id
            //                    WHERE aa.user_account=@account";
            //using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            //{
            //    if (conn.State != ConnectionState.Open)
            //        conn.Open();
            //    using (SqlCommand comm = new SqlCommand(sqlStr, conn))
            //    {
            //        comm.Parameters.AddWithValue("@account", User.Identity.Name);

            //        using (SqlDataReader sqldata = comm.ExecuteReader())
            //        {
            //            if (sqldata.HasRows)
            //            {
            //                sqldata.Read();
            //                user.User_Id = Convert.ToInt64(sqldata["user_id"]);
            //                user.EmpolyeeAccount = sqldata["user_account"].ToString();
            //                user.EmpolyeeName = sqldata["user_name"].ToString();
            //                user.GroupId = Convert.ToInt64(sqldata["usergroup_id"]);
            //                user.GroupName = sqldata["usergroup_name"].ToString();
            //                user.DeviceGroupId = sqldata["DeviceGroup"].ToString();
            //            }
            //        }
            //    }
            //}
            //return user;
        }

        /// <summary>
        /// 確認使用者是否有訪問該機台的權限
        /// </summary>
        /// <param name="MachineName"></param>
        /// <returns></returns>
        private bool ValidateMachine(string MachineName)
        {
            string UserName = User.Identity.Name;
            //string SqlStr = @$"SELECT d.remark,d.GroupName, d.img as MachineImg,c.[user_name] FROM
            //            {_ConnectStr.AccuntDB}.[dbo].Units as a
            //            LEFT JOIN Device as d ON d.GroupName = a.DeviceGroup
            //            LEFT JOIN {_ConnectStr.AccuntDB}.dbo.[User] as c ON a.usergroup_id = c.usergroup_id
            //            WHERE d.remark = @MachineName and c.user_account = @user_account ";
            string SqlStr = @$"SELECT * 
                            FROM {_ConnectStr.AccuntDB}.dbo.[User] as a
                            inner join {_ConnectStr.AccuntDB}.dbo.GroupDevice as b on a.usergroup_id=b.GroupSeq
                            inner join {_ConnectStr.APSDB}.dbo.Device as c on b.DeviceId=c.ID
                            where a.user_account=@user_account and c.remark=@MachineName";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@MachineName"), SqlDbType.NVarChar).Value = MachineName;
                    comm.Parameters.Add(("@user_account"), SqlDbType.NVarChar).Value = UserName;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 判斷是否有讀取供單權限
        /// </summary>
        /// <param name="OrderID"></param>
        /// <param name="OPID"></param>
        /// <returns></returns>

        private bool ValidOrder(string OrderID, string OPID)
        {
            string UserName = User.Identity.Name;
            //string SqlStr = @$"SELECT d.remark,d.GroupName, d.img as MachineImg,u.[user_name] 
            //                    FROM {_ConnectStr.APSDB}.dbo.Assignment as a
            //                    LEFT JOIN Device as d ON a.WorkGroup=d.remark
            //                    LEFT JOIN {_ConnectStr.AccuntDB}.dbo.Units as m ON d.GroupName=m.DeviceGroup
            //                    LEFT JOIN {_ConnectStr.AccuntDB}.dbo.[User] as u ON m.usergroup_id = u.usergroup_id
            //                    WHERE a.OrderID = @OrderID and a.OPID = @OPID  and u.user_account = @user_account ";
            string SqlStr = @$"SELECT c.remark,c.GroupName, c.img as MachineImg,a.[user_name] 
                            FROM {_ConnectStr.AccuntDB}.dbo.[User] as a
                            inner join {_ConnectStr.AccuntDB}.dbo.[GroupDevice] as b ON a.usergroup_id=b.GroupSeq
                            inner join {_ConnectStr.APSDB}.dbo.Device as c on b.DeviceId=c.id
                            inner join {_ConnectStr.APSDB}.dbo.Assignment as d on c.remark=d.WorkGroup
                            where d.OrderID=@OrderID and d.OPID=@OPID and a.user_account=@user_account";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = OPID;
                    comm.Parameters.Add(("@user_account"), SqlDbType.NVarChar).Value = UserName;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            return true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 判斷是否有QC權限
        /// </summary>
        /// <returns></returns>

        private bool ValidQCOrder()
        {
            bool isExists = false;
            var SqlStr = @$"select a.*, b.*,b.FuncName,  b.DetailSet, a.ViewRight, a.AuditRight, a.ModifyRight, a.DeleteRight , 
                            Used = case when (a.FuncName = b.FuncName) then 'True' else 'False' end from  {_ConnectStr.AccuntDB}.dbo.Functions b  left join {_ConnectStr.AccuntDB}.dbo.GroupRights as a on a.GroupSeq =
                            (select TOP(1) gm.GroupSeq FROM {_ConnectStr.AccuntDB}.dbo.GroupMembers as gm Where gm.MemberSeqNo = (SELECT TOP (1) [user_id] FROM {_ConnectStr.AccuntDB}.[dbo].[User] WHERE [user_account]=@userid))
                            and b.Status = '啟用' and a.FuncName = b.FuncName
                            order by Belong";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@userid"), SqlDbType.NVarChar).Value = User.Identity.Name;

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            List<FunctionAccess> FunctionAccess = new List<FunctionAccess>();
                            while (sqldata.Read())
                            {
                                bool used = Convert.ToBoolean(sqldata["Used"]);
                                if (used)
                                {
                                    FunctionAccess.Add(new FunctionAccess
                                    {
                                        FunctionName = sqldata["FuncName"].ToString(),
                                        ViewRight = sqldata["ViewRight"].ToString(),
                                        ModifyRight = sqldata["ModifyRight"].ToString(),
                                        DeleteRight = sqldata["DeleteRight"].ToString(),
                                        AuditRight = sqldata["AuditRight"].ToString(),
                                        Used = sqldata["Used"].ToString()
                                    });
                                }
                            }
                            isExists = FunctionAccess.Exists(x => x.FunctionName == "機台報工");
                        }
                        else
                        {
                            isExists = false;
                        }
                    }
                }
            }

            return isExists;
        }

        /// <summary>
        /// 製程詳細資料
        /// </summary>
        /// <param name="OrderID"></param>
        /// <param name="OPId"></param>
        /// <returns></returns>
        [HttpPost("OPDetail")]
        public List<OPDetail> OPDetail(string OrderID, string OPId)
        {
            string baseUrl = "http://" + Request.Host.Value + "/CADFiles/";
            UserData userdata = UserInfo();
            var result = new List<OPDetail>();

            //if (!ValidOrder(OrderID, OPId))
            //{
            //    return result;
            //}

            var SqlStr = "";
            SqlStr = @$";WITH cte AS
                        (
                           SELECT
                                 ROW_NUMBER() OVER (PARTITION BY OrderID ORDER BY [Range]) AS rn,*
                           FROM 
                           (select 
	                        ord.OrderID as ERPOrderID,a.OrderID,a.OPID,a.[Range],a.OPLTXA1,a.MAKTX,c.[Name],a.StartTime,a.EndTime,a.AssignDate,a.WorkGroup as remark,d.GroupName,a.ImgPath
	                        ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=a.Operator) as [user_name]
	                        ,Progress = cast( (cast(b.QtyGood as float) + cast(b.QtyBad as float)) / cast(a.OrderQTY as float) * 100 as int)
	                        ,RemainingCount = (a.OrderQTY - b.QtyTol)
	                        ,a.OrderQTY,b.QtyGood,b.QtyBad,b.WIPEvent,b.StartTime as wipStartTime,b.EndTime as wipEndTime
                            ,(select top(1) [user_name] from {_ConnectStr.AccuntDB}.dbo.[User] where [user_id]=f.QCman) as QCman
                            ,g.CustomerInfo
							,d.ID as DeviceID
                            ,(SELECT COUNT(*) FROM {_ConnectStr.MRPDB}.dbo.QCrule as qcr WHERE qcr.id=a.OPID) AS QC_count
                            ,(select COUNT(*) FROM {_ConnectStr.APSDB}.dbo.QCPointValue as qcpv where qcpv.WorkOrderID=a.OrderID and qcpv.OPID=a.OPID and qcpv.MAKTX=a.MAKTX) AS QCv_count
	                        from {_ConnectStr.APSDB}.[dbo].Assignment as a
	                        inner join {_ConnectStr.APSDB}.[dbo].WIP as b on a.SeriesID=b.SeriesID
                            left join {_ConnectStr.APSDB}.[dbo].[OrderOverview] as ord on a.ERPOrderID=ord.OrderID
	                        left join {_ConnectStr.MRPDB}.dbo.Part as c on a.MAKTX=c.Number
	                        left join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                            left join {_ConnectStr.APSDB}.[dbo].QCAssignment as f on a.OrderID=f.WorkOrderID and a.OPID=f.OPID
                            left join {_ConnectStr.APSDB}.[dbo].OrderOverview as g on a.ERPOrderID=g.OrderID
	                        where b.WIPEvent!=3) as a
                        )
                        SELECT *
                        FROM cte as aa
						left join {_ConnectStr.AccuntDB}.dbo.GroupDevice as bb on aa.DeviceID=bb.DeviceId
						where OrderID=@OrderID and OPID=@OPId and  bb.GroupSeq=2
                        order by OrderID,Range";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(("@OPId"), SqlDbType.NVarChar).Value = OPId;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                string imgPath = "N/A";
                                if (_ConnectStr.Debug == 1)
                                    imgPath = "https://i.imgur.com/pOxVHtB.jpg";
                                else
                                    imgPath = baseUrl + (string.IsNullOrEmpty(SqlData["ImgPath"].ToString().Trim()) ? "N/A" : SqlData["ImgPath"].ToString());

                                string delaydaydisplay = "";
                                int dd = deleyday(SqlData["AssignDate"].ToString().Trim(), SqlData["EndTime"].ToString());
                                if (dd < 0)
                                    delaydaydisplay = "延遲 " + Math.Abs(dd) + " 天";
                                else
                                    delaydaydisplay = "還有 " + Math.Abs(dd) + " 天";

                                result.Add(new OPDetail
                                {
                                    ERPorderID = string.IsNullOrEmpty(SqlData["ERPOrderID"].ToString().Trim()) ? "N/A" : SqlData["ERPOrderID"].ToString().Trim(),
                                    OrderId = string.IsNullOrEmpty(SqlData["OrderID"].ToString().Trim()) ? "N/A" : SqlData["OrderID"].ToString(),
                                    OPId = string.IsNullOrEmpty(SqlData["OPID"].ToString().Trim()) ? "N/A" : SqlData["OPID"].ToString(),
                                    OPName = string.IsNullOrEmpty(SqlData["OPLTXA1"].ToString().Trim()) ? "N/A" : SqlData["OPLTXA1"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["AssignDate"]).ToString(_dateFormat),
                                    Deviec = string.IsNullOrEmpty(SqlData["remark"].ToString().Trim()) ? "N/A" : SqlData["remark"].ToString(),
                                    //暫時替換wipEndTime
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),//checkNoword(SqlData["StartTime"].ToString().Trim()),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),//checkNoword(SqlData["EndTime"].ToString().Trim()),
                                    OperatorName = string.IsNullOrEmpty(SqlData["user_name"].ToString().Trim()) ? "N/A" : SqlData["user_name"].ToString(),
                                    ProductId = string.IsNullOrEmpty(SqlData["MAKTX"].ToString().Trim()) ? "N/A" : SqlData["MAKTX"].ToString(),
                                    ProductName = "【" + (string.IsNullOrEmpty(SqlData["CustomerInfo"].ToString().Trim()) ? "NA/NA/NA" : SqlData["CustomerInfo"].ToString().Trim()).Split("/")[1] + "】" + (string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? "N/A" : SqlData["Name"].ToString()),
                                    RequireNum = string.IsNullOrEmpty(SqlData["OrderQTY"].ToString().Trim()) ? "N/A" : SqlData["OrderQTY"].ToString(),
                                    ImgPath = imgPath,

                                    OPStatus = string.IsNullOrEmpty(SqlData["WIPEvent"].ToString().Trim()) ? "N/A" : SqlData["WIPEvent"].ToString(),
                                    wipStartTime = string.IsNullOrEmpty(SqlData["wipStartTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["wipStartTime"]).ToString(_timeFormat),
                                    wipEndTime = string.IsNullOrEmpty(SqlData["wipEndTime"].ToString()) ? "N/A" : Convert.ToDateTime(SqlData["wipEndTime"]).ToString(_timeFormat),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString().Trim()),
                                    CompleteNum = string.IsNullOrEmpty(SqlData["QtyGood"].ToString().Trim()) ? "N/A" : SqlData["QtyGood"].ToString(),
                                    DefectiveNum = string.IsNullOrEmpty(SqlData["QtyBad"].ToString().Trim()) ? "N/A" : SqlData["QtyBad"].ToString(),
                                    QtyBad = string.IsNullOrEmpty(SqlData["QtyBad"].ToString().Trim()) ? "N/A" : SqlData["QtyBad"].ToString(),

                                    DeleyDays = delaydaydisplay,
                                    DeviceGroup = string.IsNullOrEmpty(SqlData["GroupName"].ToString().Trim()) ? "N/A" : SqlData["GroupName"].ToString(),
                                    //IsQC = "N/A",//2023.06.17 Ryan修改 皆無須檢驗
                                    IsQC = QCStatus(IsQCDone(SqlData["OrderID"].ToString().Trim(), SqlData["OPID"].ToString().Trim()).ToString()),
                                    //QCman = string.IsNullOrEmpty(SqlData["QCman"].ToString().Trim()) ? "N/A" : SqlData["QCman"].ToString(),
                                    QCman = userdata.User_Id.ToString(),
                                });

                                //若需檢驗，判斷檢驗項目
                                if (result[0].IsQC=="False")
                                {
                                    //判斷是否有檢驗點要顯示
                                    var QC = QCList(OrderID, OPId);
                                    if (QC.Count == 1)
                                    {
                                        if (!QC.Exists(x => x.QCPointName == "N/A"))
                                        {
                                            result[0].QCList = QC;
                                        }
                                    }
                                    else
                                    {
                                        result[0].QCList = QC;
                                    }
                                }
                            }
                        }
                    }
                }
            }

            foreach (var item in result)
            {
                //判斷機台是否有委外廠商名稱，若有則顯示委外廠商名稱
                bool hasBrackets = item.Deviec.Contains("(") && item.Deviec.Contains(")");
                if (hasBrackets)
                {
                    int start = item.Deviec.IndexOf("(");
                    int end = item.Deviec.IndexOf(")");
                    if (start >= 0 && end > start)
                    {
                        string factory = item.Deviec.Substring(start + 1, end - start - 1);
                        item.ProductName += $"({factory})";
                    }

                }
            }
            return result;
        }

        private string checkNoword(string data)
        {
            if (data == "")
            {
                return "N/A";
            }
            else
            {
                return data;
            }
        }
        private int ChangeProgressIntFormat(string data)
        {
            if (data != "" && data != "N/A")
            {
                int n;
                if (int.TryParse(data, out n))
                {
                    if (n > 100) n = 100;
                    return n;
                }
                else
                {
                    return 0;
                }
            }
            return 0;
        }

        /// <summary>
        /// 該製程所有QC
        /// </summary>
        /// <param name="OrderID"></param>
        /// <param name="OPId"></param>
        /// <returns></returns>

        private List<QC> QCList(string OrderID, string OPId)
        {
            var result = new List<QC>();
            var SqlStr = @$"SELECT a.OrderID ,a.OPID,a.MAKTX,c.WIPEvent,b.QCPoint,b.QCPointName,b.QCLSL,b.QCUSL,d.QCPointValue,d.QCToolId,d.QCunit,d.Createtime,d.Lastupdatetime
                            FROM {_ConnectStr.APSDB}.dbo.Assignment as a
                            INNER JOIN {_ConnectStr.MRPDB}.dbo.QCrule as b on a.OPID=b.id
                            LEFT JOIN {_ConnectStr.APSDB}.dbo.QCPointValue as d ON a.OrderID=d.WorkOrderID AND b.id=d.OPID and b.QCPoint=d.QCPoint
                            LEFT JOIN {_ConnectStr.APSDB}.dbo.WIP as c on a.OrderID=c.OrderID AND a.OPID=c.OPID
                            WHERE a.OrderID= @OrderID AND a.OPID= @OPId";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(("@OPId"), SqlDbType.Float).Value = OPId;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                bool enable = true;
                                if (SqlData["WIPEvent"].ToString().Trim() == "0") enable = false;
                                result.Add(new QC
                                {
                                    QCPointName = checkNoword(SqlData["QCPointName"].ToString().Trim()),
                                    QCPointValue = checkNoword(SqlData["QCPointValue"].ToString().Trim()),
                                    QCResult = checkQCResult(SqlData),
                                    QCUnit = checkNoword(SqlData["QCunit"].ToString().Trim()),
                                    QCToolID = checkNoword(SqlData["QCToolId"].ToString().Trim()),
                                    CreateTime = checkNoword(SqlData["CreateTime"].ToString().Trim()),
                                    LastUpdateTime = checkNoword(SqlData["LastUpdateTime"].ToString().Trim()),
                                    toolDiff = getToolDiff(SqlData["QCToolId"].ToString().Trim()),
                                    Enable = enable
                                });

                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 判斷檢驗結果
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string checkQCResult(SqlDataReader data)
        {
            string result = "N/A";
            if (data["QCPointValue"].ToString() != "" && data["QCUSL"].ToString() != "" && data["QCLSL"].ToString() != "")
            {
                double value = Convert.ToDouble(data["QCPointValue"].ToString());
                double USL = Convert.ToDouble(data["QCUSL"].ToString());
                double LSL = Convert.ToDouble(data["QCLSL"].ToString());
                double min = 0, max = 0; ;
                if (USL > LSL)
                {
                    max = USL; min = LSL;
                }
                else
                {
                    max = LSL; min = USL;
                }

                if ((value >= min) && (value < max))
                {
                    result = "GO";
                }
                else
                {
                    result = "NOGO";
                }
            }
            return result;
        }

        private List<string> getToolDiff(string id)
        {
            List<string> result = new List<string>();
            if (_ConnectStr.Debug == 1)
            {
                //隨機給值
                string[] mockdata1 = { "位置1 : ", "位置2 : " };
                string[] mockdata2 = { "+0.001", "-0.001", "+0.003", "-0.003", "+0.999,", "-9.999", "+0.000", };
                Random r = new Random();

                if (id != "" || id != null)
                {
                    result.Add(mockdata1[0] + mockdata2[r.Next(0, mockdata2.Length)]);
                    result.Add(mockdata1[1] + mockdata2[r.Next(0, mockdata2.Length)]);
                }
                else
                {
                    result.Add(mockdata1[0] + "-");
                    result.Add(mockdata1[1] + "-");
                }
            }
            else
            {
                //從資料庫取得，資料庫還沒準備
                MessureService messureService = new MessureService();
                messureService.connectString = _ConnectStr.Local;
                string sqlcmd = @$"SELECT * FROM {_ConnectStr.MeasureDB}.[dbo].[Tolerance] where ToolID='{id}'";
                var ans = messureService.getToolTolerance(sqlcmd);
                int i = 0;
                foreach (var item in ans)
                {
                    result.Add("位置" + item.position + ": " + item.value);
                    i++;
                }
            }
            return result;
        }

        /// <summary>
        /// 開始加工
        /// </summary>
        /// <returns></returns>
        [HttpPost("WorkStartReport")]
        //public string WorkStartReport(WorkReporStarttRequest request)
        public string WorkStartReport(string OrderID, string OPId, string Device, int OrderQTY)
        {
            //將工單號、製程、機台寫入wipc和WipRegisterLog綁定機台，如果wip已經有重複供單請Update WIP資料，wipevent修改
            string result = "Update Failed!";
            int EffectRow = 0;
            UserData userdata = UserInfo();
            //確認一下WIP是否已經存在報工資料
            var SqlStr = @$" IF NOT EXISTS(SELECT *  FROM {_ConnectStr.APSDB}.[dbo].[WIP] WHERE OrderID=@OrderID AND OPID=@OPID)
	                            BEGIN 
		                            INSERT INTO {_ConnectStr.APSDB}.[dbo].[WIP]
		                            ([OrderID],[OPID],[OrderQTY],[QtyTol],[QtyGood],[QtyBad],[WIPEvent],[WorkGroup],[FirstStart],[FirstEnd],[StartTime],[EndTime],[CreateTime],[UpdateTime]) 
		                            VALUES
		                            (@OrderID, @OPID, @OrderQTY, 0, 0, 0, 1, @Device, GETDATE(), NULL, GETDATE(), NULL, GETDATE(), GETDATE()) ;

                                    INSERT INTO  {_ConnectStr.APSDB}.[dbo].[WIPLog](OrderID, OPID, QtyGood, QtyBad, DeviceID, ResultID, WIPEvent, CreateTime, StaffID)
                                    VALUES (@OrderID, @OPID, 0, 0, @Device ,NULL ,1 ,GETDATE() ,@UserID );

                                    UPDATE {_ConnectStr.APSDB}.[dbo].[WipRegisterLog]
		                            SET WorkOrderID=@OrderID, OPID=@OPID, OperatorID=@UserID, LastUpdateTime=GETDATE()
		                            WHERE DeviceID=(SELECT TOP(1) ID FROM {_ConnectStr.APSDB}.dbo.Device WHERE remark=@Device);

                                    UPDATE  {_ConnectStr.APSDB}.[dbo].[Assignment] SET Operator = @UserID WHERE OrderID = @OrderID AND OPID= @OPID

                                END 
                            ELSE
	                            BEGIN
		                            UPDATE {_ConnectStr.APSDB}.[dbo].[WIP]
		                            SET WIPEvent=1, WorkGroup=@Device, StartTime = GETDATE(), UpdateTime = GETDATE()
		                            WHERE OrderID = @OrderID AND OPID = @OPID;

                                    INSERT INTO  {_ConnectStr.APSDB}.[dbo].[WIPLog](OrderID, OPID, QtyGood, QtyBad, DeviceID, ResultID, WIPEvent, CreateTime, StaffID)
                                    VALUES (@OrderID, @OPID, 0, 0, @Device ,NULL ,1 ,GETDATE() ,@UserID );

                                    UPDATE {_ConnectStr.APSDB}.[dbo].[WipRegisterLog]
		                            SET WorkOrderID=@OrderID, OPID=@OPID, OperatorID=@UserID, LastUpdateTime=GETDATE()
		                            WHERE DeviceID=(SELECT TOP(1) ID FROM {_ConnectStr.APSDB}.dbo.Device WHERE remark=@Device);

                                    UPDATE  {_ConnectStr.APSDB}.[dbo].[Assignment] SET Operator = @UserID WHERE OrderID = @OrderID AND OPID= @OPID
	                            END";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = OPId;
                    comm.Parameters.Add(("@OrderQTY"), SqlDbType.Int).Value = OrderQTY;
                    comm.Parameters.Add(("@Device"), SqlDbType.VarChar).Value = Device;
                    comm.Parameters.Add(("@UserID"), SqlDbType.Char).Value = userdata.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = "Update Successful!";
                try
                {
                    UpdateERPOrderStatus(OrderID, OPId, 6);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                result = "Update Failed!";
            }
            return result;
        }

        //開工:6、暫停:2、完工:7
        private void UpdateERPOrderStatus(string orderID, string opid, int status)
        {
            string SeriesID = getSeriesID(orderID, opid);

            string cmd = @"[{""Id"": " + SeriesID + @",""Status"": " + status + "}]";

            string ERPServerUrl = _ConnectStr.ERPurl;
            string SetAccessToken = WebAPIservice.gettoken(ERPServerUrl);

            WebAPIservice.RequestWebAPI_PUT(cmd, ERPServerUrl + "api/WorkShare/JobOrderUpdateStatus", SetAccessToken);
        }

        private string getSeriesID(string orderID, string opid)
        {
            string result = "";
            string SqlStr = @$"SELECT * FROM {_ConnectStr.APSDB}.[dbo].[Assignment] where OrderID=@OrderID and OPID=@OPID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = orderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = opid;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            result = SqlData["SeriesID"].ToString();
                        }
                    }
                }
            }
            return result;
        }

        private bool canWork(string OrderID, float OPID)
        {
            string sqlStr1 = @$"SELECT TOP (1) [WIPEvent]
                              FROM {_ConnectStr.APSDB}.[dbo].[Assignment] as a
                              LEFT JOIN {_ConnectStr.APSDB}.[dbo].[WIP] as b
                              ON a.[OrderID] =  b.[OrderID] and a.[OPID] = b.[OPID]
                              WHERE a.[OrderID] = @OrderID and a.[OPID] < @OPID 
                              ORDER BY b.[OPID] DESC";
            string sqlStr2 = @$"SELECT TOP (1) [WIPEvent]
                              FROM {_ConnectStr.APSDB}.[dbo].[Assignment] as a
                              LEFT JOIN {_ConnectStr.APSDB}.[dbo].[WIP] as b
                              ON a.[OrderID] =  b.[OrderID] and a.[OPID] = b.[OPID]
                              WHERE a.[OrderID] = @OrderID and a.[OPID] = @OPID";

            string last_WIPevent = ""; //上一個製程
            string cur_WIPevent = ""; //當前製程
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(sqlStr1, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(@"OrderID", SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(@"OPID", SqlDbType.Int).Value = OPID;
                    using (SqlDataReader sqlData = comm.ExecuteReader())
                    {
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                last_WIPevent = sqlData["WIPEvent"].ToString().Trim();
                            }
                        }
                    }
                }
                using (var comm = new SqlCommand(sqlStr2, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(@"OrderID", SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(@"OPID", SqlDbType.Int).Value = OPID;
                    using (SqlDataReader sqlData = comm.ExecuteReader())
                    {
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                cur_WIPevent = sqlData["WIPEvent"].ToString().Trim();
                            }
                        }
                    }
                }
                // 上一個製程WIPEvent = 3 則可開始加工
                if (cur_WIPevent != "0")
                    return true;
                else
                {
                    if ((last_WIPevent == "3" || last_WIPevent == "") && cur_WIPevent == "0")
                        return true;
                    else
                        return false;
                }


            }
        }

        ////暫時無解，需要請載另外判斷製程名稱來篩選
        //private void UpdateERPOrderStatus(string data, string Mode)
        //{
        //    erpService erpService = new erpService();
        //    string cmd = @" [{ ""Number"": """+RenameOrderID(data)+""",""ProcessName"":""PMC測試製程E_替換機台""}]"" ";
        //    string recivestring = erpService.RequestWebAPI(@"", erpService.ERPServerUrl + "/api/JobOrders/Query", erpService.SetAccessToken);
        //    List<JobOrder_Query.Root> dataList = JsonConvert.DeserializeObject<List<JobOrder_Query.Root>>(recivestring);


        //    cmd = @"[{""Id"": " + dataList[0].Id + @",""Status"": " + Mode + "}]"; //報工開始
        //    string joborder = erpService.RequestWebAPI_PUT(cmd, erpService.ERPServerUrl + "/api/WorkShare/JobOrderUpdateStatus", erpService.SetAccessToken);
        //}

        private string RenameOrderID(string data)
        {
            try
            {
                string type = "";
                string year = "";
                string mouth = "";
                string day = "";
                string number = "";

                type = "JO";
                year = "20" + data.Substring(1, 2);
                mouth = data.Substring(3, 2);
                day = data.Substring(5, 2);
                number = "00" + data.Substring(7, 3);

                return type + year + mouth + day + number;
            }
            catch
            {
                return "";
            }
        }

        /// <summary>
        /// 暫停加工
        /// </summary>
        /// <param name="request">{"orderID": "1219111329","opId": "30","device": "2M302","qtyTol": 2,"reason": [{"qcReason": "外徑過小","qtyNum": 1}]}</param>
        /// <returns></returns>
        [HttpPost("WorkPassReport")]
        public string WorkPassReport(WorkPassReportRequest request)
        {
            string result = "Update Failed!";
            int EffectRow = 0;
            UserData userdata = UserInfo();

            //計算良品數量&不良品數量
            int QtyBadCount = 0;
            if (request.Reason.Count != 0)
            {
                for (int i = 0; i < request.Reason.Count; i++)
                {
                    string ReportReason = request.Reason[i].QCReason;
                    int QtyNum = Convert.ToInt32(request.Reason[i].QtyNum);
                    QtyBadCount += QtyNum;
                }
            }

            var SqlStr = @$"IF EXISTS(SELECT * FROM {_ConnectStr.APSDB}.[dbo].[WIP] WHERE OrderID=@OrderID AND OPID=@OPID)
	                            BEGIN 
		                            UPDATE {_ConnectStr.APSDB}.[dbo].[WIP]
		                            SET QtyGood += @QtyGood, QtyBad += @QtyBad, QtyTol += @QtyTol, WIPEvent=2, UpdateTime = GETDATE()
		                            WHERE OrderID = @OrderID AND OPID = @OPID AND WorkGroup = @Device;

                                    update {_ConnectStr.APSDB}.[dbo].Assignment set Operator=@UserID where OrderID=@OrderID and OPID=@OPID

                                    INSERT INTO {_ConnectStr.APSDB}.[dbo].WIPLog
                                    VALUES (@OrderID, @OPID, @QtyGood, @QtyBad, @Device ,NULL ,2 ,GETDATE() ,@UserID );

                                    UPDATE {_ConnectStr.APSDB}.[dbo].[WipRegisterLog]
		                            SET WorkOrderID=NULL, OPID=NULL, OperatorID=NULL, LastUpdateTime=GETDATE()
		                            WHERE DeviceID=(SELECT TOP(1) ID FROM {_ConnectStr.APSDB}.dbo.Device WHERE remark=@Device);
	                            END";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = request.OrderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.Float).Value = request.OPId;
                    comm.Parameters.Add(("@QtyTol"), SqlDbType.Int).Value = request.QtyTol;
                    comm.Parameters.Add(("@QtyGood"), SqlDbType.Int).Value = request.QtyTol;
                    comm.Parameters.Add(("@QtyBad"), SqlDbType.Int).Value = QtyBadCount;
                    comm.Parameters.Add(("@Device"), SqlDbType.VarChar).Value = request.Device;
                    comm.Parameters.Add(("@UserID"), SqlDbType.Char).Value = userdata.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = "Update Successful!";
                //UpdateERPOrderStatus(request.OrderID, "2");
                try
                {
                    UpdateERPOrderStatus(request.OrderID, request.OPId, 2);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                result = "Update Failed!";
            }
            return result;
        }

        /// <summary>
        /// 取得人員/機台待機原因清單
        /// </summary>
        /// <param name="Category">人員:1、機台:2</param>
        /// <returns>"type":1(人員)、2(機台)</returns>
        [HttpPost("Idlereasonlist")]
        public List<idlereason> Idlereasonlist(int Category)
        {
            var result = new List<idlereason>();
            string SqlStr = @$"SELECT * FROM {_ConnectStr.APSDB}.[dbo].[IdleResult]
                                where Category=@Category and exist=1
                                ORDER BY id DESC";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add(("@Category"), SqlDbType.Int).Value = Category;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new idlereason
                                {
                                    idleReasonId = SqlData["ID"].ToString(),
                                    idleReasonType = SqlData["Category"].ToString(),
                                    idleReasonTitle = checkNoword(SqlData["item"].ToString()),
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 回報待機開始時間
        /// </summary>
        /// <param name="request">{"orderId": "1219111329","opid": "30","device": "2M302","reasonCode": "80"}</param>
        /// <returns></returns>
        [HttpPost("StartIdleReport")]
        public string StartIdleReport(StartIdleReportRequest request)
        {
            string result = "";
            int EffectRow = 0;
            UserData userData = UserInfo();
            //判斷待機是否已經存在

            //DECLARE @OrderID  NVARCHAR(30)
            //SET @OrderID  =(SELECT TOP(1) OrderID FROM WIPLog WHERE StaffID=@StaffID and WIPEvent=1 ORDER BY CreateTime DESC)
            //DECLARE @OPID  FLOAT
            //SET @OPID=(SELECT TOP(1) OPID FROM WIPLog WHERE StaffID=@StaffID and WIPEvent=1 ORDER BY CreateTime DESC)
            string SqlStr = @$"

                        IF NOT EXISTS(SELECT * FROM {_ConnectStr.APSDB}.[dbo].[IdleReasonBinding] WHERE [OrderID]=@OrderID AND [OPID]=@OPID AND ReasonCode=@ReasonCode AND StaffID=@StaffID)
                            BEGIN 
                                INSERT INTO {_ConnectStr.APSDB}.[dbo].[IdleReasonBinding] ([OrderID],[OPID],[StartTime],[EndTime],[ReasonCode],[StaffID],[Device],[Updated])
                                VALUES (@OrderID, @OPID, GETDATE(), NULL, @ReasonCode, @StaffID,@Device, 0);
                            END 
                        ELSE
                            BEGIN
	                             UPDATE {_ConnectStr.APSDB}.[dbo].[IdleReasonBinding] SET [StartTime]=GETDATE(), [ReasonCode]=@ReasonCode, [Updated]+=1
		                         WHERE [OrderID]=@OrderID AND [OPID]=@OPID AND [StaffID]=@StaffID;
                            END";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = request.OrderId;
                    comm.Parameters.Add("@OPID", SqlDbType.Float).Value = request.OPID;
                    comm.Parameters.Add("@Device", SqlDbType.NVarChar).Value = request.Device;
                    comm.Parameters.Add("@ReasonCode", SqlDbType.Int).Value = request.ReasonCode;
                    comm.Parameters.Add("@StaffID", SqlDbType.Int).Value = userData.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = "Update Successful!";
            }
            else
            {
                result = "Update Failed!";
            }
            return result;
        }

        /// <summary>
        /// 回報待機結束時間
        /// </summary>
        /// <param name="request">{"orderId": "1219111329","opid": "30","device": "2M302","reasonCode": "80"}</param>
        /// <returns></returns>
        [HttpPost("EndIdleReport")]
        public string EndIdleReport(EndIdleReportRequest request)
        {
            string result = "";
            int EffectRow = 0;
            UserData userData = UserInfo();
            //判斷待機是否已經存在
            string SqlStr = @$"UPDATE {_ConnectStr.APSDB}.[dbo].[IdleReasonBinding] SET EndTime=GETDATE(), Updated += 1
                            WHERE StaffID = @StaffID AND ReasonCode = @ReasonCode AND OrderID=@OrderID AND OPID=@OPID
                            ";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = request.OrderId;
                    comm.Parameters.Add("@OPID", SqlDbType.Float).Value = request.OPID;
                    comm.Parameters.Add("@Device", SqlDbType.NVarChar).Value = request.Device;
                    comm.Parameters.Add("@ReasonCode", SqlDbType.Int).Value = request.ReasonCode;
                    comm.Parameters.Add("@StaffID", SqlDbType.Int).Value = userData.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = "Update Successful!";
            }
            else
            {
                result = "Update Failed!";
            }
            return result;
        }


        /// <summary>
        /// 繼續開始加工
        /// </summary>
        /// <returns></returns>
        [HttpPost("WorkReSatrtReport")]
        //public string WorkReSatrtReport(WorkReportRequest request)
        public string WorkReSatrtReport(string OrderID, string OPId, string Device)
        {
            //將工單號、製程、機台寫入wipc和WipRegisterLog綁定機台，如果wip已經有重複供單請Update WIP資料，wipevent修改
            string result = "Update Failed!";
            int EffectRow = 0;
            UserData userdata = UserInfo();
            //確認一下WIP是否已經存在報工資料
            var SqlStr = @$" IF EXISTS(SELECT *  FROM {_ConnectStr.APSDB}.[dbo].[WIP] WHERE OrderID=@OrderID AND OPID=@OPID)
	                            BEGIN
		                            UPDATE {_ConnectStr.APSDB}.[dbo].[WIP]
		                            SET WIPEvent=1, WorkGroup=@Device, StartTime = GETDATE(), UpdateTime = GETDATE()
		                            WHERE OrderID = @OrderID AND OPID = @OPId;

                                    update {_ConnectStr.APSDB}.[dbo].Assignment set Operator=@UserID where OrderID=@OrderID and OPID=@OPID

                                    INSERT INTO {_ConnectStr.APSDB}.[dbo].[WIPLog](OrderID, OPID, QtyGood, QtyBad, DeviceID, ResultID, WIPEvent, CreateTime, StaffID)
                                    VALUES (@OrderID, @OPId, 0, 0, @Device ,NULL ,1 ,GETDATE() ,@UserID );

                                    UPDATE {_ConnectStr.APSDB}.[dbo].[WipRegisterLog]
		                            SET WorkOrderID=@OrderID, OPID=@OPID, OperatorID=@UserID, LastUpdateTime=GETDATE()
		                            WHERE DeviceID=(SELECT TOP(1) ID FROM {_ConnectStr.APSDB}.dbo.Device WHERE remark=@Device);

                                    UPDATE {_ConnectStr.APSDB}.[dbo].[Assignment] SET Operator = @UserID WHERE OrderID = @OrderID AND OPID= @OPID
	                            END";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = OPId;
                    comm.Parameters.Add(("@Device"), SqlDbType.VarChar).Value = Device;
                    comm.Parameters.Add(("@UserID"), SqlDbType.Char).Value = userdata.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = "Update Successful!";
                //UpdateERPOrderStatus(OrderID, "6");
                try
                {
                    UpdateERPOrderStatus(OrderID, OPId, 6);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                result = "Update Failed!";
            }
            return result;
        }

        /// <summary>
        /// 填寫QCPoint
        /// </summary>
        /// <param name="request"></param>
        /// <returns></returns>
        [HttpPost("UpdateQCPoint")]
        public string UpdateQCPoint(WorkQCReportRequest request)
        {
            //先檢查是否已經有data，如果沒有要先insert，如果有要update value和 lastupdatetime
            //判斷是否有QC權限
            //bool Valid = ValidQCOrder();
            //if (Valid == false)
            //{
            //    return Unauthorized();
            //}
            string result = "Update Failed!";
            UserData userData = UserInfo();
            int EffectRow = 0;
            var SqlStr = @$"IF NOT EXISTS(SELECT * FROM {_ConnectStr.APSDB}.[dbo].[QCPointValue] WHERE WorkOrderID = @OrderID AND OPID = @OPId AND QCPoint = (SELECT TOP(1) QCPoint
                                      FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] as a
                                      LEFT join {_ConnectStr.MRPDB}.[dbo].Process as b
                                      on a.ProcessID = b.ProcessNo
                                      where b.ID=@OPId and a.QCPointName=@QCPointName
                                      order by ProcessName,QCPoint))
                                BEGIN 
                                     INSERT INTO {_ConnectStr.APSDB}.[dbo].[QCPointValue]
                                     ([WorkOrderID], [OPID], [MAKTX], [QCPoint], [QCPointValue], [QCToolId], [QCunit], [Createtime], [Lastupdatetime],[QCman],[QCMode]) 
                                     VALUES
                                     (@OrderID, @OPId, @MAKTX, 
                                     (SELECT TOP(1) QCPoint
                                      FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] as a
                                      LEFT join {_ConnectStr.MRPDB}.[dbo].Process as b
                                      on a.ProcessID = b.ProcessNo
                                      where b.ID=@OPId and a.QCPointName=@QCPointName
                                      order by ProcessName,QCPoint), 
                                     @QCValue,
                                     @QCToolId, 
                                     'mm'
                                     ,GETDATE() ,GETDATE(),@UserName,@QCMode ) 
                                END 
                           ELSE
                                BEGIN
                                     UPDATE {_ConnectStr.APSDB}.[dbo].[QCPointValue]
                                     SET QCPointValue = @QCValue,
                                     QCunit='mm',
                                     Lastupdatetime = GETDATE(), QCman=@UserName,QCMode = @QCMode, QCToolId = @QCToolId
                                     WHERE WorkOrderID = @OrderID AND OPID = @OPId AND MAKTX = @MAKTX 
                                     AND QCPoint = (SELECT TOP(1) QCPoint
                                      FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] as a
                                      LEFT join {_ConnectStr.MRPDB}.[dbo].Process as b
                                      on a.ProcessID = b.ProcessNo
                                      where b.ID=@OPId and a.QCPointName=@QCPointName
                                      order by ProcessName,QCPoint)
                                END";

            ////Ver1. SQL語法
            //SqlStr = @$"IF NOT EXISTS(SELECT * FROM {_ConnectStr.APSDB}.[dbo].[QCPointValue] WHERE WorkOrderID = @OrderID AND OPID = @OPId AND QCPoint = (SELECT TOP(1)[QCPoint] FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] WHERE id=@OPId AND QCPointName=@QCPointName))
            //                    BEGIN 
            //                         INSERT INTO {_ConnectStr.APSDB}.[dbo].[QCPointValue]
            //                         ([WorkOrderID], [OPID], [MAKTX], [QCPoint], [QCPointValue], [QCToolId], [QCunit], [Createtime], [Lastupdatetime]) 
            //                         VALUES
            //                         (@OrderID, @OPId, @MAKTX, 
            //                         (SELECT TOP(1)[QCPoint] FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] WHERE id=@OPId AND QCPointName=@QCPointName), 
            //                         (@QCPointValue),
            //                         @QCToolId, 
            //                         (SELECT TOP(1)[unit] FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] WHERE MAKTX=@MAKTX AND opid=@OPId AND QCPointName=@QCPointName)
            //                         ,GETDATE() ,GETDATE()) 
            //                    END 
            //               ELSE
            //                    BEGIN
            //                         UPDATE {_ConnectStr.APSDB}.[dbo].[QCPointValue]
            //                         SET QCPointValue = (@QCPointValue),
            //                         Lastupdatetime = GETDATE()
            //                         WHERE WorkOrderID = @OrderID AND OPID = @OPId AND MAKTX = @MAKTX 
            //                         AND QCPoint = (SELECT TOP(1)[QCPoint] FROM {_ConnectStr.MRPDB}.[dbo].[QCrule] WHERE id=@OPId AND QCPointName=@QCPointName)
            //                    END";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = request.OrderID;
                    comm.Parameters.Add(("@OPId"), SqlDbType.Float).Value = request.OPId;
                    comm.Parameters.Add(("@MAKTX"), SqlDbType.NVarChar).Value = request.ProductId;
                    comm.Parameters.Add(("@QCPointName"), SqlDbType.NVarChar).Value = request.QCPointName;
                    comm.Parameters.Add(("@QCToolId"), SqlDbType.NVarChar).Value = request.QCToolID;
                    comm.Parameters.Add(("@UserName"), SqlDbType.NVarChar).Value = userData.EmpolyeeName;
                    comm.Parameters.Add(("@QCMode"), SqlDbType.NVarChar).Value = request.QCMode == 0 ? "Auto" : "Manual";
                    comm.Parameters.Add(("@QCValue"), SqlDbType.Float).Value = request.QCValue;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = "Update Successful!";
            }
            else
            {
                result = "Update Failed!";
            }

            ////更新QCAssignment，判斷檢測點是否都有量測值
            //var checkQC = QCList(request.OrderID, request.OPId.ToString().Trim());

            //if (!checkQC.Exists(x => x.QCPointValue == "N/A"))
            //{
            //    SqlStr = @$"update QCAssignment set IsQC=1,QCman=@Userid where WorkOrderID=@OrderID and opid=@opid";
            //    using (var conn = new SqlConnection(_ConnectStr.Local))
            //    {
            //        using (var comm = new SqlCommand(SqlStr, conn))
            //        {
            //            if (conn.State != ConnectionState.Open)
            //                conn.Open();
            //            comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = request.OrderID;
            //            comm.Parameters.Add(("@opid"), SqlDbType.Float).Value = request.OPId;
            //            comm.Parameters.Add(("@Userid"), SqlDbType.NVarChar).Value = userData.User_Id;

            //            EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
            //        }
            //    }
            //    if (EffectRow > 0)
            //    {
            //        result = "Update QCPoint & QCman Successful!";
            //    }
            //    else
            //    {
            //        result = "Update QCPoint & Update QCman Failed!";
            //    }
            //}

            return result;
        }

        /// <summary>
        /// 取得最新無線量測數據
        /// </summary>
        /// <param name="MeasureDeviceID">輸入量具編號名稱 Ex: 03</param>
        /// <returns></returns>
        [HttpPost("NewQCPointValue")]
        public ActionResponse<List<Measuredata>> NewQCPointValue(string MeasureDeviceID)
        {
            var SqlStr = @$"SELECT TOP (1) *
                            FROM {_ConnectStr.MeasureDB}.[dbo].[MeasureLog]
                            WHERE U_WAVE_T_ID=@MeasureDeviceID
                            ORDER BY [_time] DESC";
            var result = new List<Measuredata>();
            //開啟連線
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open) conn.Open();
                    comm.Parameters.Add(("@MeasureDeviceID"), SqlDbType.NVarChar).Value = MeasureDeviceID;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new Measuredata
                                {
                                    U_WAVE_T_ID = SqlData["U_WAVE_T_ID"].ToString().Replace("\r", "").Trim(),
                                    measurement_data = SqlData["measurement_data"].ToString().Replace("\r", "").Trim(),
                                    unit = IsMetric(SqlData["unit"].ToString().Replace("\r", "").Trim()),
                                    _time = SqlData["_time"].ToString().Replace("\r", "").Trim()
                                });
                            }
                        }
                    }
                }
            }
            //取得回傳資料給新的Response
            return new ActionResponse<List<Measuredata>>
            {
                Data = result
            };
        }

        private string IsMetric(string data)
        {
            string result = "mm";
            if (data != "M") result = "in";
            return result;
        }

        /// <summary>
        /// 更新量測結果
        /// </summary>
        /// <param name="OrderID"></param>
        /// <param name="OPId"></param>
        /// <returns></returns>
        [HttpPost("ReflashQCPoint")]
        public List<QC> ReflashQCPoint(string OrderID, string OPId)
        {
            var result = new List<QC>();
            result = QCList(OrderID, OPId);

            return result;
        }


        private object IsQCDone(string OrderID, string OPId)
        {
            string ans = "N/A";
            var result = new List<QC>();

            result = QCList(OrderID, OPId);

            if (result.Count == 0)
            {
                ans = "N/A";
            }
            else
            {
                //確認是否有量測點未量測
                if (result.Exists(x => x.QCPointName == "N/A"))
                {
                    ans = "N/A";
                }
                else
                {
                    if (result.Exists(x => x.QCPointValue == "N/A"))
                    {
                        ans = "False";
                    }
                    else
                    {
                        ans = "True";
                    }
                }
            }

            return ans;
        }

        private string IsQCDone_2(string requiredQty, string DoneQty)
        {
            string ans = "N/A";
            if (int.Parse(requiredQty) != 0)
            {
                ans = "False";//未完成，需檢驗
            }
            else if (int.Parse(requiredQty) - int.Parse(DoneQty) == 0)
            {
                ans = "Ture";//已完成，無須檢驗
            }
            return ans;
        }

        /// <summary>
        /// 結束加工
        /// </summary>
        /// <param name="request">{"orderID": "1219111325","opId": 50,"qtyTol": 18,"reason": []}</param>
        /// <returns></returns>
        [HttpPost("WorkDoneReport")]
        public object WorkDoneReport(RequestReportInfo request)
        {
            //bool Valid = ValidOrder(requestReportInfo.OrderId, requestReportInfo.OPId);
            //if (Valid == false)
            //{
            //    return Unauthorized();
            //}
            UserData userdata = UserInfo();

            //完工回報數量和不良品原因，要修改WIP和新增WIPLOG，如果有不良品原因要在新增至QCRresult
            string result = "";
            int EffectRow = 0;
            string SqlStr = @$"INSERT INTO {_ConnectStr.APSDB}.[dbo].[WIPLog] 
                            VALUES (@OrderID, @OPID, @QtyGood, @QtyBad, (SELECT TOP(1) WorkGroup FROM WIP WHERE OrderID=@OrderID and OPID=@OPID), 0, '3', GETDATE(), @StaffID)

                            update {_ConnectStr.APSDB}.[dbo].Assignment set Operator=@StaffID where OrderID=@OrderID and OPID=@OPID

                            update {_ConnectStr.APSDB}.[dbo].WipRegisterLog set WorkOrderID=null,OPID=null,OperatorID=null,LastUpdateTime=GETDATE() 
                            where DeviceID=(select ID from Device where remark=(select top(1)WorkGroup from {_ConnectStr.APSDB}.[dbo].Assignment where OrderID=@OrderID and OPID=@OPID))

                            INSERT INTO {_ConnectStr.APSDB}.[dbo].QCResultLog (OrderId, OPID, QCReason, QtyNum) VALUES";
            int QtyBadCount = 0;
            if (request.Reason.Count != 0)
            {
                for (int i = 0; i < request.Reason.Count; i++)
                {

                    string ReportReason = request.Reason[i].QCReason;
                    int QtyNum = Convert.ToInt32(request.Reason[i].QtyNum);
                    QtyBadCount += QtyNum;
                    if (i == request.Reason.Count - 1)
                    {
                        SqlStr += $"(@OrderId, @OPID, '{ReportReason}',  '{QtyNum}' )";
                    }
                    else
                    {
                        SqlStr += $"('{request.OrderID}', '{request.OPId}', '{ReportReason}',  '{QtyNum}' ),";
                    }

                }
            }
            else
            {
                SqlStr = @$"INSERT INTO {_ConnectStr.APSDB}.[dbo].[WIPLog] 
                            VALUES (@OrderID, @OPID, @QtyGood, @QtyBad, (SELECT TOP(1) WorkGroup FROM WIP WHERE OrderID=@OrderID and OPID=@OPID), 0, '3', GETDATE(), @StaffID)

";
            }

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@QtyGood"), SqlDbType.Int).Value = request.QtyTol;
                    comm.Parameters.Add(("@QtyBad"), SqlDbType.Int).Value = QtyBadCount;
                    comm.Parameters.Add(("@QtyTol"), SqlDbType.Int).Value = request.QtyTol;
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = request.OrderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.Float).Value = request.OPId;
                    comm.Parameters.Add(("@StaffID"), SqlDbType.NVarChar).Value = userdata.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {
                result = UpdateWIP(request.QtyTol, request.OrderID, request.OPId, QtyBadCount);
                //UpdateERPOrderStatus(request.OrderID, "7");
                try
                {
                    UpdateERPOrderStatus(request.OrderID, request.OPId, 7);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                result = "Insert Failed!";
            }
            return result;
        }

        /// <summary>
        /// UpdateWIP
        /// </summary>
        /// <param name="QtyTol"></param>
        /// <param name="OrderId"></param>
        /// <param name="OPId"></param>
        /// <param name="QtyBadCount"></param>
        /// <returns></returns>
        private string UpdateWIP(int QtyTol, string OrderId, string OPId, int QtyBadCount)
        {
            UserData userdata = UserInfo();
            string result = "";
            int EffectRow = 0;
            string SqlStr = @$"
                            UPDATE {_ConnectStr.APSDB}.[dbo].[WIP] SET 
                            WIPEvent='3', QtyGood += @QtyGood, QtyBad += @QtyBad, QtyTol += @QtyTol, EndTime=GETDATE(), UpdateTime=GETDATE()
                            WHERE OrderID=@OrderID AND OPID=@OPID AND WIPEvent=1

                            UPDATE {_ConnectStr.APSDB}.dbo.Assignment SET Operator=@Operator where OrderID=@OrderID and OPID=@OPID
                            ";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@QtyTol"), SqlDbType.NVarChar).Value = QtyTol;
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderId;
                    comm.Parameters.Add(("@OPID"), SqlDbType.NVarChar).Value = OPId;
                    comm.Parameters.Add(("@QtyGood"), SqlDbType.NVarChar).Value = QtyTol;
                    comm.Parameters.Add(("@QtyBad"), SqlDbType.NVarChar).Value = QtyBadCount;
                    comm.Parameters.Add(("@Operator"), SqlDbType.Char).Value = userdata.User_Id;
                    EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                }
            }
            if (EffectRow > 0)
            {

                SqlStr = @$"UPDATE {_ConnectStr.APSDB}.[dbo].Assignment SET Operator = (SELECT [user_id] FROM {_ConnectStr.AccuntDB}.dbo.[User] WHERE user_account = @user_account) WHERE OrderID = @OrderID AND OPID= @OPID

                            update {_ConnectStr.APSDB}.[dbo].WipRegisterLog set WorkOrderID=null,OPID=null,OperatorID=null 
                            where DeviceID=(
	                            select top(1)ID from {_ConnectStr.APSDB}.[dbo].Device where remark=(
		                            select top(1)WorkGroup from {_ConnectStr.APSDB}.[dbo].WIP where OrderID=@OrderID and OPID=@OPID
	                            )
                            )";
                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(SqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        comm.Parameters.Add(("@user_account"), SqlDbType.NVarChar).Value = User.Identity.Name;
                        comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderId;
                        comm.Parameters.Add(("@OPId"), SqlDbType.NVarChar).Value = OPId;
                        EffectRow = Convert.ToInt32(comm.ExecuteNonQuery());
                    }
                }
                if (EffectRow > 0)
                {
                    result = "WIP and Assignment Update Successful!";
                }
                else
                {
                    result = "UpdateWIP Update Successfull but Assignment Update Failed";
                }
            }
            else
            {
                result = "WIP Update Failed!";
            }
            return result;
        }

        /// <summary>
        /// 不良品原因列表
        /// </summary>
        /// <returns></returns>
        [HttpGet("DefectiveList")]
        public List<DefectiveItem> DefectiveList()
        {
            var result = new List<DefectiveItem>();
            var SqlStr = @$"SELECT * FROM {_ConnectStr.MRPDB}.[dbo].[DefectiveItem] WHERE [Status]=1";
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
                                result.Add(new DefectiveItem
                                {
                                    Id = SqlData["Id"].ToString(),
                                    DefectiveName = SqlData["DefectiveItems"].ToString(),
                                    Remark = checkNoword(SqlData["Remark"].ToString()),
                                    FilePath = checkNoword(SqlData["FilePath"].ToString())
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        public class DefectiveItem
        {
            /// <summary>
            /// 不良原因編號'
            /// </summary>
            public string Id { get; set; }
            /// <summary>
            /// 不良原因名稱
            /// </summary>
            public string DefectiveName { get; set; }
            /// <summary>
            /// 不良原因備註
            /// </summary>
            public string Remark { get; set; }
            /// <summary>
            /// 檔案路徑
            /// </summary>
            public string FilePath { get; set; }
        }


        /// <summary>
        /// 讀取量測設備取得數值
        /// </summary>
        /// <param name="query"></param>
        /// <returns></returns>
        [HttpPost("measuredvalue")]
        public ActionResponse<List<string>> measuredvalue(measurervalueQuery query)
        {
            if (query.id == null || query.id == "" || query.id == "string"
                || query.time == null || query.time == "" || query.time == "string")
            {
                throw new Exception("No Tool Id, Time data");
            }

            var result = new List<string>();
            if (_ConnectStr.Debug == 1)
            {
                Random r = new Random(Guid.NewGuid().GetHashCode());
                int count = r.Next(1, 6);
                for (int i = 0; i < count; i++)
                {
                    result.Add((r.Next(205, 206) + r.NextDouble()).ToString("F3"));
                }
            }
            else
            {
                //var SqlStr = @$"SELECT TOP(6)* FROM {_ConnectStr.MeasureDB}.[dbo].[MeasureLog] where U_WAVE_T_ID='{query.id}' and _time > '{query.time}' order by _time desc";
                var SqlStr = @$"SELECT TOP(3)* FROM {_ConnectStr.MeasureDB}.[dbo].[MeasureLog] where U_WAVE_T_ID='{query.id}' order by _time desc";

                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(SqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        //comm.Parameters.Add(("@toolid"), SqlDbType.NChar).Value = query.id;
                        //comm.Parameters.Add(("@time"), SqlDbType.DateTime).Value = query.time;
                        using (SqlDataReader SqlData = comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                while (SqlData.Read())
                                {
                                    result.Add(SqlData["measurement_data"].ToString().Trim());
                                }
                            }
                        }
                    }
                }
            }

            result.Reverse();//倒序一下

            return new ActionResponse<List<string>>
            {
                Data = result
            };
        }


    }
}