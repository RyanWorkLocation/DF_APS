using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using PMCDash.Services;
using PMCDash.Models;
using PMCDash.DTO;
using static PMCDash.Services.AccountService;
using Newtonsoft.Json;
using System.Threading.Tasks;

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    //[ApiController]
    //public class ScheduleController : BaseApiController
    public class ScheduleController : ControllerBase
    {
        private readonly FitnessService fitnessService = new FitnessService();
        ConnectStr _ConnectStr = new ConnectStr();
        private readonly string _timeFormat = "yyyy-MM-dd HH:mm";
        private readonly string _timeFormat_ToMin = "yyyy-MM-dd HH:mm";

        /// <summary>
        /// 取得工單列表資料 (O)
        /// </summary>
        /// <returns></returns>
        [HttpGet("DashboardData")]
        public ActionResponse<List<Schedule>> DashboardData()
        {
            //checkLoginPermission();

            var result = new List<Schedule>();
            var SqlStr = $@"SELECT  
                            a.OrderID, a.ERPOrderID, a.OPID, 
                            a.Range, a.WorkGroup, a.OPLTXA1, 
                            a.MachOpTime, a.HumanOpTime, a.OrderQTY, 
                            a.StartTime, a.EndTime, a.AssignDate, 
                            a.AssignDate_PM, a.Important, 
                            a.MAKTX, a.CPK, a.Scheduled, w.WIPEvent, 
                            Progress=ROUND(cast(cast(w.QtyTol as float)/cast(w.OrderQTY as float) * 100 as decimal), 0), 
                            a.Note, 
                            (select top(1) CustomerInfo from OrderOverview where OrderID=a.ERPOrderID) as CustomerInfo
                            ,e.[Name]
                            FROM {_ConnectStr.APSDB}.[dbo].Assignment as a
                            Left JOIN {_ConnectStr.APSDB}.[dbo].WIP as w ON a.SeriesID=w.SeriesID
                            Left JOIN {_ConnectStr.MRPDB}.[dbo].[Part] as e on a.MAKTX=e.Number
                            where (w.EndTime is null or w.EndTime > GETDATE() - 30) and (w.WIPEvent is null or w.WIPEvent!=3)";
            //and substring(a.ERPOrderID,2,4)>= DATEADD(MONTH, -3, GETDATE())

            Random rnd = new Random(Guid.NewGuid().GetHashCode());

            SqlStr += " ORDER BY a.OrderID, a.Range";
            try
            {
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
                                int i = 0;
                                while (SqlData.Read())
                                {
                                    //var CustomerLocation_choose = CustomerLocation_List[1];
                                    //var CustomerName_choose = CustomerName_List[rnd.Next(CustomerName_List.Count)];

                                    string[] customer = checkCustomerInfo(string.IsNullOrEmpty(SqlData["CustomerInfo"].ToString()) ? "-" : SqlData["CustomerInfo"].ToString());


                                    result.Add(new Schedule
                                    {
                                        MAKTX = _checkNoword(SqlData["MAKTX"].ToString().Trim()),
                                        OrderID = _checkNoword(SqlData["OrderID"].ToString().Trim()),
                                        OPID = _checkNoword(SqlData["OPID"].ToString()),
                                        OPName = _checkNoword(SqlData["OPLTXA1"].ToString()),
                                        WorkGroup = _checkNoword(SqlData["WorkGroup"].ToString()),
                                        OrderQTY = ChangeIntFormat(SqlData["OrderQTY"].ToString()),
                                        StartTime = _ChangeTimeFormat(checkNoword(SqlData["StartTime"].ToString())),
                                        EndTime = _ChangeTimeFormat(checkNoword(SqlData["EndTime"].ToString())),
                                        AssignDate_PM = _checkNoword(DateTime.Parse(SqlData["AssignDate_PM"].ToString()).ToString("yyyy-MM-dd")),
                                        AssignDate = _checkNoword(DateTime.Parse(SqlData["AssignDate"].ToString()).ToString("yyyy-MM-dd")),
                                        Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString()),
                                        Note = _checkNoword(SqlData["Note"].ToString()),
                                        CustomerName = string.IsNullOrEmpty(customer[1]) ? "-" : customer[1],
                                        CustomerLocation = string.IsNullOrEmpty(new Address(string.IsNullOrEmpty(customer[2]) ? "-" : customer[2]).City) ? "-" : new Address(string.IsNullOrEmpty(customer[2]) ? "-" : customer[2]).City,
                                        Important = Convert.ToBoolean(SqlData["Important"].ToString()),
                                        PartName = string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? " - " : SqlData["Name"].ToString().Trim(),
                                        HunmanOpTime = getHumanOpTime(SqlData).ToString(@"hh\:mm\:ss"),
                                        MachOpTime = getMachOpTime(SqlData).ToString(@"hh\:mm\:ss"),
                                        WIPEvent = int.Parse(string.IsNullOrEmpty(SqlData["WIPEvent"].ToString().Trim()) ? "0" : SqlData["WIPEvent"].ToString().Trim()),
                                    });

                                    if (result[i].AssignDate != "－")
                                    {
                                        var days = Delayday(result[i].AssignDate, DateTime.Now.ToString());
                                        if (days > 0)
                                        {
                                            result[i].ProcessStatus = "DELAY";
                                            result[i].DelayDays = Convert.ToInt32(Math.Round(days, 0, MidpointRounding.AwayFromZero).ToString());
                                            //result[i].OrderHasDelayed = true;
                                        }
                                        else
                                        {
                                            if (Delayday(DateTime.Now.ToString(), result[i].AssignDate_PM) >= 0 && Delayday(DateTime.Now.ToString(), result[i].AssignDate_PM) <= 5)
                                            {
                                                result[i].ProcessStatus = "WILLDELAY";
                                                //result[i].AssignDelayIsComing = true;
                                            }
                                            else
                                            {
                                                result[i].ProcessStatus = "NORMAL";

                                            }
                                            result[i].DelayDays = 0;
                                        }
                                    }
                                    else
                                    {
                                        result[i].DelayDays = 0;
                                    }

                                    if (result[i].StartTime == "－")
                                    {
                                        result[i].Assign = false;
                                    }
                                    else
                                    {
                                        result[i].Assign = true;
                                    }
                                    i += 1;
                                }
                            }
                        }
                    }
                }
            }
            catch (SqlException sqlex)
            {
                result.Add(new Schedule
                {
                    MAKTX = sqlex.ToString()
                });
            }

            //移除已經完成的工單
            var sgg = IsDoneWorkorderList(result);
            foreach (var number in sgg)
            {
                var temp = result.Where(x => x.OrderID == number).ToList();
                foreach (var k in temp)
                {
                    result.Remove(k);
                }
            }

            return new ActionResponse<List<Schedule>>
            {
                Data = result
            };
        }

        private List<string> IsDoneWorkorderList(List<Schedule> data)
        {
            List<string> result = new List<string>();

            var orders = data.Distinct(x => x.OrderID).Select(x => x.OrderID).ToList();

            foreach (var item in orders)
            {
                var temp = data.Where(x => x.OrderID == item).ToList();
                if ((temp.Count * 3) == temp.Sum(x => x.WIPEvent))
                {
                    result.Add(item);
                }
            }

            return result;

        }

        private void checkLoginPermission()
        {
            var result = UserInfo();
            if (!result.FunctionAccess.Exists(x => x.FunctionName == "生產排程"))
                throw new UnauthorizedAccessException($"Some message");
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
        }

        private string[] checkCustomerInfo(string data)
        {
            string[] customer;
            if (data.Split("/").Count() <= 2)
            {
                if (data.Split("/")[0] != "")
                {
                    customer = new string[3];
                    customer[0] = data.Split("/")[0];
                    customer[1] = "-";
                    customer[2] = "-";
                }
                else
                {
                    customer = new string[3];
                    customer[0] = "-";
                    customer[1] = "-";
                    customer[2] = "-";
                }
            }
            else
            {
                customer = data.Split("/");
            }
            return customer;
        }

        private int ChangeIntFormat(string data)
        {
            if (data != "" && data != "N/A")
            {
                int n;
                if (int.TryParse(data, out n))
                {
                    return n;
                }
                else
                {
                    return 0;
                }
            }
            return 0;
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
        private string ChangeTimeFormat(string data)
        {
            if (data != "－" && data != "N/A")
            {
                return DateTime.Parse(data).ToString(_timeFormat);
            }
            else
            {
                if (data == "N/A") data = "－";
                return data;
            }
        }
        private string _ChangeTimeFormat(string data)
        {
            if (data != "－" && data != "N/A")
            {
                return DateTime.Parse(data).ToString(_timeFormat_ToMin);
            }
            else
            {
                if (data == "N/A") data = "－";
                return data;
            }
        }


        /// <summary>
        /// 取得工單列表資料(預覽模式) (O)*
        /// </summary>
        /// <param name="mode">【0:原始排程、1:交期優先-推薦權重、2:交期優先-自訂權重、3:機台優先、4:插單優先、5:手動排程、6:設備故障排程重排、7:設備故障排程延後工單右移】</param>
        /// <returns></returns>
        [HttpGet("DashboardPreview/{mode}")]
        public ActionResponse<List<Schedule>> DashboardPreview(string mode)
        {
            List<string> Mode_List = new List<string> { "0", "1", "2", "3", "4", "5", "6", "7" };
            var result = new List<Schedule>();
            if (Mode_List.Exists(x => x == mode))
            {
                string Assign = string.Empty;

                if (mode == "0")
                {
                    Assign = "Assignment";
                }
                else
                {
                    Assign = "AssignmentTemp" + mode;
                }

                var SqlStr = $@"SELECT  a.OrderID, a.OPID, a.WorkGroup,a.OPLTXA1, a.OrderQTY, a.StartTime, a.EndTime, a.AssignDate, a.AssignDate_PM,  a.Important, a.MAKTX, w.WIPEvent,
                                a.Scheduled,
                                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0), 
                                a.Note
                                FROM {_ConnectStr.APSDB}.[dbo].{Assign} as a
                                LEFT JOIN {_ConnectStr.APSDB}.[dbo].WIP as w ON (a.SeriesID=w.SeriesID)
                                where (w.EndTime is null or w.EndTime > GETDATE() - 30) and (w.WIPEvent is null or w.WIPEvent!=3)
                                and (a.EndTime is null or a.EndTime > GETDATE() - 30)";
                //var temp = "substring(a.ERPOrderID,2,4)>=DATEADD(MONTH, -3, GETDATE())";
                SqlStr += " ORDER BY a.WorkGroup, a.StartTime DESC";
                try
                {
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
                                    int i = 0;
                                    while (SqlData.Read())
                                    {
                                        string[] customer = checkCustomerInfo(".../.../...");
                                        result.Add(new Schedule
                                        {
                                            MAKTX = _checkNoword(SqlData["MAKTX"].ToString().Trim()),
                                            OrderID = _checkNoword(SqlData["OrderID"].ToString().Trim()),
                                            OPID = _checkNoword(SqlData["OPID"].ToString()),
                                            OPName = _checkNoword(SqlData["OPLTXA1"].ToString()),
                                            WorkGroup = _checkNoword(SqlData["WorkGroup"].ToString()),
                                            OrderQTY = ChangeIntFormat(SqlData["OrderQTY"].ToString()),
                                            StartTime = _ChangeTimeFormat(checkNoword(SqlData["StartTime"].ToString())),
                                            EndTime = _ChangeTimeFormat(checkNoword(SqlData["EndTime"].ToString())),
                                            AssignDate_PM = _checkNoword(DateTime.Parse(SqlData["AssignDate_PM"].ToString()).ToString("yyyy-MM-dd")),
                                            AssignDate = _checkNoword(DateTime.Parse(SqlData["AssignDate"].ToString()).ToString("yyyy-MM-dd")),
                                            Progress = ChangeIntFormat(SqlData["Progress"].ToString()),
                                            Note = _checkNoword(SqlData["Note"].ToString()),
                                            CustomerName = string.IsNullOrEmpty(customer[1]) ? "-" : customer[1],
                                            CustomerLocation = string.IsNullOrEmpty(new Address(string.IsNullOrEmpty(customer[2]) ? "-" : customer[2]).City) ? "-" : new Address(string.IsNullOrEmpty(customer[2]) ? "-" : customer[2]).City,
                                            Important = Convert.ToBoolean(SqlData["Important"].ToString())
                                            //Assign = Convert.ToBoolean(SqlData["Scheduled"].ToString())
                                        });



                                        if (result[i].AssignDate != "－")
                                        {
                                            var days = Delayday(result[i].AssignDate, DateTime.Now.ToString());
                                            if (days > 0)
                                            {
                                                result[i].ProcessStatus = "DELAY";
                                                result[i].DelayDays = Convert.ToInt32(Math.Round(days, 0, MidpointRounding.AwayFromZero).ToString());
                                                //result[i].OrderHasDelayed = true;
                                            }
                                            else
                                            {
                                                if (Delayday(DateTime.Now.ToString(), result[i].AssignDate_PM) >= 0 && Delayday(DateTime.Now.ToString(), result[i].AssignDate_PM) <= 5)
                                                {
                                                    result[i].ProcessStatus = "WILLDELAY";
                                                    //result[i].AssignDelayIsComing = true;
                                                }
                                                else
                                                {
                                                    result[i].ProcessStatus = "NORMAL";

                                                }
                                                result[i].DelayDays = 0;
                                            }
                                        }
                                        else
                                        {
                                            result[i].DelayDays = 0;
                                        }

                                        if (result[i].StartTime == "－")
                                        {
                                            result[i].Assign = false;
                                        }
                                        else
                                        {
                                            result[i].Assign = true;
                                        }
                                        i += 1;
                                    }
                                }
                            }
                        }
                    }
                }
                catch (SqlException sqlex)
                {
                    result.Add(new Schedule
                    {
                        MAKTX = sqlex.ToString()
                    });

                }
            }


            return new ActionResponse<List<Schedule>>
            {
                Data = result
            };
        }


        /// <summary>
        /// 取得 Dashboard 表頭順序 (O)*
        /// </summary>
        /// <returns></returns>
        [HttpGet("DashboardSeq")]
        public ActionResponse<List<Dashboard>> DashboardSeq()
        {
            var result = new List<Dashboard>();
            var SqlStr = @"SELECT *
                           FROM Dashboard
                           where page=1
                           ORDER BY Sequence";

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
                                result.Add(new Dashboard
                                {
                                    //ID = Convert.ToInt16(SqlData["ID"]),
                                    Name = checkNoword(SqlData["Name"].ToString()),
                                    Sequence = Convert.ToInt16(SqlData["Sequence"]),
                                    Key = checkNoword(SqlData["EngName"].ToString()),
                                    //CanBeSort = false,
                                    //Freeze = false,
                                    //TableWidth = "170",
                                    DateType = checkNoword(SqlData["DataType"].ToString().Trim()),
                                    IsShow = true,
                                    IsAdvancedSearch = true
                                });

                            }
                        }
                    }
                }
            }

            for (int i = 0; i < 3; i++)
            {
                result[i].Freeze = true;
            }

            return new ActionResponse<List<Dashboard>>
            {
                Data = result.OrderBy(x => x.Sequence).ToList()
            };
        }

        /// <summary>
        /// 更新 Dashboard 順序 (O)*
        /// </summary>
        /// <param name="DashboardSeq">請輸入欲更改欄位之中文名稱與順序</param>
        /// <returns></returns>
        [HttpPost("DashboardSeqUpdate")]
        public ActionResponse<string> DashboardSeqUpdate([FromBody] List<DashboardUpdate> DashboardSeq)
        {
            string result = "";
            int EffectRow = 0;
            if (DashboardSeq.Count == 0 || DashboardSeq == null)
            {
                result = "Not Change";
            }
            else
            {
                try
                {
                    string SqlStr = "";
                    if (DashboardSeq.Distinct(x => x.Sequence).ToList().Count() != DashboardSeq.Count || DashboardSeq.Distinct(x => x.Name).ToList().Count() != DashboardSeq.Count)
                    {
                        result = "Can not be Updated";
                    }
                    else
                    {
                        for (int i = 0; i < DashboardSeq.Count; i++)
                        {
                            SqlStr += $"UPDATE Dashboard SET Sequence = {DashboardSeq[i].Sequence} WHERE Name = '{DashboardSeq[i].Name}' and Page = 1" + Environment.NewLine;
                        }
                        if (SqlStr != "")
                        {
                            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                            {
                                if (conn.State != ConnectionState.Open)
                                    conn.Open();
                                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                                {
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
                        }
                    }

                }
                catch
                {
                    result = "Update Failed";
                }
            }
            return new ActionResponse<string>
            {
                Data = result
            };
        }


        /// <summary>
        /// 顯示一張工單製程詳細資訊 (O)
        /// </summary>
        /// <param name="request">工單編號、製程編號</param>
        /// <returns></returns>
        [HttpPost("OPdetail")]
        public ActionResponse<List<DetailData>> OPdetail([FromBody] ScheduleOneOrderRequest request)
        {
            var result = new List<DetailData>();
            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.OrderQTY,w.QtyTol, a.StartTime, a.EndTime, a.AssignDate_PM, a.AssignDate, (CAST(w.QtyTol AS float)/CAST(w.OrderQTY AS float)*100) as precent 
                            FROM Assignment as a 
                            LEFT JOIN WIP as w ON a.OrderID=w.OrderID AND a.OPID=w.OPID
                            WHERE a.OrderID=@OrderID AND a.OPID=@OPID 
                            ORDER BY  a.StartTime ASC";
            if (request.OPID != "" && request.OrderID != "")
            {
                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(SqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                        {
                            conn.Open();
                        }
                        comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = request.OrderID;
                        comm.Parameters.Add("@OPID", SqlDbType.Float).Value = request.OPID;
                        using (SqlDataReader SqlData = comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                double WT_sum = 0.0;
                                double CT_sum = 0.0;
                                string temp_WT_end = "";
                                string temp_WT_start = "";
                                string temp_CT_start = "";
                                string temp_CT_end = "";
                                double DDay = 0;

                                int i = 0;
                                while (SqlData.Read())
                                {
                                    if (i == 0)
                                    {
                                        temp_WT_end = SqlData["EndTime"].ToString();
                                        temp_CT_start = SqlData["StartTime"].ToString();
                                    }
                                    else if (i >= 1)
                                    {
                                        temp_WT_start = SqlData["StartTime"].ToString();
                                        WT_sum += Costtime(temp_WT_end, temp_WT_start);
                                        temp_WT_end = SqlData["EndTime"].ToString();
                                    }

                                    if (SqlData["OPID"].ToString() == request.OPID)
                                    {
                                        if (Delayday(SqlData["AssignDate"].ToString(), DateTime.Now.ToString()) > 0)
                                        {
                                            DDay = Delayday(SqlData["AssignDate"].ToString(), DateTime.Now.ToString());
                                        }
                                        else
                                        {
                                            DDay = 0;
                                        }
                                        //DateTime st = DateTime.Parse(SqlData["StartTime"].ToString());
                                        //DateTime et = DateTime.Parse(SqlData["EndTime"].ToString());
                                        //DateTime asssign_pm = DateTime.Parse(SqlData["AssignDate_PM"].ToString());
                                        //DateTime assign = DateTime.Parse(SqlData["AssignDate"].ToString());
                                        //var ttt = SqlData["precent"].ToString();

                                        result.Add(new DetailData
                                        {
                                            OrderID = (SqlData["OrderID"].ToString()).Trim(),
                                            Total = string.IsNullOrEmpty(SqlData["OrderQTY"].ToString()) ? 0 : Convert.ToInt32(SqlData["OrderQTY"].ToString()),
                                            PredictST = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "－" : Convert.ToString(SqlData["StartTime"].ToString()),
                                            PredictET = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "－" : Convert.ToString(SqlData["EndTime"].ToString()),
                                            AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "－" : Convert.ToString(SqlData["AssignDate_PM"].ToString()),
                                            AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "－" : Convert.ToString(SqlData["AssignDate"].ToString()),
                                            //Progress = ChangeProgressIntFormat(SqlData["precent"].ToString()),
                                            Progress = OrderProcessRate(request.OrderID).ToString(),
                                            DelayDay = int.Parse(Math.Round(DDay, 0, MidpointRounding.AwayFromZero).ToString()),
                                        });
                                    }
                                    temp_CT_end = SqlData["EndTime"].ToString();
                                    i++;
                                }

                                if (result.Count > 0)
                                {
                                    CT_sum = Costtime(temp_CT_start, temp_CT_end);
                                    result[0].WT = WT_sum;
                                    result[0].CT = CT_sum;
                                }
                                else
                                {
                                    result.Add(new DetailData());
                                }

                                foreach (var item in result)
                                {
                                    if (item.PredictET != "－")
                                    {
                                        item.PredictET = DateTime.Parse(item.PredictET).ToString("yyyy-MM-dd HH:mm");
                                    }
                                    if (item.PredictST != "－")
                                    {
                                        item.PredictST = DateTime.Parse(item.PredictST).ToString("yyyy-MM-dd HH:mm");
                                    }
                                    if (item.AssignDate != "－")
                                    {
                                        item.AssignDate = DateTime.Parse(item.AssignDate).ToString("yyyy-MM-dd");
                                    }
                                    if (item.AssignDate_PM != "－")
                                    {
                                        item.AssignDate_PM = DateTime.Parse(item.AssignDate_PM).ToString("yyyy-MM-dd");
                                    }
                                }


                            }
                        }
                    }
                }
            }
            return new ActionResponse<List<DetailData>>
            {
                Data = result
            };
        }

        private TimeSpan getHumanOpTime(SqlDataReader data)
        {
            int QTY = string.IsNullOrEmpty(data["OrderQTY"].ToString()) ? 0 : Convert.ToInt32(data["OrderQTY"].ToString());
            double Ht = string.IsNullOrEmpty(data["HumanOpTime"].ToString()) ? 0 : Convert.ToDouble(data["HumanOpTime"].ToString()); ;
            TimeSpan result = TimeSpan.FromMinutes(Ht * QTY);
            return result;
        }

        private TimeSpan getMachOpTime(SqlDataReader data)
        {
            int QTY = string.IsNullOrEmpty(data["OrderQTY"].ToString()) ? 0 : Convert.ToInt32(data["OrderQTY"].ToString());
            double Mt = string.IsNullOrEmpty(data["MachOpTime"].ToString()) ? 0 : Convert.ToDouble(data["MachOpTime"].ToString()); ;
            TimeSpan result = TimeSpan.FromMinutes(Mt * QTY);
            return result;
        }

        //以工單為單位取得整張工單的進度(個別製程進度取平均)
        private int OrderProcessRate(string orderid)
        {
            int result = 0;
            List<int> data = new List<int>();
            var SqlStr = $@"select ROUND(cast(cast(b.QtyTol as float) / cast(b.OrderQTY as float) * 100 as decimal), 0) as precent
                            from Assignment as a
                            left join WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                            where a.OrderID=@OrderID AND b.QtyTol is not null AND b.OrderQTY is not null";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = orderid;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        while (SqlData.Read())
                        {
                            if (SqlData.HasRows)
                            {
                                data.Add(int.Parse(SqlData["precent"].ToString()));
                            }
                        }
                    }
                }
            }
            if (data.Count() > 0)
            {
                result = int.Parse(Math.Round(data.Average(), 0, MidpointRounding.AwayFromZero).ToString());
            }
            //result = int.Parse(Math.Round(data.Average(), 0, MidpointRounding.AwayFromZero).ToString());
            return result;
        }

        /// <summary>
        /// 計算延遲天數(今日-預交日)
        /// </summary>
        /// <param name="time1"></param>
        /// <param name="time2"></param>
        /// <returns></returns>         
        private double Delayday(string time1, string time2)
        {
            DateTime date1 = DateTime.Parse(time1);
            DateTime date2 = DateTime.Parse(time2);
            double days = Math.Round(
                new TimeSpan(date2.Ticks - date1.Ticks).TotalDays
                , 0, MidpointRounding.AwayFromZero);
            return days;
        }

        /// <summary>
        /// 計算花費時數(用於OPdetail->hour)
        /// </summary>
        /// <param name="time1"></param>
        /// <param name="time2"></param>
        /// <returns></returns>         
        private double Costtime(string time1, string time2)
        {
            if (time1 == "" || time2 == "")
            {
                return 0;
            }
            DateTime date1 = DateTime.Parse(time1);
            DateTime date2 = DateTime.Parse(time2);
            double hours = Math.Round(
                new TimeSpan(date2.Ticks - date1.Ticks).TotalHours
                , 0, MidpointRounding.AwayFromZero);
            if (hours < 0) hours = 0;
            return hours;
        }

        /// <summary>
        /// 工單延遲與提早訊息 (O)
        /// </summary>
        /// <returns></returns>
        [HttpGet("ListWorkCondition")]
        public ActionResponse<List<ListWorkConditions>> ListWorkCondition()
        {
            var result = new List<ListWorkConditions>();
            List<OP_Data> OP_Datas = new List<OP_Data>();
            int i;
            var SqlStr = $@"SELECT TOP(10) w.OrderID, w.OPID, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST, w.EndTime as Real_ET, w.WorkGroup, a.Important AS Important
                             FROM WIP as w
                             LEFT JOIN Assignment as a ON a.OrderID=w.OrderID AND a.OPID=w.OPID
                             WHERE WIPEvent='3'
                             ORDER BY w.WorkGroup ASC, Real_ET DESC";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {

                            while (SqlData.Read())
                            {
                                OP_Datas.Add(new OP_Data
                                {
                                    OrderID = SqlData["OrderID"].ToString(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Assign_ST = SqlData["Assign_ST"].ToString(),
                                    Assign_ET = SqlData["Assign_ET"].ToString(),
                                    Real_ST = SqlData["Real_ST"].ToString(),
                                    Real_ET = SqlData["Real_ET"].ToString(),
                                    Workgroup = SqlData["WorkGroup"].ToString()
                                });
                            }
                        }
                    }
                }
            }

            double Timediff = 0.0;

            for (i = 0; i < OP_Datas.Count(); i++)
            {
                if (OP_Datas[i].Real_ET == "" ||
                    OP_Datas[i].Real_ST == "" ||
                    OP_Datas[i].Assign_ET == "" ||
                    OP_Datas[i].Assign_ST == "")
                    continue;

                if (i > 0)
                {
                    if (OP_Datas[i].Workgroup == OP_Datas[i - 1].Workgroup)
                    {
                        continue;
                    }
                }


                Timediff = TimeDiff(OP_Datas[i].Assign_ET, OP_Datas[i].Real_ET);

                if (Timediff > 0) //延遲
                {
                    var Next_op = WorkInfo(OP_Datas[i].Workgroup, OP_Datas[i].Assign_ET);
                    result.Add(new ListWorkConditions
                    {
                        Endtime = DateTime.Parse(OP_Datas[i].Real_ET).ToString("yyyy-MM-dd HH:mm"),
                        OrderID = Next_op[0].OrderID,
                        OPID = Next_op[0].OPID,
                        Message = "延後開始時間，請確認製程執行方式",
                        IsOPDeley = true
                    }); ;
                }
                else if (Timediff < 0)//提早
                {
                    result.Add(new ListWorkConditions
                    {
                        Endtime = DateTime.Parse(OP_Datas[i].Real_ET).ToString("yyyy-MM-dd HH:mm"),
                        OrderID = OP_Datas[i].OrderID.Trim(),
                        OPID = OP_Datas[i].OPID,
                        Message = "提早完成",
                        IsOPDeley = false
                    });
                }
            }

            return new ActionResponse<List<ListWorkConditions>>
            {
                Data = result
            };
        }

        private double TimeDiff(string time1, string time2)
        {
            double minutes = 0;
            if (IsDate(time1) && IsDate(time2))
            {
                DateTime date1 = DateTime.Parse(time1);
                DateTime date2 = DateTime.Parse(time2);
                minutes = Math.Round(
                    new TimeSpan(date2.Ticks - date1.Ticks).TotalMinutes
                    , 0, MidpointRounding.AwayFromZero);
            }
            return minutes;
        }


        private bool IsDate(string strDate)
        {
            try
            {
                DateTime.Parse(strDate);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private List<WorkInfos> WorkInfo(string Workgroup, string Endtime)
        {
            var result = new List<WorkInfos>();

            var SqlStr = $@"SELECT TOP(1) a.OrderID, a.OPID, a.WorkGroup, a.StartTime, a.EndTime
                            FROM　Assignment as a
                            LEFT JOIN WIP as w
                            ON a.OrderID=w.OrderID AND a.OPID=w.OPID
                            WHERE a.WorkGroup=@Workgroup AND a.StartTime>=@Endtime AND w.WIPEvent!='3' 
                            ORDER BY a.StartTime ASC";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    DateTime et = DateTime.Parse(Endtime);
                    comm.Parameters.Add("@Workgroup", SqlDbType.NVarChar).Value = Workgroup;
                    comm.Parameters.Add("@Endtime", SqlDbType.NVarChar).Value = et.ToString("yyyy-MM-dd HH:mm");

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {

                            while (SqlData.Read())
                            {
                                result.Add(new WorkInfos
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),
                                });
                            }
                        }
                        else
                        {
                            result.Add(new WorkInfos
                            {
                                OrderID = "",
                                OPID = "",
                            });
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 工作上班時間資訊 (O)
        /// </summary>
        /// <returns></returns>
        [HttpGet("WorkingHours")]
        public ActionResponse<CommuterTime> WorkingHours()
        {
            var result = new CommuterTime();
            result.StartTime = "08:00";
            result.CloseTime = "17:00";
            return new ActionResponse<CommuterTime>
            {
                Data = result
            };
        }

        /// <summary>
        /// 甘特圖排程模式標籤 (O)*
        /// </summary>
        /// <returns></returns>
        [HttpGet("ScheduleItems")]
        public ActionResponse<List<chartTab>> ScheduleItems()
        {
            //判斷temp1~7是否有數據
            var sqlStr = "";

            List<int> templist = new List<int>();
            for (int i = 1; i <= 7; i++)
            {
                var Assign = "AssignmentTemp" + i.ToString();
                sqlStr = $@"SELECT * FROM {Assign}";
                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(sqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                        {
                            conn.Open();
                        }
                        using (SqlDataReader SqlData = comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                templist.Add(i);
                            }
                        }
                    }
                }
            }

            var result = new List<chartTab>();

            result.Add(new chartTab
            {
                ModeCode = "0",
                DisplayTitle = "原始排程",
            });

            foreach (var item in templist)
            {
                if (item == 1)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "1",
                        DisplayTitle = "交期優先-系統權重",
                    });
                }

                if (item == 2)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "2",
                        DisplayTitle = "交期優先-自訂權重",
                    });
                }

                if (item == 3)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "3",
                        DisplayTitle = "機台優先",
                    });
                }

                if (item == 4)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "4",
                        DisplayTitle = "插單優先",
                    });
                }

                if (item == 5)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "5",
                        DisplayTitle = "手動優先",
                    });
                }

                if (item == 6)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "6",
                        DisplayTitle = "設備故障排程重排",
                    });
                }

                if (item == 7)
                {
                    result.Add(new chartTab
                    {
                        ModeCode = "7",
                        DisplayTitle = "設備故障排程延後工單右移",
                    });
                }
            }

            return new ActionResponse<List<chartTab>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 工單製程目前狀態(甘特圖基本狀態) (O)
        /// </summary>
        /// <returns></returns>
        [HttpGet("BasicChartWorkCondition")]
        public ActionResponse<List<BasicChartWorkConditions>> BasicChartWorkCondition()
        {
            var result = new List<BasicChartWorkConditions>();
            List<BasicChartOPData> Chart_OP_Datas = new List<BasicChartOPData>();
            string groupNumber = "ALL";
            string groupsearch = $"and d.GroupName='{groupNumber}'";
            if (groupNumber == "ALL")
                groupsearch = "";

            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.OPLTXA1, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                            w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,
                            Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0)
                            ,d.GroupName
                            FROM  {_ConnectStr.APSDB}.[dbo].[Assignment] as a 
                            inner join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                            LEFT JOIN  {_ConnectStr.APSDB}.[dbo].[WIP] as w ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                            left join {_ConnectStr.APSDB}.[dbo].Outsourcing as o on d.ID=o.Id
                            where a.Scheduled = 1 {groupsearch} and a.StartTime is not NULL and a.AssignDate>DATEADD(DAY,-60,GETDate()) and a.StartTime>='2024-04-01 00:00:00'" +
                            //"and o.Outsource=0"+
                            @" ORDER BY w.WorkGroup ASC, Assign_ST ASC";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {

                            while (SqlData.Read())
                            {
                                Chart_OP_Datas.Add(new BasicChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),//SqlData["OPID"].ToString()+"_"+ SqlData["OPLTXA1"].ToString(),
                                    Workgroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "N/A" : SqlData["WorkGroup"].ToString(),
                                    GroupNumber = string.IsNullOrEmpty(SqlData["GroupName"].ToString()) ? "N/A" : SqlData["GroupName"].ToString(),
                                    AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString()),
                                    AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString()),
                                    Real_ST = SqlData["Real_ST"].ToString(),
                                    Real_ET = SqlData["Real_ET"].ToString(),
                                    PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString(_timeFormat),
                                    ASDate = SqlData["ASDate"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                    Priority = SqlData["PRIORITY"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString())
                                });
                            }
                        }
                    }
                }
            }

            //刪除過時的完工資料
            var overOrders = getOverOrders();
            for (int i = 0; i < overOrders.Count; i++)
            {
                Chart_OP_Datas.Remove(overOrders[i]);
            }

            var tempResult = Chart_OP_Datas;
            var CData = new List<BasicChartWorkConditions>();//甘特圖圖表資料
            //取得各機台是否為委外機台
            int chromosomesCount = _ConnectStr.chromosomesCount;
            var devices = DeviceInfo();
            var PMC = new DelayMthod(chromosomesCount, devices);
            var OutsourcingList = PMC.getOutsourcings();
            var UnlimitMC = OutsourcingList.Where(x => x.isOutsource == "1").Select(x=>x.remark).ToList();
            if(tempResult.Count>0)
            {
                foreach (var item in tempResult.Where(x=> !UnlimitMC.Exists(y => y == x.Workgroup)))
                {
                    var temp = CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                    if (temp is null)
                    {
                        CData.Add(new BasicChartWorkConditions
                        {
                            GroupID = "",
                            AssignST = item.AssignST.ToString(_timeFormat_ToMin),
                            AssignET = item.AssignET.ToString(_timeFormat_ToMin),
                            OPID = item.OPID,
                            OrderID = item.OrderID,
                            OPSTATE = CheckOPState(item),
                            DueDate = item.PMDate,
                            WorkGroup = item.Workgroup,
                            GroupNumber = item.GroupNumber,
                            OPICON = CheckOpICON(item, "NP"),
                            actExpCompare = checkActExptatus(item, "NP"),
                            Progress = item.Progress,
                            isLock = false
                        }); 
                    }
                    else
                    {
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).AssignST = item.AssignST.ToString(_timeFormat_ToMin);
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).AssignET = item.AssignET.ToString(_timeFormat_ToMin);
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).WorkGroup = item.Workgroup;
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPSTATE = "UNSTARTED";
                    }
                    CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPICON.IsUnconfirmed = false;
                    CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPICON.IsAffected = false;
                }
                int groupID = 1;
                foreach (var item in UnlimitMC)
                {
                    var specialMC = tempResult.Where(x => x.Workgroup == item).ToList().OrderBy(x => x.AssignST).ToList();
                    if (specialMC.Count > 0)
                    {
                        var lastopET = specialMC[0].AssignET;
                        for (int i = 0; i < specialMC.Count(); i++)
                        {
                            if (specialMC[i].AssignST <= lastopET)
                            {
                                if (specialMC[i].AssignET > lastopET)
                                {
                                    lastopET = specialMC[i].AssignET;
                                }

                            }
                            else
                            {
                                lastopET = specialMC[i].AssignET;
                                groupID += 1;
                            }

                            var temp = CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID);
                            if (temp is null)
                            {
                                CData.Add(new BasicChartWorkConditions
                                {
                                    GroupID = groupID.ToString(),
                                    AssignST = specialMC[i].AssignST.ToString(_timeFormat_ToMin),
                                    AssignET = specialMC[i].AssignET.ToString(_timeFormat_ToMin),
                                    OPID = specialMC[i].OPID,
                                    OrderID = specialMC[i].OrderID,
                                    OPSTATE = CheckOPState(specialMC[i]),
                                    DueDate = specialMC[i].PMDate,
                                    WorkGroup = specialMC[i].Workgroup,
                                    GroupNumber = specialMC[i].GroupNumber,
                                    OPICON = CheckOpICON(specialMC[i], "NP"),
                                    actExpCompare = checkActExptatus(specialMC[i], "NP"),
                                    Progress = specialMC[i].Progress,
                                    isLock = true
                                });
                            }
                            else
                            {
                                CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID).AssignST = specialMC[i].AssignST.ToString(_timeFormat_ToMin);
                                CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID).AssignET = specialMC[i].AssignET.ToString(_timeFormat_ToMin);
                                CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID).WorkGroup = specialMC[i].Workgroup;
                                CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID).OPSTATE = "UNSTARTED";
                            }
                            CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID).OPICON.IsUnconfirmed = false;
                            CData.Find(x => x.OrderID == specialMC[i].OrderID && x.OPID == specialMC[i].OPID).OPICON.IsAffected = false;
                        }
                    }
                    groupID += 1;

                }
            }
            

            result = CData;
            return new ActionResponse<List<BasicChartWorkConditions>>
            {
                Data = result
            };
        }

        private List<BasicChartOPData> getOverOrders()
        {
            var result = new List<BasicChartOPData>();

            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.OPLTXA1, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                            w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,
                            Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0)
                            ,d.GroupName
                            FROM  {_ConnectStr.APSDB}.[dbo].[Assignment] as a 
                            inner join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                            LEFT JOIN  {_ConnectStr.APSDB}.[dbo].[WIP] as w
                            ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                            where w.WIPEvent=3 and w.UpdateTime > GETDATE()-10
                            ORDER BY w.WorkGroup ASC, Assign_ST ASC";
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
                                result.Add(new BasicChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),//SqlData["OPID"].ToString()+"_"+ SqlData["OPLTXA1"].ToString(),
                                    Workgroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "N/A" : SqlData["WorkGroup"].ToString(),
                                    GroupNumber = string.IsNullOrEmpty(SqlData["GroupName"].ToString()) ? "N/A" : SqlData["GroupName"].ToString(),
                                    AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString()),
                                    AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString()),
                                    Real_ST = SqlData["Real_ST"].ToString(),
                                    Real_ET = SqlData["Real_ET"].ToString(),
                                    PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString(_timeFormat),
                                    ASDate = SqlData["ASDate"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                    Priority = SqlData["PRIORITY"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString())
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// 工單製程甘特圖基本狀態預覽 (O)*
        /// </summary>
        /// <param name="mode">【0:原始排程、1:交期優先-推薦權重、2:交期優先-自訂權重、3:機台優先、4:插單優先、5:手動排程、6:設備故障排程重排、7:設備故障排程延後工單右移】</param>
        /// <returns></returns>
        [HttpGet("BasicChartPreview/{mode}")]
        public ActionResponse<List<BasicChartWorkConditions>> BasicChartPreview(string mode)
        {
            List<string> Mode_List = new List<string> { "0", "1", "2", "3", "4", "5", "6", "7" };
            var result = new List<BasicChartWorkConditions>();
            var changeds = new List<int>();
            if (Mode_List.Exists(x => x == mode))
            {
                string Assign = string.Empty;

                if (mode == "0")
                {
                    Assign = "Assignment";
                }
                else
                {
                    Assign = "AssignmentTemp" + mode;
                }
                List<BasicChartOPData> Chart_OP_Datas = new List<BasicChartOPData>();
                string groupNumber = "ALL";
                string groupsearch = $"and d.GroupName='{groupNumber}'";
                if (groupNumber == "ALL")
                    groupsearch = "";

                var SqlStr = $@"SELECT a.OrderID, a.OPID, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                                w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,                                
                                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0), a.Scheduled
                                ,d.GroupName
                                FROM  {_ConnectStr.APSDB}.[dbo].{Assign} as a 
                                left join {_ConnectStr.APSDB}.[dbo].Device as d on a.WorkGroup=d.remark
                                left JOIN  {_ConnectStr.APSDB}.[dbo].[WIP] as w ON w.OrderID=a.OrderID AND w.OPID=a.OPID AND w.SeriesID=a.SeriesID
                                where (a.Scheduled = 1 or a.Scheduled = 2 or a.Scheduled = 3) {groupsearch} and a.StartTime is not NULL and a.AssignDate>DATEADD(DAY,-60,GETDate()) and a.StartTime>='2024-04-01 00:00:00'
                                ORDER BY w.WorkGroup ASC, Assign_ST ASC";


                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(SqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                        {
                            conn.Open();
                        }

                        using (SqlDataReader SqlData = comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {

                                while (SqlData.Read())
                                {
                                    Chart_OP_Datas.Add(new BasicChartOPData
                                    {
                                        OrderID = SqlData["OrderID"].ToString().Trim(),
                                        OPID = SqlData["OPID"].ToString(),
                                        Workgroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "N/A" : SqlData["WorkGroup"].ToString(),
                                        GroupNumber = string.IsNullOrEmpty(SqlData["GroupName"].ToString()) ? "N/A" : SqlData["GroupName"].ToString(),
                                        AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString(), null),
                                        AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString(), null),
                                        Real_ST = string.IsNullOrEmpty(SqlData["Real_ST"].ToString()) ? string.Empty : SqlData["Real_ST"].ToString(),
                                        Real_ET = string.IsNullOrEmpty(SqlData["Real_ET"].ToString()) ? string.Empty : SqlData["Real_ET"].ToString(),
                                        PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString(_timeFormat),
                                        ASDate = SqlData["ASDate"].ToString(),
                                        WIPEvent = SqlData["WIPEvent"].ToString(),
                                        Priority = SqlData["PRIORITY"].ToString(),
                                        Important = SqlData["Important"].ToString(),
                                        Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString())
                                    });
                                    changeds.Add(Convert.ToInt32(SqlData["Scheduled"].ToString()));
                                }
                            }
                        }
                    }
                }

                var CData = new List<BasicChartWorkConditions>();//甘特圖圖表資料
                //取得各機台是否為委外機台
                int chromosomesCount = _ConnectStr.chromosomesCount;
                var devices = DeviceInfo();
                var PMC = new DelayMthod(chromosomesCount, devices);
                var OutsourcingList = PMC.getOutsourcings();
                var UnlimitMC = OutsourcingList.Where(x => x.isOutsource == "1").Select(x => x.remark).ToList();
                int i = 0;
                foreach (var item in Chart_OP_Datas.Where(x => !UnlimitMC.Exists(y => y == x.Workgroup)))
                {
                    var temp = CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                    if (temp is null)
                    {
                        CData.Add(new BasicChartWorkConditions
                        {
                            GroupID = "",
                            AssignST = item.AssignST.ToString(_timeFormat_ToMin),
                            AssignET = item.AssignET.ToString(_timeFormat_ToMin),
                            OPID = item.OPID,
                            OrderID = item.OrderID,
                            OPSTATE = CheckOPState(item),
                            WorkGroup = item.Workgroup,
                            GroupNumber = item.GroupNumber,
                            DueDate = item.PMDate,
                            OPICON = CheckOpICON(item, "P"),
                            actExpCompare = checkActExptatus(item, "P"),
                            Progress = item.Progress,
                            isLock = false
                        });
                    }
                    else
                    {
                        //var findidx = CData.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                        //CData[findidx].AssignST = item.AssignST.ToString(_timeFormat_ToMin);
                        //CData[findidx].AssignET = item.AssignET.ToString(_timeFormat_ToMin);
                        //CData[findidx].WorkGroup = item.Workgroup;
                        //CData[findidx].OPSTATE = "UNSTARTED";
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).AssignST = item.AssignST.ToString(_timeFormat_ToMin);
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).AssignET = item.AssignET.ToString(_timeFormat_ToMin);
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).WorkGroup = item.Workgroup;
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPSTATE = "UNSTARTED";
                    }
                    
                    
                    var idx = Chart_OP_Datas.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);

                    if (changeds[idx] == 2)
                    {
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPICON.IsAffected = true;
                        //temp.OPICON.IsAffected = true;
                    }
                    else if (changeds[idx] == 3)
                    {
                        CData.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPICON.IsUnconfirmed = true;
                        //temp.OPICON.IsUnconfirmed = true;
                    }
                    i += 1;
                    //CData.Add(temp);
                }
                int groupID = 1;
                foreach (var item in UnlimitMC)
                {
                    var specialMC = Chart_OP_Datas.Where(x => x.Workgroup == item).ToList().OrderBy(x => x.AssignST).ToList();
                    if (specialMC.Count > 0)
                    {
                        var lastopET = specialMC[0].AssignET;
                        for (int j = 0; j < specialMC.Count(); j++)
                        {
                            if (specialMC[j].AssignST <= lastopET)
                            {
                                if (specialMC[j].AssignET > lastopET)
                                {
                                    lastopET = specialMC[j].AssignET;
                                }

                            }
                            else
                            {
                                lastopET = specialMC[j].AssignET;
                                groupID += 1;
                            }


                            var temp = CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID);
                            if (temp is null)
                            {
                                CData.Add(new BasicChartWorkConditions
                                {
                                    GroupID = groupID.ToString(),
                                    AssignST = specialMC[j].AssignST.ToString(_timeFormat_ToMin),
                                    AssignET = specialMC[j].AssignET.ToString(_timeFormat_ToMin),
                                    OPID = specialMC[j].OPID,
                                    OrderID = specialMC[j].OrderID,
                                    OPSTATE = CheckOPState(specialMC[j]),
                                    DueDate = specialMC[j].PMDate,
                                    WorkGroup = specialMC[j].Workgroup,
                                    GroupNumber = specialMC[j].GroupNumber,
                                    OPICON = CheckOpICON(specialMC[j], "P"),
                                    actExpCompare = checkActExptatus(specialMC[j], "P"),
                                    Progress = specialMC[j].Progress,
                                    isLock = true
                                });
                            }
                            else
                            {
                                CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID).AssignST = specialMC[j].AssignST.ToString(_timeFormat_ToMin);
                                CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID).AssignET = specialMC[j].AssignET.ToString(_timeFormat_ToMin);
                                CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID).WorkGroup = specialMC[j].Workgroup;
                                CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID).OPSTATE = "UNSTARTED";
                            }


                            //var source = new BasicChartWorkConditions
                            //{
                            //    GroupID = groupID.ToString(),
                            //    AssignST = specialMC[j].AssignST.ToString(_timeFormat_ToMin),
                            //    AssignET = specialMC[j].AssignET.ToString(_timeFormat_ToMin),
                            //    OPID = specialMC[j].OPID,
                            //    OrderID = specialMC[j].OrderID,
                            //    OPSTATE = CheckOPState(specialMC[j]),
                            //    DueDate = specialMC[j].PMDate,
                            //    WorkGroup = specialMC[j].Workgroup,
                            //    GroupNumber = specialMC[j].GroupNumber,
                            //    OPICON = CheckOpICON(specialMC[j], "NP"),
                            //    actExpCompare = checkActExptatus(specialMC[j], "NP"),
                            //    Progress = specialMC[j].Progress,
                            //    isLock = true
                            //};

                            var idx = Chart_OP_Datas.FindIndex(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID);
                            if (changeds[idx] == 2)
                            {
                                CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID).OPICON.IsAffected = true;
                                //temp.OPICON.IsAffected = true;
                            }
                            else if (changeds[idx] == 3)
                            {
                                CData.Find(x => x.OrderID == specialMC[j].OrderID && x.OPID == specialMC[j].OPID).OPICON.IsUnconfirmed = true;
                                //temp.OPICON.IsUnconfirmed = true;
                            }
                            i += 1;
                            //CData.Add(temp);


                        }
                    }
                    groupID += 1;
                }
                result = CData;
            }

            return new ActionResponse<List<BasicChartWorkConditions>>
            {
                Data = result
            };
        }

        public class ChartWorkConditionsrequest
        {
            public string mode { get; set; }
            public string groupNumber { get; set; }
        }

        private int CheckPcriority(string data)
        {
            int ans = 0;
            switch (data.Trim())
            {
                case "4":
                    ans = 1;
                    break;
                case "3":
                    ans = 2;
                    break;
                case "2":
                    ans = 3;
                    break;
            }
            return ans;
        }

        private string OrderInfo(string OrderID)
        {
            var result = "";
            var SqlStr = $@"SELECT TOP(1) a.OrderID, a.OPID, a.AssignDate_PM as PMDate, a.AssignDate as ASDate
                            FROM  [Assignment] as a
                            WHERE a.OrderID=@OrderID AND a.WorkGroup is not null
                            ORDER BY a.OPID DESC";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = OrderID;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {

                                result = SqlData["OPID"].ToString();
                            }
                        }
                    }
                }
            }
            return result;
        }

        private ActExpCompare checkActExptatus(BasicChartOPData item, string version)
        {
            var result = new ActExpCompare();
            switch (item.WIPEvent.Trim())
            {
                //製程已完成
                case "3":
                    //生管預交與訂單交期確認                    
                    if (Delayday(item.ASDate, item.Real_ET) <= 0)//訂單未延遲
                    {
                        if (Delayday(item.PMDate, item.Real_ET) < 0)//生管提早
                        {
                            result.AssignPMState = 0;
                            result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                            result.PMIntervalET = string.IsNullOrEmpty(item.Real_ET) ? "N/A" : DateTime.Parse(item.Real_ET).ToString("yyyy-MM-dd HH:mm");
                        }
                        else if (Delayday(item.PMDate, item.Real_ET) > 0)//生管延遲
                        {
                            result.AssignPMState = 2;
                            result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                            result.PMIntervalET = string.IsNullOrEmpty(item.Real_ET) ? "N/A" : DateTime.Parse(item.Real_ET).ToString("yyyy-MM-dd HH:mm");
                        }
                        else
                        {
                            result.AssignPMState = 1;//生管正常
                            result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                            result.PMIntervalET = string.IsNullOrEmpty(item.Real_ET) ? "N/A" : DateTime.Parse(item.Real_ET).ToString("yyyy-MM-dd HH:mm");
                        }

                    }
                    else if (Delayday(item.ASDate, item.Real_ET) > 0)//訂單延遲
                    {
                        result.AssignPMState = 2;//生管延遲
                        result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                        result.PMIntervalET = string.IsNullOrEmpty(item.Real_ET) ? "N/A" : DateTime.Parse(item.Real_ET).ToString("yyyy-MM-dd HH:mm");
                    }

                    break;

                //製程進行中
                case "1":
                    //生管預交與訂單交期確認
                    if (version == "NP")
                    {
                        if (Delayday(item.ASDate, DateTime.Now.ToString()) <= 0)//訂單未延遲
                        {
                            if (Delayday(item.PMDate, DateTime.Now.ToString()) > 0)//生管延遲
                            {
                                result.AssignPMState = 2;
                                result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                                result.PMIntervalET = "N/A";
                            }
                            else
                            {
                                result.AssignPMState = 1;//生管正常
                                result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                                result.PMIntervalET = "N/A";

                            }

                        }
                        else if (Delayday(item.ASDate, DateTime.Now.ToString()) > 0)//訂單延遲
                        {
                            result.AssignPMState = 2;//生管延遲
                            result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                            result.PMIntervalET = "N/A";
                        }
                    }

                    else if (version == "P")
                    {
                        if (Delayday(item.ASDate, item.AssignET.ToString()) <= 0)//訂單未延遲
                        {
                            if (Delayday(item.PMDate, item.AssignET.ToString()) > 0)//生管延遲
                            {
                                result.AssignPMState = 2;
                                result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                                result.PMIntervalET = "N/A";
                            }
                            else
                            {
                                result.AssignPMState = 1;//生管正常
                                result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                                result.PMIntervalET = "N/A";

                            }

                        }
                        else if (Delayday(item.ASDate, item.AssignET.ToString()) > 0)//訂單延遲
                        {
                            result.AssignPMState = 2;//生管延遲
                            result.PMIntervalST = string.IsNullOrEmpty(item.Real_ST) ? "N/A" : DateTime.Parse(item.Real_ST).ToString("yyyy-MM-dd HH:mm");
                            result.PMIntervalET = "N/A";
                        }
                    }


                    break;

                //製程未開工
                case "0":
                    //生管預交與訂單交期確認

                    if (version == "NP")
                    {
                        if (Delayday(item.ASDate, DateTime.Now.ToString()) <= 0)//訂單未延遲
                        {
                            if (Delayday(item.PMDate, DateTime.Now.ToString()) > 0)//生管延遲
                            {
                                result.AssignPMState = 2;
                                result.PMIntervalST = "N/A";
                                result.PMIntervalET = "N/A";
                            }
                            else
                            {
                                result.AssignPMState = 1;//生管正常
                                result.PMIntervalST = "N/A";
                                result.PMIntervalET = "N/A";

                            }

                        }
                        else if (Delayday(item.ASDate, DateTime.Now.ToString()) > 0)//訂單延遲
                        {
                            //OrderHasDelayed = true;
                            result.AssignPMState = 2;//生管延遲
                            result.PMIntervalST = "N/A";
                            result.PMIntervalET = "N/A";

                        }
                    }
                    else if (version == "P")
                    {
                        //if (Delayday(item.ASDate, item.AssignET.ToString()) <= 0)//訂單未延遲
                        if (Delayday(item.ASDate, DateTime.Now.ToString()) <= 0)//訂單未延遲
                        {
                            //if (Delayday(item.PMDate, item.AssignET.ToString()) > 0)//生管延遲
                            if (Delayday(item.PMDate, DateTime.Now.ToString()) > 0)//生管延遲
                            {
                                result.AssignPMState = 2;
                                result.PMIntervalST = "N/A";
                                result.PMIntervalET = "N/A";
                            }
                            else
                            {
                                result.AssignPMState = 1;//生管正常
                                result.PMIntervalST = "N/A";
                                result.PMIntervalET = "N/A";

                            }

                        }
                        else if (Delayday(item.ASDate, item.AssignET.ToString()) > 0)//訂單延遲
                        {
                            result.AssignPMState = 2;//生管延遲
                            result.PMIntervalST = "N/A";
                            result.PMIntervalET = "N/A";

                        }
                    }

                    break;

                default:

                    result.AssignPMState = 1;//生管正常
                    result.PMIntervalST = "N/A";
                    result.PMIntervalET = "N/A";
                    break;


            }
            result.AssignPMState = 1;//生管正常
            return result;
        }

        private BasicOPstates CheckOpICON(BasicChartOPData item, string version)
        {
            var result = new BasicOPstates();
            result.ProcessStatus = "NORMAL";
            result.IsAffected = false;
            result.IsImportant = false;
            result.IsUnconfirmed = false;
            result.OrderPriority = 5;


            //是否為重要製程
            if (item.Important.Trim() == "False")
            {
                result.IsImportant = false;
            }
            else
            {
                result.IsImportant = true;
            }

            result.OrderPriority = CheckPcriority(item.Priority);

            switch (item.WIPEvent.Trim())
            {
                //製程已完成
                case "3":
                    //生管預交與訂單交期確認
                    if (Delayday(item.ASDate, item.Real_ET) <= 0)//訂單未延遲
                    {

                        if (Delayday(item.Real_ET, item.PMDate) >= 0 && Delayday(item.Real_ET, item.PMDate) <= 5)//生管延遲
                        {
                            result.ProcessStatus = "WILLDELAY";
                        }
                    }
                    else if (Delayday(item.ASDate, item.Real_ET) > 0)//訂單延遲
                    {
                        result.ProcessStatus = "DELAY";
                    }
                    break;

                //製程進行中
                case "1":
                    //生管預交與訂單交期確認
                    if (version == "NP") //非預覽模式
                    {
                        if (Delayday(item.ASDate, DateTime.Now.ToString()) <= 0)//訂單未延遲
                        {

                            if (Delayday(DateTime.Now.ToString(), item.PMDate) >= 0 && Delayday(DateTime.Now.ToString(), item.PMDate) <= 5)//生管延遲
                            {
                                result.ProcessStatus = "WILLDELAY";
                            }
                        }
                        else if (Delayday(item.ASDate, DateTime.Now.ToString()) > 0)//訂單延遲
                        {
                            result.ProcessStatus = "DELAY";
                        }
                    }
                    else if (version == "P") //預覽模式
                    {
                        if (Delayday(item.ASDate, item.AssignET.ToString()) <= 0)//訂單未延遲
                        {

                            if (Delayday(item.AssignET.ToString(), item.PMDate) >= 0 && Delayday(item.AssignET.ToString(), item.PMDate) <= 5)//生管延遲
                            {
                                result.ProcessStatus = "WILLDELAY";
                            }
                        }
                        else if (Delayday(item.ASDate, item.AssignET.ToString()) > 0)//訂單延遲
                        {
                            result.ProcessStatus = "DELAY";
                        }
                    }
                    break;

                //製程未開工
                case "0":
                    //生管預交與訂單交期確認
                    if (version == "NP") //非預覽模式
                    {
                        if (Delayday(item.ASDate, DateTime.Now.ToString()) <= 0)//訂單未延遲
                        {

                            if (Delayday(item.AssignET.ToString(), item.PMDate) >= 0 && Delayday(item.AssignET.ToString(), item.PMDate) <= 5)//生管延遲
                            {
                                result.ProcessStatus = "WILLDELAY";
                            }
                        }
                        //else if (Delayday(item.ASDate, item.AssignET.ToString()) > 0)//訂單延遲
                        else if (Delayday(item.ASDate, DateTime.Now.ToString()) > 0)//訂單延遲
                        {
                            result.ProcessStatus = "DELAY";
                        }
                    }
                    else if (version == "P") //預覽模式
                    {
                        //if (Delayday(item.ASDate, item.AssignET.ToString()) <= 0)//訂單未延遲
                        if (Delayday(item.ASDate, DateTime.Now.ToString()) <= 0)//訂單未延遲
                        {

                            //if (Delayday(item.AssignET.ToString(), item.PMDate) >= 0 && Delayday(item.AssignET.ToString(), item.PMDate) <= 5)//生管延遲
                            if (Delayday(item.AssignET.ToString(), item.PMDate) >= 0 && Delayday(item.AssignET.ToString(), item.PMDate) <= 5)//生管延遲
                            {
                                result.ProcessStatus = "WILLDELAY";
                            }
                        }
                        //else if (Delayday(item.ASDate, item.AssignET.ToString()) > 0)//訂單延遲
                        else if (Delayday(item.ASDate, DateTime.Now.ToString()) > 0)//訂單延遲
                        {
                            result.ProcessStatus = "DELAY";
                        }
                    }
                    break;
            }
            return result;

        }


        private string CheckOPState(BasicChartOPData item)
        {
            string resulr = "UNSTARTED";
            switch (item.WIPEvent.Trim())
            {
                case "0":
                    resulr = "UNSTARTED";
                    break;
                case "1":
                    resulr = "PROCESSING";
                    break;
                case "2":
                    resulr = "PAUSED";
                    break;
                case "3":
                    resulr = "FINISH";
                    break;
                case "5":
                    resulr = "BREAKDOWN";
                    break;
            }
            return resulr;
        }

        /// <summary>
        /// 工單製程目前狀態(甘特圖進度) (O)
        /// </summary>
        /// <returns>進度狀態(2:延遲 1:正常)</returns>
        [HttpGet("ProgressChartWorkCondition")]
        public ActionResponse<List<ProgressChartWorkConditions>> ProgressChartWorkCondition()
        {
            var result = new List<ProgressChartWorkConditions>();
            List<ProgressChartOPData> Chart_OP_Datas = new List<ProgressChartOPData>();
            int i;
            var SqlStr = $@"SELECT w.OrderID, w.OPID, w.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST, w.EndTime as Real_ET, w.WIPEvent
                            FROM  [WIP] as w
                            LEFT JOIN  [Assignment] as a ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                            WHERE w.WIPEvent='1' OR w.WIPEvent='3'
                            ORDER BY w.WorkGroup ASC, Assign_ST ASC";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                    {
                        conn.Open();
                    }

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {

                            while (SqlData.Read())
                            {
                                Chart_OP_Datas.Add(new ProgressChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Workgroup = SqlData["WorkGroup"].ToString(),
                                    AssignST = SqlData["Assign_ST"].ToString(),
                                    AssignET = SqlData["Assign_ET"].ToString(),
                                    Real_ST = SqlData["Real_ST"].ToString(),
                                    Real_ET = SqlData["Real_ET"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                });
                            }
                        }
                    }
                }
            }


            string MachineIDTmp = "";//機台名稱
            int ProgressState; //進度狀態(2:延遲 1:正常)
            string ProgressST;//進度延遲or正常區間
            string ProgressET;
            bool ChooseMethod;//是否要選擇執行方式


            var OP_stateTmp = new List<ProgressOPstates>();
            for (i = 0; i < Chart_OP_Datas.Count(); i++)
            {
                ProgressState = 0;
                ProgressST = "";
                ProgressET = "";
                ChooseMethod = false;

                if (i == 0)
                {

                    MachineIDTmp = Chart_OP_Datas[i].Workgroup;
                }

                //換到下一機台
                if (Chart_OP_Datas[i].Workgroup != MachineIDTmp)
                {
                    if (OP_stateTmp.Count != 0)
                    {
                        result.Add(new ProgressChartWorkConditions
                        {
                            WorkGroup = MachineIDTmp,
                            OPConditions = OP_stateTmp
                        });
                    }


                    MachineIDTmp = Chart_OP_Datas[i].Workgroup;
                    OP_stateTmp = new List<ProgressOPstates>();
                }


                switch (Chart_OP_Datas[i].WIPEvent.Trim())
                {
                    //製程已完成
                    case "3":
                        //檢查是否ST/ET都有正常數值
                        if (!string.IsNullOrEmpty(Chart_OP_Datas[i].AssignET) && !string.IsNullOrEmpty(Chart_OP_Datas[i].Real_ET))
                        {
                            //進度確認
                            if (TimeDiff(Chart_OP_Datas[i].AssignET, Chart_OP_Datas[i].Real_ET) > 0)
                            {
                                ProgressState = 2; //進度延遲
                                ProgressST = DateTime.Parse(Chart_OP_Datas[i].Real_ST).ToString(_timeFormat_ToMin);//Convert.ToDateTime(Chart_OP_Datas[i].Real_ST).ToString("yyyy-MM-dd HH:mm");
                                ProgressET = DateTime.Parse(Chart_OP_Datas[i].Real_ET).ToString(_timeFormat_ToMin);
                                ChooseMethod = true; //確認執行方式
                            }
                        }
                        break;

                    //製程進行中
                    case "1":
                        //檢查是否ST有正常數值
                        if (!string.IsNullOrEmpty(Chart_OP_Datas[i].AssignET))
                        {
                            //進度確認
                            if (TimeDiff(Chart_OP_Datas[i].AssignET, DateTime.Now.ToString()) <= 0)
                            {
                                ProgressState = 1;//進度正常
                                ProgressST = DateTime.Parse(Chart_OP_Datas[i].Real_ST).ToString(_timeFormat_ToMin);
                                ProgressET = DateTime.Now.ToString(_timeFormat_ToMin);
                            }
                            else //進度延遲
                            {
                                ProgressST = string.IsNullOrEmpty(Chart_OP_Datas[i].Real_ST) ? "" : Convert.ToDateTime(Chart_OP_Datas[i].Real_ST).ToString(_timeFormat_ToMin);
                                ProgressET = DateTime.Now.ToString(_timeFormat_ToMin);
                                ProgressState = 2;//進度延遲
                                ChooseMethod = true;//確認執行方式
                            }
                        }
                        break;
                }


                if (ProgressState == 1 || ProgressState == 2)
                {
                    OP_stateTmp.Add(new ProgressOPstates
                    {
                        OrderID = Chart_OP_Datas[i].OrderID.Trim(),
                        OPID = Chart_OP_Datas[i].OPID,
                        ProgressState = ProgressState,
                        ProgressST = ProgressST,
                        ProgressET = ProgressET,
                        ChooseMethod = ChooseMethod,
                    });
                }


                //加入最後一列資料
                if (i == Chart_OP_Datas.Count() - 1)
                {
                    if (OP_stateTmp.Count != 0)
                    {
                        result.Add(new ProgressChartWorkConditions
                        {
                            WorkGroup = MachineIDTmp,
                            OPConditions = OP_stateTmp
                        });
                    }

                }


            }
            return new ActionResponse<List<ProgressChartWorkConditions>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 工單製程甘特圖進度預覽 (O)*
        /// </summary>
        /// <param name="mode">【0:原始排程、1:交期優先-推薦權重、2:交期優先-自訂權重、3:機台優先、4:插單優先、5:手動排程、6:設備故障排程重排、7:設備故障排程延後工單右移】</param>
        /// <returns>
        /// 測試問題:目前測試交期優先推薦權重有發現資料庫temp01和wip工單不吻合，導致資料不合理(已解決)
        /// </returns>
        [HttpGet("ProgressChartPreview/{mode}")]
        public ActionResponse<List<ProgressChartWorkConditions>> ProgressChartPreview(string mode)
        {
            List<string> Mode_List = new List<string> { "0", "1", "2", "3", "4", "5", "6", "7" };
            var result = new List<ProgressChartWorkConditions>();
            if (Mode_List.Exists(x => x == mode))
            {
                string Assign = string.Empty;

                if (mode == "0")
                {
                    Assign = "Assignment";
                }
                else
                {
                    Assign = "AssignmentTemp" + mode;
                }

                List<ProgressChartOPData> Chart_OP_Datas = new List<ProgressChartOPData>();
                int i;
                var SqlStr = $@"SELECT w.OrderID, w.OPID, w.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST, w.EndTime as Real_ET, w.WIPEvent
                            FROM  WIP as w
                            LEFT JOIN  {Assign} as a ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                            WHERE w.WIPEvent='1' OR w.WIPEvent='3'
                            ORDER BY w.WorkGroup ASC, Assign_ST ASC";

                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(SqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                        {
                            conn.Open();
                        }

                        using (SqlDataReader SqlData = comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                while (SqlData.Read())
                                {
                                    Chart_OP_Datas.Add(new ProgressChartOPData
                                    {
                                        OrderID = SqlData["OrderID"].ToString().Trim(),
                                        OPID = SqlData["OPID"].ToString(),
                                        Workgroup = SqlData["WorkGroup"].ToString(),
                                        AssignST = SqlData["Assign_ST"].ToString(),
                                        AssignET = SqlData["Assign_ET"].ToString(),
                                        Real_ST = SqlData["Real_ST"].ToString(),
                                        Real_ET = SqlData["Real_ET"].ToString(),
                                        WIPEvent = SqlData["WIPEvent"].ToString(),
                                    });
                                }
                            }
                        }
                    }
                }


                string MachineIDTmp = "";//機台名稱
                int ProgressState; //進度狀態(2:延遲 1:正常)
                string ProgressST;//進度延遲or正常區間
                string ProgressET;
                bool ChooseMethod;//是否要選擇執行方式


                var OP_stateTmp = new List<ProgressOPstates>();
                for (i = 0; i < Chart_OP_Datas.Count(); i++)
                {
                    ProgressState = 0;
                    ProgressST = "";
                    ProgressET = "";
                    ChooseMethod = false;

                    if (i == 0)
                    {

                        MachineIDTmp = Chart_OP_Datas[i].Workgroup;
                    }

                    //換到下一機台
                    if (Chart_OP_Datas[i].Workgroup != MachineIDTmp)
                    {
                        if (OP_stateTmp.Count != 0)
                        {
                            result.Add(new ProgressChartWorkConditions
                            {
                                WorkGroup = MachineIDTmp,
                                OPConditions = OP_stateTmp
                            });
                        }


                        MachineIDTmp = Chart_OP_Datas[i].Workgroup;
                        OP_stateTmp = new List<ProgressOPstates>();
                    }


                    switch (Chart_OP_Datas[i].WIPEvent.Trim())
                    {
                        //製程已完成
                        case "3":
                            //檢查是否ST/ET都有正常數值
                            if (!string.IsNullOrEmpty(Chart_OP_Datas[i].AssignET) && !string.IsNullOrEmpty(Chart_OP_Datas[i].Real_ET))
                            {
                                //進度確認
                                if (TimeDiff(Chart_OP_Datas[i].AssignET, Chart_OP_Datas[i].Real_ET) > 0)
                                {
                                    ProgressState = 2; //進度延遲
                                    ProgressST = DateTime.Parse(Chart_OP_Datas[i].Real_ST).ToString(_timeFormat_ToMin);
                                    ProgressET = DateTime.Parse(Chart_OP_Datas[i].Real_ET).ToString(_timeFormat_ToMin);
                                    ChooseMethod = true; //確認執行方式
                                }
                            }
                            break;

                        //製程進行中
                        case "1":
                            //檢查是否ST有正常數值
                            if (!string.IsNullOrEmpty(Chart_OP_Datas[i].AssignET))
                            {
                                //進度確認
                                if (TimeDiff(Chart_OP_Datas[i].AssignET, DateTime.Now.ToString()) <= 0)
                                {
                                    ProgressState = 1;//進度正常
                                    ProgressST = DateTime.Parse(Chart_OP_Datas[i].Real_ST).ToString(_timeFormat_ToMin);
                                    ProgressET = DateTime.Now.ToString(_timeFormat_ToMin);
                                }
                                else //進度延遲
                                {
                                    ProgressST = string.IsNullOrEmpty(Chart_OP_Datas[i].Real_ST) ? "" : Convert.ToDateTime(Chart_OP_Datas[i].Real_ST).ToString(_timeFormat_ToMin);
                                    ProgressET = DateTime.Now.ToString(_timeFormat_ToMin);
                                    ProgressState = 2;//進度延遲
                                    ChooseMethod = true;//確認執行方式
                                }
                            }
                            break;
                    }


                    if (ProgressState == 1 || ProgressState == 2)
                    {
                        OP_stateTmp.Add(new ProgressOPstates
                        {
                            OrderID = Chart_OP_Datas[i].OrderID.Trim(),
                            OPID = Chart_OP_Datas[i].OPID,
                            ProgressState = ProgressState,
                            ProgressST = ProgressST,
                            ProgressET = ProgressET,
                            ChooseMethod = ChooseMethod,
                        });
                    }


                    //加入最後一列資料
                    if (i == Chart_OP_Datas.Count() - 1)
                    {
                        if (OP_stateTmp.Count != 0)
                        {
                            result.Add(new ProgressChartWorkConditions
                            {
                                WorkGroup = MachineIDTmp,
                                OPConditions = OP_stateTmp
                            });
                        }

                    }
                }
            }

            return new ActionResponse<List<ProgressChartWorkConditions>>
            {
                Message = "N/A",
                Data = result
            };
        }

        /// <summary>
        /// 修改生管預交日 (O)
        /// </summary>
        /// <param name="request">工單編號、製程編號、要更改的時間</param>
        /// <returns>
        /// 測資:
        /// {
        ///   "orderID": "1219111325",
        ///   "opid": "30",
        ///   "assignDate_PM": "2022-01-01 00:00:00"
        /// }
        /// </returns>
        [HttpPost("EditPredictDate")]
        public ActionResponse<string> EditPredictDate([FromBody] ScheduleEditPredictDateRequest request)
        {
            string result = "";
            string SqlStr = @"UPDATE Assignment SET AssignDate_PM =@AssignDate_PM WHERE OrderID= @OrderID and OPID =@OPID
                        ";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    try
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        comm.Parameters.Add("@AssignDate_PM", SqlDbType.DateTime).Value = request.AssignDate_PM;
                        comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = request.OrderID.Trim();
                        comm.Parameters.Add("@OPID", SqlDbType.Float).Value = request.OPID;

                        int t = comm.ExecuteNonQuery();
                        if (t != 0)
                        {
                            result = "AssignDate_PM Update Successful";
                        }
                        else
                        {
                            result = "AssignDate_PM Update Failed";
                        }
                    }
                    catch
                    {
                        return new ActionResponse<string>
                        {
                            Data = "輸入資料有誤 !"
                        };
                    }
                }
            }

            return new ActionResponse<string>
            {

                Data = result
            };
        }

        /// <summary>
        /// 標註重要工單 (O)
        /// </summary>
        /// <param name="request">工單編號、製程編號、是否重要(true、false)</param>
        /// <returns>
        /// 測資:
        /// {
        ///   "orderID": "1219111325",
        ///   "opid": "30",
        ///   "important": true
        /// }
        /// </returns>
        [HttpPost("EditImportantOrder")]
        public ActionResponse<string> EditImportantOrder([FromBody] ScheduleEditImportantOrderRequest request)
        {
            string result = "";
            string SqlStr = @"UPDATE  [Assignment] SET Important=@Important WHERE OrderID= @OrderID AND OPID=@opid
                        ";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    try
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        comm.Parameters.Add("@Important", SqlDbType.NVarChar).Value = request.Important;
                        comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = request.OrderID.Trim();
                        comm.Parameters.Add("@opid", SqlDbType.Float).Value = request.OPID.Trim();

                        int t = comm.ExecuteNonQuery();
                        if (t != 0)
                        {
                            result = "ImportantOrder Update Successful";
                        }
                        else
                        {
                            result = "ImportantOrder Update Failed";
                        }
                    }
                    catch
                    {
                        return new ActionResponse<string>
                        {
                            Data = "輸入資料有誤 !"
                        };
                    }
                }
            }

            return new ActionResponse<string>
            {

                Data = result
            };
        }

        /// <summary>
        /// 捨棄排程 (O)*
        /// </summary>
        /// <param name="request">排程模式【1:交期優先-推薦權重、2:交期優先-自訂權重、3:機台優先、4:插單優先、5:手動排程、6:設備故障排程重排、7:設備故障排程延後工單右移】</param>
        /// <returns>測資:{"mode": "3"} </returns>
        [HttpPost("Cancel")]
        public ActionResponse<bool> Cancel([FromBody] CancelScheduleRequest request)
        {
            var Mode_List = new List<string> { "1", "2", "3", "4", "5", "6", "7" };
            if (!Mode_List.Exists(x => x == request.Mode))
            {
                return new ActionResponse<bool>
                {
                    Data = false
                };
            }

            bool result = true;

            //刪除tempXX的DB table
            string SqlStr = @"DELETE AssignmentTemp" + request.Mode.ToString();
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    try
                    {
                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        int t = comm.ExecuteNonQuery();
                        if (t != 0)
                        {
                            result = true;
                        }
                        else
                        {
                            result = false;
                        }
                    }
                    catch
                    {
                        result = false;
                        return new ActionResponse<bool>
                        {
                            Data = result
                        };
                    }
                }
            }

            if (request.Mode == "6")
            {
                string sqlStr = $@"delete [DeviceBreakdownInfoTemp]";

                using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                    {
                        int impactrow = comm.ExecuteNonQuery();
                    }
                }
            }

            return new ActionResponse<bool>
            {
                Data = result
            };
        }

        /// <summary>
        /// 所有排程結果代碼 (!!!)
        /// </summary>
        /// <returns></returns>
        private ActionResponse<List<ScheduleCode>> ScheduleCode()
        {
            var result = new List<ScheduleCode>();
            result.Add(new Controllers.ScheduleCode
            {
                mode = "0",
                ScheduleMethod = "原始排程"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "1",
                ScheduleMethod = "交期優先-系統權重"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "2",
                ScheduleMethod = "交期優先-自訂權重"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "3",
                ScheduleMethod = "機台優先"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "4",
                ScheduleMethod = "插單優先"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "5",
                ScheduleMethod = "手動排程"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "6",
                ScheduleMethod = "設備故障排程重排"
            });
            result.Add(new Controllers.ScheduleCode
            {
                mode = "7",
                ScheduleMethod = "設備故障排程延後工單右移"
            });

            return new ActionResponse<List<ScheduleCode>>
            {
                Message = "N/A",
                Data = result
            };
        }

        //---------------------------------------------------------------------------------------------------
        //以下為交期優先排程流程相關API

        /// <summary>
        /// 交期優先排程法 (O)*
        /// </summary>
        /// <returns>
        /// {
        ///   "startTime": "2021-05-01 00:00:00",
        ///   "endTime": "2022-02-01 00:00:00",
        ///   "mode": "C",
        ///   "selectOrders": [
        ///     {
        ///       "orderID": "1219111460",
        ///       "opid": "30"
        ///     }
        ///   ],
        ///   "weights": 
        ///   {
        ///     "orderFillRate": "40",
        ///     "utilizationRate": "30",
        ///     "movingDistance": "10",
        ///     "loadBalance": "20"
        ///   }
        /// }
        /// </returns>
        [HttpPost("GenerateInitialSolutionforDueDayFirst")]
        public ActionResponse<string> GenerateInitialSolutionforDueDayFirst([FromBody] ScheduleDateRangeRequest request)
        {
            DateTime starttime = DateTime.Now;
            if (request.SelectOrders.Count <= 0)
            {
                throw new Exception("No work order, process selected");
            }
            int iteration = _ConnectStr.iteration;
            int chromosomesCount = _ConnectStr.chromosomesCount;
            var current = DateTime.Now;

            //取得機台列表
            var devices = DeviceInfo();

            //var result = new List<Schedulelist>();
            var result = string.Empty;

            List<int> modelist = new List<int>();
            switch (request.Mode)
            {
                case "A":
                    modelist.Add(1);
                    break;
                case "B":
                    modelist.Add(2);
                    break;
                case "C":
                    modelist.Add(1);
                    modelist.Add(2);
                    break;
            }

            // 2024.05.27 合并后的数据库查询
            var sqlStr = $@"
                            SELECT a.OrderID, a.OPID, DATEDIFF(MINUTE,a.StartTime,a.EndTime) as Optime, a.StartTime, a.EndTime, a.WorkGroup,
                                   c.BreakdownET, a.SeriesID, a.OrderQTY, a.Scheduled
                            FROM {_ConnectStr.APSDB}.dbo.Assignment as a
                            INNER JOIN {_ConnectStr.APSDB}.dbo.WIP as b ON a.SeriesID = b.SeriesID
                            LEFT JOIN {_ConnectStr.APSDB}.dbo.DeviceBreakdownInfo as c ON a.WorkGroup = c.MachineName
                            WHERE a.Scheduled = 1 AND b.WIPEvent != 3 AND a.StartTime IS NOT NULL AND c.BreakdownET >= GETDATE() AND c.IsFixed = 0";
            var originSchedule = new List<ScheduleDto>();
            var mcbkresult = new List<DelayScheduleOP>();
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    using (var sqlData = comm.ExecuteReader())
                    {
                        
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                originSchedule.Add(new ScheduleDto(
                                    sqlData["OrderID"].ToString().Trim(),
                                    Convert.ToInt32(sqlData["OPID"]),
                                    sqlData["WorkGroup"].ToString(),
                                    sqlData["StartTime"].ToString(),
                                    sqlData["EndTime"].ToString(),
                                    (int)Convert.ToDouble(sqlData["Optime"].ToString())));

                                mcbkresult.Add(new DelayScheduleOP
                                {
                                    MachineName = sqlData["WorkGroup"].ToString().Trim(),
                                    BreakdownET = _ChangeTimeFormat(checkNoword(sqlData["BreakdownET"].ToString())),
                                });
                            }
                        }
                    }
                }
            }



            //工單資訊
            var OrderList = new List<OrderList>();
            var listsqistr = $@"SELECT OrderID,OPID,Range,StartTime,EndTime
                              FROM {_ConnectStr.APSDB}.[dbo].[Assignment]
                              where '{request.EndTime}'>=AssignDate and AssignDate>='{request.StartTime}'";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (var comm = new SqlCommand(listsqistr, conn))
                {
                    using (var sqlData = comm.ExecuteReader())
                    {
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                OrderList.Add(new OrderList{
                                    OrderID = sqlData["OrderID"].ToString().Trim(),
                                    OPID = sqlData["OPID"].ToString().Trim(),
                                    Range = Convert.ToInt32(sqlData["Range"])

                                });
                            }
                        }
                    }
                }
            }

            var PMC = new DelayMthod(chromosomesCount, devices);
            // 2024.05.27新增
            // 使用HashSet替代List进行查找操作
            //HashSet<string> orderSet = new HashSet<string>(request.SelectOrders.Select(o => o.OrderID + "-" + o.OPID));

            Dictionary<string, int> orderseq = new Dictionary<string, int>();
            var selectedOrders = new HashSet<string>();
            //// 创建一个新的HashSet用于存储已选择的订单和操作ID
            //foreach (var item in request.SelectOrders)
            //{
            //    //取得排序最小值
            //    if(!orderseq.ContainsKey(item.OrderID))
            //    {
            //        orderseq[item.OrderID] = OrderList
            //                                .Where(x => x.OrderID == item.OrderID && x.OPID == item.OPID)
            //                                .Select(x => x.Range)
            //                                .Min();
            //    }
            //    selectedOrders.Add($"{item.OrderID}-{item.OPID}-{orderseq[item.OrderID]}");
            //    orderseq[item.OrderID]++;
            //}
            ////var selectedOrders = new HashSet<string>(request.SelectOrders.Select(o => $"{o.OrderID}-{o.OPID}"));
            //var addedOrders = new HashSet<string>();


            //挑選未開工且預交日期在所選區間的製程(可以重排的)
            var FirSolution = PMC.CreateDataSet(request.StartTime, request.EndTime);
            if (FirSolution.Count != 0)
            {
                foreach (int a in modelist)
                {
                    var ChromosomeList = new Dictionary<int, List<Chromsome>>();
                    var requestOrderSets = new List<GaSchedule>();

                    if (FirSolution.Count != 0)
                    {
                        foreach (var item in FirSolution)
                        {
                            // 2024.05.27修改
                            string key = $"{item.OrderID}-{item.OPID}-{item.Range}";
                            if (item.Scheduled == 1 || request.SelectOrders.Exists(x=>x.OrderID==item.OrderID && Convert.ToDouble(x.OPID)==item.OPID))
                            {
                                requestOrderSets.Add(item);
                            }
                        }
                    }






                    //取得當前機台最後時間
                    var reportedMachine = new Dictionary<string, DateTime>();
                    foreach (var item in originSchedule.Distinct(x => x.WorkMachine).Select(x => x.WorkMachine))
                    {
                        var endTimes = originSchedule.Where(x => x.WorkMachine == item && (!requestOrderSets.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID))).ToList();
                        if (endTimes.Count != 0)
                        {
                            var lastET = endTimes.Max(x => x.EndTime);
                            if (mcbkresult.Exists(x => x.MachineName == item))
                            {
                                var mcbkET = Convert.ToDateTime(mcbkresult.Find(x => x.MachineName == item).BreakdownET);
                                reportedMachine.Add(item, mcbkET > lastET ? mcbkET : lastET);
                            }
                            else
                            {
                                if (DateTime.Compare(lastET, current) >= 0)
                                {
                                    reportedMachine.Add(item, lastET);
                                }
                            }
                        }
                        else if (mcbkresult.Exists(x => x.MachineName == item))
                        {
                            var mcbkET = Convert.ToDateTime(mcbkresult.Find(x => x.MachineName == item).BreakdownET);
                            reportedMachine.Add(item, mcbkET);
                        }
                    }

                    //取得當前工單最後時間
                    var reportedOrder = new Dictionary<string, DateTime>();
                    foreach (var item in originSchedule.Distinct(x => x.OrderID).Select(x => x.OrderID))
                    {
                        var endTimes = originSchedule.Where(x => x.OrderID == item && (!requestOrderSets.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID)))
                                                     .ToList();
                        if (endTimes.Count != 0)
                        {
                            var lastET = endTimes.Max(x => x.EndTime);
                            if (DateTime.Compare(lastET, current) >= 0)
                            {
                                reportedOrder.Add(item, lastET);
                            }
                        }
                    }

                    PMC.ReportedMachine = reportedMachine;
                    PMC.ReportedOrder = reportedOrder;

                    
                    // 2024.05.27 使用Parallel.For并行生成初始解和计算适应度
                    Parallel.For(0, PMC.Chromvalue, i =>
                    {
                        var sequence = PMC.CreateSequence(requestOrderSets);
                        var train = PMC.Scheduled(sequence);
                        lock (ChromosomeList)
                        {
                            ChromosomeList.Add(i, train);
                        }
                    });

                    int noImprovementCount = 0;
                    int maxNoImprovementIterations = 10; // 最大無改進迭代次數
                    int recorditeration = 0;
                    //iteration迭代
                    for (int i = 0; i < iteration; i++)
                    {
                        PMC.EvaluationFitness(ref ChromosomeList,ref noImprovementCount);
                        // 如果無改進迭代次數達到最大值，提前終止迭代
                        if (noImprovementCount >= maxNoImprovementIterations)
                        {
                            recorditeration = i;
                            break;
                        }
                    }

                    //最佳解
                    List<Evafitnessvalue> fitness_idx_value = new List<Evafitnessvalue>();
                    for (int i = 0; i < ChromosomeList.Count; i++)
                    {
                        int sumDelay = ChromosomeList[i].Sum(x => x.Delay);
                        fitness_idx_value.Add(new Evafitnessvalue(i, sumDelay));
                    }
                    //calculate and sorting
                    fitness_idx_value.Sort((x, y) => { return x.Fitness.CompareTo(y.Fitness); });
                    var tempResult = ChromosomeList[fitness_idx_value[0].Idx];

                    # region 總排程調整
                    var OutsourcingList = PMC.getOutsourcings();
                    foreach (var item in originSchedule)
                    {
                        if (!tempResult.Exists(x => x.OrderID == item.OrderID && x.OPID == item.OPID))
                        {
                            tempResult.Add(new Chromsome
                            {
                                OrderID = item.OrderID,
                                OPID = item.OPID,
                                StartTime = item.StartTime,
                                EndTime = item.EndTime,
                                WorkGroup = item.WorkMachine,
                                Duration = item.Duration,
                            });
                        }
                    }
                    var orderList = tempResult.Distinct(x => x.OrderID)
                                  .Select(x => x.OrderID)
                                  .ToList();

                    for (int k = 0; k < 2; k++)
                    {
                        foreach (var one_order in orderList)
                        {
                            //挑選同工單製程
                            var itt = tempResult.Where(x => x.OrderID == one_order)
                                             .OrderBy(x => x.Range)
                                             .ToList();

                            for (int i = 1; i < itt.Count; i++)
                            {
                                int idx;

                                //調整同工單製程
                                if (DateTime.Compare(Convert.ToDateTime(itt[i - 1].EndTime), Convert.ToDateTime(itt[i].StartTime)) > 0)
                                {
                                    idx = tempResult.FindIndex(x => x.OrderID == itt[i].OrderID && x.OPID == itt[i].OPID);
                                    tempResult[idx].StartTime = itt[i - 1].EndTime;
                                    tempResult[idx].EndTime = itt[i - 1].EndTime + itt[i].Duration;
                                    itt[i].StartTime = itt[i - 1].EndTime;
                                    itt[i].EndTime = itt[i - 1].EndTime + itt[i].Duration;
                                }
                                //再調整同機台製程
                                //確認機台列表中有該機台
                                if (OutsourcingList.Exists(x => x.remark == itt[i].WorkGroup))
                                {
                                    //該機台非委外
                                    if (OutsourcingList.Where(x => x.remark == itt[i].WorkGroup).First().isOutsource == "0")
                                    {
                                        //調整同機台製程
                                        if (tempResult.Exists(x => itt[i].WorkGroup == x.WorkGroup))
                                        {
                                            var sequence = tempResult.Where(x => x.WorkGroup == itt[i].WorkGroup)
                                                                 .OrderBy(x => x.StartTime)
                                                                 .ToList();
                                            for (int j = 1; j < sequence.Count; j++)
                                            {
                                                if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                                {
                                                    idx = tempResult.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                                    tempResult[idx].StartTime = sequence[j - 1].EndTime;
                                                    tempResult[idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                                    sequence[j].StartTime = sequence[j - 1].EndTime;
                                                    sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region 第二次調整(避免前後製程中間有空閒時間)
                    //tempResult = tempResult.OrderBy(x => x.WorkGroup).ThenBy(x => x.StartTime).ToList();
                    tempResult = tempResult.OrderBy(x => x.Range).ThenBy(x => x.StartTime).ToList();
                    for (int k = 0; k < 3; k++)
                    {
                        foreach (var item in tempResult)
                        {
                            //同機台前面製程
                            var samewg = tempResult.FindAll(x => x.WorkGroup == item.WorkGroup && Convert.ToDateTime(x.EndTime) < Convert.ToDateTime(item.EndTime)).OrderByDescending(x => x.EndTime).ToList();
                            //同工單前面製程 
                            var sameod = tempResult.FindAll(x => x.OrderID == item.OrderID && Convert.ToInt32(x.Range) < Convert.ToInt32(item.Range)).OrderByDescending(x => x.Range).ToList();

                            TimeSpan TS = item.Duration;

                            if (samewg.Count() > 0 || sameod.Count() > 0)
                            {
                                //比較同工單前製程與同機台前製程結束時間
                                if (samewg.Count() > 0 && sameod.Count() > 0)
                                {
                                    //同步機台，直接看同工單即可
                                    if (OutsourcingList.Where(x => x.remark == item.WorkGroup).First().isOutsource == "1")
                                    {
                                        item.StartTime = sameod[0].EndTime;
                                        item.EndTime = item.StartTime + TS;
                                    }
                                    //非同步機台，要看同機台同工單
                                    else
                                    {
                                        var time1 = Convert.ToDateTime(samewg[0].EndTime);
                                        var time2 = Convert.ToDateTime(sameod[0].EndTime);
                                        item.StartTime = time1 >= time2 ? time1 : time2;
                                        item.EndTime = item.StartTime + TS;
                                    }

                                }
                                else if (samewg.Count() == 0 && sameod.Count() > 0)
                                {
                                    item.StartTime = sameod[0].EndTime;
                                    item.EndTime = item.StartTime + TS;
                                }
                                else if (sameod.Count() == 0 && samewg.Count() > 0)
                                {
                                    // 同步機台，直接看當前時間即可
                                    if (OutsourcingList.Where(x => x.remark == item.WorkGroup).First().isOutsource == "1")
                                    {
                                        item.StartTime = current;
                                        item.EndTime = Convert.ToDateTime(item.StartTime) + TS;
                                    }
                                    //非同步機台，看同機台時間
                                    else
                                    {
                                        item.StartTime = samewg[0].EndTime;
                                        item.EndTime = item.StartTime + TS;
                                    }   
                                }
                                else
                                {
                                    if (DateTime.Compare(Convert.ToDateTime(current), Convert.ToDateTime(item.StartTime)) < 0)
                                    {
                                        item.StartTime = current;
                                        item.EndTime = Convert.ToDateTime(item.StartTime) + TS;
                                    }
                                }

                                ////更新Tep
                                //var idx = tempResult.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                                //tempResult[idx].StartTime = item.StartTime;
                                //tempResult[idx].EndTime = item.EndTime;
                            }
                        }
                    }
                    #endregion

                    var temp = new List<Schedule>();//甘特圖圖表資料
                    clearntemp(@$"DELETE {_ConnectStr.APSDB}.dbo.AssignmentTemp" + a.ToString());
                    string SqlStr = @$"INSERT INTO {_ConnectStr.APSDB}.dbo.AssignmentTemp{a} ([SeriesID],[OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime],[StartTime]
                            ,[EndTime],[WorkGroup],[Operator],[AssignDate],[Parent],[SAP_WorkGroup]
                              ,[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice],[IsSkip],[MAKTX]
                             ,[PRIORITY],[Note],[Important])
                           SELECT [SeriesID],[OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime],[StartTime]
                              ,[EndTime],[WorkGroup],[Operator],[AssignDate],[Parent],[SAP_WorkGroup]
                              ,[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice],[IsSkip],[MAKTX]
                              ,[PRIORITY],[Note],[Important]
                              FROM {_ConnectStr.APSDB}.dbo.Assignment";
                    using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                    {
                        using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                        {

                            if (conn.State != ConnectionState.Open)
                                conn.Open();
                            int t = comm.ExecuteNonQuery();
                        }
                    }

                    foreach (var item in tempResult)
                    {
                        SqlStr = @$" update {_ConnectStr.APSDB}.dbo.AssignmentTemp{a}
                                 set StartTime = @StartTime, EndTime = @EndTime, WorkGroup = @WorkGroup, Scheduled = 2
                                 where OrderID = @OrderID and OPID = @OPID";
                        using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                        {
                            using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                            {
                                if (conn.State != ConnectionState.Open)
                                    conn.Open();
                                comm.Parameters.Add("@OrderID", SqlDbType.VarChar).Value = item.OrderID;
                                comm.Parameters.Add("@OPID", SqlDbType.VarChar).Value = item.OPID;
                                comm.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = item.StartTime.ToString(_timeFormat);
                                comm.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = item.EndTime.ToString(_timeFormat);
                                comm.Parameters.Add("@WorkGroup", SqlDbType.NVarChar).Value = item.WorkGroup;
                                int t = comm.ExecuteNonQuery();
                            }
                        }
                    }
                    result = $"Data save successful, iteration {recorditeration} time";
                }
            }
            else
            {
                throw new Exception("No data available in the interval");
            }
            DateTime endtime = DateTime.Now;
            TimeSpan Duration = endtime - starttime;
            //CountDelay(Tep);
            //Tep.OrderBy(x => x.OrderID).ThenBy(x => x.OPID);
            return new ActionResponse<string>
            {
                Data = result+$" ,{request.SelectOrders.Count()} OP, Spend {Duration.TotalSeconds} seconds"
            };
        }

        private List<Device> DeviceInfo()
        {
            string SqlStr = $@"SELECT  a.*, w.WIPEvent, CPK,
                            Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0)
                            FROM {_ConnectStr.APSDB}.[dbo].Assignment as a
                            LEFT JOIN {_ConnectStr.APSDB}.[dbo].WIP as w ON (a.OrderID = w.OrderID AND a.OPID = w.OPID)";
            List<Device> devices = new List<Device>();
            SqlStr = $@"SELECT * FROM  {_ConnectStr.APSDB}.[dbo].Device";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    //取得工單列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                devices.Add(new Device
                                {
                                    ID = int.Parse(SqlData["ID"].ToString()),
                                    MachineName = SqlData["MachineName"].ToString(),
                                    Remark = SqlData["Remark"].ToString(),
                                    GroupName = SqlData["GroupName"].ToString(),
                                });
                            }
                        }
                    }
                }
            }
            return devices;
        }

        [ApiExplorerSettings(IgnoreApi = true)]
        public void CountDelay(List<OrderInfo> Tep)
        {
            TimeSpan temp;
            int itemDelay;
            DateTime CurrentPredictTime = new DateTime();
            foreach (var item in Tep)
            {
                if (string.IsNullOrEmpty(item.StartTime))
                    continue;

                //判斷目前預交日
                if (string.IsNullOrEmpty(item.AssignDate_PM)) CurrentPredictTime = Convert.ToDateTime(item.AssignDate);
                else CurrentPredictTime = Convert.ToDateTime(item.AssignDate_PM);

                // 結束時間 > 預交日
                temp = CurrentPredictTime - Convert.ToDateTime(item.EndTime);
                itemDelay = (temp.TotalDays > 0) ? 0 : Math.Abs(temp.Days);
                item.DelayDays = itemDelay;
            }
        }

        private void clearntemp(string SqlStr)
        {
            //string SqlStr = @"DELETE AssignmentTemp1";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    int t = comm.ExecuteNonQuery();
                }
            }
        }

        /// <summary>
        /// 取得排程權重資料 (O)
        /// </summary>
        /// <returns></returns>
        [HttpGet("ScheduleWeight")]
        public ActionResponse<List<ScheduleWeight>> ScheduleWeight()
        {
            var result = new List<ScheduleWeight>();
            result.Add(new ScheduleWeight
            {
                DisplayName = "準交率",
                Key = "OrderFillRate",
                WeightDefaultValue = 40,
                MiniWeight = 0,
                MaxWeight = 100

            });
            result.Add(new ScheduleWeight
            {
                DisplayName = "稼動率",
                Key = "UtilizationRate",
                WeightDefaultValue = 30,
                MiniWeight = 0,
                MaxWeight = 100

            });
            result.Add(new ScheduleWeight
            {
                DisplayName = "移動距離",
                Key = "MovingDistance",
                WeightDefaultValue = 10,
                MiniWeight = 0,
                MaxWeight = 100

            });
            result.Add(new ScheduleWeight
            {
                DisplayName = "負載平衡",
                Key = "LoadBalance",
                WeightDefaultValue = 20,
                MiniWeight = 0,
                MaxWeight = 100

            });
            return new ActionResponse<List<ScheduleWeight>> { Data = result };
        }



        /// <summary>
        /// 數據比較(O)
        /// </summary>
        /// <returns></returns>
        [HttpGet("Comparison")]
        public ActionResponse<List<ScheduleDataComparison>> Comparison()
        {
            //判斷tempXX是否有數據
            var sqlStr = "";

            List<int> templist = new List<int>();
            for (int i = 1; i <= 7; i++)
            {
                var Assign = "AssignmentTemp" + i.ToString();
                sqlStr = $@"SELECT * FROM {_ConnectStr.APSDB}.dbo.{Assign} WHERE StartTime>='2024-04-01 00:00:00'";
                using (var conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (var comm = new SqlCommand(sqlStr, conn))
                    {
                        if (conn.State != ConnectionState.Open)
                        {
                            conn.Open();
                        }
                        using (SqlDataReader SqlData = comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                templist.Add(i);
                            }
                        }
                    }
                }
            }

            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            List<string> CTVal_List = new List<string> { "12:34:50", "10:18:35", "15:37:04", "11:20:49" };
            List<string> WTVal_List = new List<string> { "00:48:54", "00:56:30", "00:43:11", "00:28:44" };

            var WTVal_choose = WTVal_List[rnd.Next(WTVal_List.Count)];
            var STVal_choose = CTVal_List[rnd.Next(CTVal_List.Count)];


            var result = new List<ScheduleDataComparison>();

            result.Add(new ScheduleDataComparison
            {
                Item = "原始排程",
                WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                TotalDelay = fitnessService.TotalDelay("Assignment").ToString() + "(minute)",
                MovingTime = fitnessService.Moving_times("Assignment").ToString() + " 次",
                UtilizationRate = fitnessService.Utilize_rate("Assignment").ToString("0.0") + "%",
                OrderFillRate = fitnessService.ontimerate("Assignment").ToString() + "%"
            });

            foreach (var item in templist)
            {
                switch (item)
                {
                    case 1:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "交期優先-系統權重",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp1").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp1").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp1").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp1").ToString() + "%"
                        });
                        break;
                    case 2:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "交期優先-自訂權重",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp2").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp2").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp2").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp2").ToString() + "%"
                        });
                        break;
                    case 3:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "機台優先",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp3").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp3").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp3").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp3").ToString() + "%"
                        });
                        break;
                    case 4:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "插單",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp4").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp4").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp4").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp4").ToString() + "%"
                        });
                        break;
                    case 5:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "手動排程",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp5").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp5").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp5").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp5").ToString() + "%"
                        });
                        break;
                    case 6:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "設備故障排程重排",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp6").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp6").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp6").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp6").ToString() + "%"
                        });
                        break;
                    case 7:
                        result.Add(new ScheduleDataComparison
                        {
                            Item = "設備故障排程延後工單右移",
                            WTValue = WTVal_List[rnd.Next(WTVal_List.Count)],
                            CTValue = CTVal_List[rnd.Next(CTVal_List.Count)],
                            TotalDelay = fitnessService.TotalDelay("AssignmentTemp7").ToString() + "(minute)",
                            MovingTime = fitnessService.Moving_times("AssignmentTemp7").ToString() + " 次",
                            UtilizationRate = fitnessService.Utilize_rate("AssignmentTemp7").ToString("0.0") + "%",
                            OrderFillRate = fitnessService.ontimerate("AssignmentTemp7").ToString() + "%"
                        });
                        break;
                }
            }

            return new ActionResponse<List<ScheduleDataComparison>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 確定選擇運算方式 (O)*
        /// </summary>
        /// <param name="request">排程結果代號(0:原始排程、1:交期優先-推薦權重、2:交期優先-自訂權重、3:機台優先、4:插單優先、5:手動排程、6:設備故障排程重排、7:設備故障排程延後工單右移)</param>                                                              
        /// <returns>
        /// 測資:
        /// {
        ///     "mode": "1"
        /// }
        /// </returns>
        [HttpPost("ScheduledMethodSelected")]
        public ActionResponse<string> ScheduledMethodSelected([FromBody] ScheduleMethodSelectionRequest request)
        {
            var Mode_List = new List<string> { "0", "1", "2", "3", "4", "5", "6", "7" };
            if (!Mode_List.Exists(x => x == request.Mode) || Convert.ToInt32(request.Mode) < 0)
            {
                return new ActionResponse<string>
                {
                    Data = "false"
                };
            }
            //判斷選用哪一種Method

            //進行Assigment的更新
            var Assign = string.Empty;
            var sqlStr = string.Empty;
            var result = string.Empty;
            if (request.Mode != "0")
            {
                Assign = "AssignMentTemp" + request.Mode;
                sqlStr = $@"update ass
                            set ass.StartTime = temp.StartTime, ass.EndTime = temp.EndTime, ass.WorkGroup = temp.WorkGroup, ass.Scheduled = 1
                            from Assignment ass left join {Assign} as temp
                            on ass.OrderID = temp.OrderID and ass.OPID = temp.OPID
                            where temp.OrderID is not null";


                using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                    {
                        int impactrow = comm.ExecuteNonQuery();
                        if (impactrow > 0)
                        {
                            result = "Update Successful";
                        }
                        else
                        {
                            result = "Can not update";
                        }

                    }
                }
                sqlStr = @"UPDATE Assignment
                           SET Scheduled = 0
                           WHERE StartTime IS NULL";
                using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                    {
                        int impactrow = comm.ExecuteNonQuery();
                    }
                }
            }
            else if (request.Mode == "0")
            {
                result = "Update Successful";
            }


            if (result == "Update Successful")
            {
                //清除所有AssignmentTemp的資料
                foreach (string i in Mode_List)
                {
                    if (i != "0")
                    {
                        var DelAssign = "AssignmentTemp" + i.ToString();

                        var DelsqlStr = $@"TRUNCATE TABLE {DelAssign}";

                        using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                        {

                            if (conn.State != ConnectionState.Open)
                                conn.Open();
                            using (SqlCommand comm = new SqlCommand(DelsqlStr, conn))
                            {
                                int impactrow = comm.ExecuteNonQuery();
                            }
                        }
                    }
                }
            }

            if (request.Mode != "0")
            {
                //更新WIP
                sqlStr = $@"update w
                        set w.WorkGroup = a.WorkGroup
                        from Assignment as a left join WIP as w
                        on a.OrderID = w.OrderID and a.OPID = w.OPID";

                using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                    {
                        int impactrow = comm.ExecuteNonQuery();
                    }
                }
            }

            // 將暫存的機台故障修復訊息寫入正式表內
            if (request.Mode == "6" || request.Mode == "7")
            {
                saveMachineBreakdownInfo();
            }

            //篩選Assigement有，但WIP沒有的工單
            List<Assignment> tempNewWIPorder = getWIPhaveNew();
            if (tempNewWIPorder.Count > 0)
            {
                List<string> diffdata = new List<string>();


                foreach (var item in tempNewWIPorder)
                {
                    if (item.StartTime != null && item.EndTime != null && item.WorkGroup != null)

                        try
                        {
                            //修改ERP工單狀態 功能已經好了 之後在串上
                            //updateERPOrderstatus(item.SeriesID);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }


                    if (InsertToWIP(item))
                    {
                        diffdata.Add(item.OrderID + " " + item.OPID);
                    }
                    else
                    {
                        diffdata.Add(item.OrderID + " " + item.OPID);
                    }
                }
            }

            return new ActionResponse<string>
            {
                Data = result
            };

        }
        private List<Assignment> getWIPhaveNew()
        {
            var result = new List<Assignment>();
            string SqlStr = @"select A.*,B.OrderID from  
                                Assignment as A
                                left join  
                                WIP as B on A.OrderID=B.OrderID 
                                WHERE B.OrderID is NULL";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new Assignment
                                {
                                    SeriesID = SqlData["SeriesID"].ToString(),
                                    OrderID = SqlData["OrderID"].ToString(),
                                    ERPOrderID = SqlData["ERPOrderID"].ToString(),
                                    OPID = SqlData["OPID"].ToString(),
                                    OPLTXA1 = SqlData["OPLTXA1"].ToString(),
                                    MachOpTime = SqlData["MachOpTime"].ToString(),
                                    HumanOpTime = SqlData["HumanOpTime"].ToString(),
                                    StartTime = SqlData["StartTime"].ToString(),
                                    EndTime = SqlData["EndTime"].ToString(),
                                    WorkGroup = SqlData["WorkGroup"].ToString(),
                                    Operator = SqlData["Operator"].ToString(),
                                    AssignDate = SqlData["AssignDate"].ToString(),
                                    OrderQTY = SqlData["OrderQTY"].ToString(),
                                    Scheduled = SqlData["Scheduled"].ToString(),
                                    AssignDate_PM = SqlData["AssignDate_PM"].ToString(),
                                    MAKTX = SqlData["MAKTX"].ToString(),
                                    PRIORITY = SqlData["PRIORITY"].ToString(),
                                    ImgPath = SqlData["ImgPath"].ToString(),
                                    Note = SqlData["Note"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                    CPK = SqlData["CPK"].ToString(),
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }

        private void saveMachineBreakdownInfo()
        {
            string sqlStr = $@"INSERT INTO [DeviceBreakdownInfo] ([MachineId],[MachineName],[BreakdownST],[BreakdownET],[IsFixed],[Createdtime],[Lastupdatetime])
                                SELECT [MachineId],[MachineName],[BreakdownST],[BreakdownET],[IsFixed],[Createdtime],[Lastupdatetime]
                                FROM [DeviceBreakdownInfoTemp]

                                delete [DeviceBreakdownInfoTemp]
                                ";

            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {

                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    int impactrow = comm.ExecuteNonQuery();
                }
            }
        }

        //更新ERP上工單狀台【待排:4=>已排程:0】
        private void updateERPOrderstatus(string Id)
        {
            string cmd = @"[{""Id"": " + Id + @",""Status"": 0}]";

            string ERPServerUrl = _ConnectStr.ERPurl;
            string SetAccessToken = WebAPIservice.gettoken(ERPServerUrl);

            WebAPIservice.RequestWebAPI_PUT(cmd, ERPServerUrl + "api/WorkShare/JobOrderUpdateStatus", SetAccessToken);
        }

        private bool InsertToWIP(Assignment data)
        {
            bool result = false;
            int EffectRow = 0;
            string SqlStr = $"insert into WIP values (@SeriesID,@OrderID,@OPID,@OrderQTY,0,0,0,0,@WorkGroup,NULL,NULL,NULL,NULL,GETDATE(),GETDATE())";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add(("@SeriesID"), SqlDbType.BigInt).Value = data.SeriesID;
                    comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = data.OrderID;
                    comm.Parameters.Add(("@OPID"), SqlDbType.Float).Value = data.OPID;
                    comm.Parameters.Add(("@OrderQTY"), SqlDbType.Int).Value = data.OrderQTY;
                    comm.Parameters.Add(("@WorkGroup"), SqlDbType.NVarChar).Value = data.WorkGroup;

                    EffectRow = comm.ExecuteNonQuery();
                }
            }
            if (EffectRow > 0)
            {
                result = true;
            }
            return result;
        }
        /// <summary>
        /// 機台優先排程法 (O)
        /// </summary>
        /// <returns>{"selectOrders": [ { "orderID": "1219111460", "opid": "30" } ]}</returns>
        [HttpPost("GenerateInitialSolutionForDeviceFirst")]
        public ActionResponse<string> GenerateInitialSolutionForDeviceFirst([FromBody] ScheduleOrdersRequest request)
        {
            if (request.SelectOrders.Count <= 0)
            {
                throw new Exception("No work order, process selected");
            }

            int iteration = _ConnectStr.iteration;
            int chromosomesCount = _ConnectStr.chromosomesCount;

            var result = string.Empty;
            var current = DateTime.Now;
            var originSchedule = new List<ScheduleDto>();

            string SqlStr = "";

            var devices = DeviceInfo();


            var sqlStr = $@"select OrderID, OPID, Optime =  HumanOpTime+(MachOpTime* OrderQTY), StartTime, EndTime, WorkGroup
                            from Assignment
                            where Scheduled = 1 and StartTime is not NULL";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    using (var sqlData = comm.ExecuteReader())
                    {
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                originSchedule.Add(new ScheduleDto(
                                    sqlData["OrderID"].ToString().Trim(),
                                    Convert.ToInt32(sqlData["OPID"]),
                                    sqlData["WorkGroup"].ToString(),
                                    sqlData["StartTime"].ToString(),
                                    sqlData["EndTime"].ToString(),
                                    (int)Convert.ToDouble(sqlData["Optime"].ToString())));
                            }
                        }
                    }
                }
            }


            var ChromosomeList = new Dictionary<int, List<Chromsome>>();
            SetupMethod PMC = new SetupMethod(chromosomesCount, devices);
            var FirSolution = PMC.CreateDataSet();
            var reportedMachine = new Dictionary<string, DateTime>();


            var requestOrderSets = new List<GaSchedule>();
            //撈出實際要重新排程的製程
            foreach (var item in FirSolution)
            {
                if (request.SelectOrders.Exists(x => x.OrderID == item.OrderID && x.OPID == item.OPID.ToString()))
                {
                    var id = FirSolution.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                    requestOrderSets.Add(FirSolution[id]);
                }
            }

            var mcbkresult = new List<DelayScheduleOP>();
            var mcbkstr = $@"
                            SELECT MachineName,BreakdownET
                            FROM DeviceBreakdownInfo
                            WHERE BreakdownET>=GETDATE() AND IsFixed=0";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(mcbkstr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                mcbkresult.Add(new DelayScheduleOP
                                {
                                    MachineName = SqlData["MachineName"].ToString().Trim(),
                                    BreakdownET = _ChangeTimeFormat(checkNoword(SqlData["BreakdownET"].ToString())),
                                });

                            }
                        }
                    }
                }
            }

            var originSchedule_WorkMachine_List = originSchedule.Distinct(x => x.WorkMachine).Select(x => x.WorkMachine).ToList();

            //找出當前各機台時間
            foreach (var item in originSchedule_WorkMachine_List)
            {
                var endTimes = originSchedule.Where(x => x.WorkMachine == item && (!requestOrderSets.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID)))
                                             .ToList();
                if (endTimes.Count != 0)
                {
                    var lastET = endTimes.Max(x => x.EndTime);
                    if (mcbkresult.Exists(x => x.MachineName == item))
                    {
                        var mcbkET = Convert.ToDateTime(mcbkresult.Find(x => x.MachineName == item).BreakdownET);
                        reportedMachine.Add(item, mcbkET > lastET ? mcbkET : lastET);
                    }
                    else
                    {
                        if (DateTime.Compare(lastET, current) >= 0)
                        {
                            reportedMachine.Add(item, lastET);
                        }
                    }
                }
                else if (mcbkresult.Exists(x => x.MachineName == item))
                {
                    var mcbkET = Convert.ToDateTime(mcbkresult.Find(x => x.MachineName == item).BreakdownET);
                    reportedMachine.Add(item, mcbkET);
                }

            }
            //取得當前工單最後時間
            var reportedOrder = new Dictionary<string, DateTime>();
            foreach (var item in originSchedule.Distinct(x => x.OrderID).Select(x => x.OrderID))
            {
                var endTimes = originSchedule.Where(x => x.OrderID == item && (!requestOrderSets.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID)))
                                             .ToList();
                if (endTimes.Count != 0)
                {
                    var lastET = endTimes.Max(x => x.EndTime);
                    if (DateTime.Compare(lastET, current) >= 0)
                    {
                        reportedOrder.Add(item, lastET);
                    }
                }
            }
            PMC.ReportedMachine = reportedMachine;
            PMC.ReportedOrder = reportedOrder;


            for (int i = 0; i < PMC.Chromvalue; i++)
            {
                var rnd = new Random(Guid.NewGuid().GetHashCode());
                var randomized = requestOrderSets.OrderBy(item => rnd.Next()).ToList();
                var sequence = PMC.CreateSequence(randomized);
                var train = PMC.Scheduled(sequence);
                ChromosomeList.Add(i, train.Select(x => (Chromsome)x.Clone()).ToList());
            }
            int noImprovementCount = 0;
            int maxNoImprovementIterations = 10;
            int recorditeration = 0;
            //iteration
            for (int i = 0; i < iteration; i++)
            {
                PMC.EvaluationFitness(ref ChromosomeList, ref noImprovementCount);
                if(noImprovementCount>= maxNoImprovementIterations)
                {
                    recorditeration = i;
                    break;
                }
            }
            // find the best solution
            List<Evafitnessvalue> fitness_idx_value = new List<Evafitnessvalue>();
            for (int i = 0; i < ChromosomeList.Count; i++)
            {
                int sumDelay = ChromosomeList[i].Sum(x => x.Delay);
                fitness_idx_value.Add(new Evafitnessvalue(i, sumDelay));
            }
            //calculate and sorting
            fitness_idx_value.Sort((x, y) => { return x.Fitness.CompareTo(y.Fitness); });
            var tempResult = ChromosomeList[fitness_idx_value[0].Idx];

            #region 總排程調整
            var OutsourcingList = PMC.getOutsourcings();
            foreach (var item in originSchedule)
            {
                if(!tempResult.Exists(x=>x.OrderID==item.OrderID && x.OPID==item.OPID))
                {
                    tempResult.Add(new Chromsome
                    {
                        OrderID = item.OrderID,
                        OPID = item.OPID,
                        StartTime = item.StartTime,
                        EndTime = item.EndTime,
                        WorkGroup = item.WorkMachine,
                        Duration = item.Duration,
                    });
                }
            }
            var orderList = tempResult.Distinct(x => x.OrderID)
                          .Select(x => x.OrderID)
                          .ToList();

            for (int k = 0; k < 2; k++)
            {
                foreach (var one_order in orderList)
                {
                    //挑選同工單製程
                    var itt = tempResult.Where(x => x.OrderID == one_order)
                                     .OrderBy(x => x.Range)
                                     .ToList();

                    for (int i = 1; i < itt.Count; i++)
                    {
                        int idx;

                        //調整同工單製程
                        if (DateTime.Compare(Convert.ToDateTime(itt[i - 1].EndTime), Convert.ToDateTime(itt[i].StartTime)) > 0)
                        {
                            idx = tempResult.FindIndex(x => x.OrderID == itt[i].OrderID && x.OPID == itt[i].OPID);
                            tempResult[idx].StartTime = itt[i - 1].EndTime;
                            tempResult[idx].EndTime = itt[i - 1].EndTime + itt[i].Duration;
                            itt[i].StartTime = itt[i - 1].EndTime;
                            itt[i].EndTime = itt[i - 1].EndTime + itt[i].Duration;
                        }
                        //若委外再調整同機台製程
                        if (OutsourcingList.Exists(x => x.remark == itt[i].WorkGroup))
                        {
                            if (OutsourcingList.Where(x => x.remark == itt[i].WorkGroup).First().isOutsource == "0")
                            {
                                //調整同機台製程
                                if (tempResult.Exists(x => itt[i].WorkGroup == x.WorkGroup))
                                {
                                    var sequence = tempResult.Where(x => x.WorkGroup == itt[i].WorkGroup)
                                                         .OrderBy(x => x.StartTime)
                                                         .ToList();
                                    for (int j = 1; j < sequence.Count; j++)
                                    {
                                        if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                        {
                                            idx = tempResult.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                            tempResult[idx].StartTime = sequence[j - 1].EndTime;
                                            tempResult[idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                            sequence[j].StartTime = sequence[j - 1].EndTime;
                                            sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            //先清空暫存資料庫table
            clearntemp($@"delete {_ConnectStr.APSDB}.[dbo].AssignmentTemp3");
            SqlStr = $@"INSERT INTO {_ConnectStr.APSDB}.[dbo].AssignmentTemp3([OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime],[StartTime]
                    ,[EndTime],[WorkGroup],[Operator],[AssignDate],[Parent],[SAP_WorkGroup]
                      ,[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice],[IsSkip],[MAKTX]
                     ,[PRIORITY],[Note],[Important])
                   SELECT [OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime],[StartTime]
                      ,[EndTime],[WorkGroup],[Operator],[AssignDate],[Parent],[SAP_WorkGroup]
                      ,[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice],[IsSkip],[MAKTX]
                      ,[PRIORITY],[Note],[Important]
                      FROM {_ConnectStr.APSDB}.[dbo].Assignment";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {

                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    int t = comm.ExecuteNonQuery();
                }
            }
            foreach (var item in tempResult)
            {
                SqlStr = @" update AssignmentTemp3
                            set StartTime = @StartTime, EndTime = @EndTime, WorkGroup = @WorkGroup, Scheduled = 2
                            where OrderID = @OrderID and OPID =  @OPID";
                using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                    {

                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        comm.Parameters.Add("@OrderID", SqlDbType.VarChar).Value = item.OrderID;
                        comm.Parameters.Add("@OPID", SqlDbType.VarChar).Value = item.OPID;
                        comm.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = item.StartTime.ToString(_timeFormat);
                        comm.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = item.EndTime.ToString(_timeFormat);
                        comm.Parameters.Add("@WorkGroup", SqlDbType.NVarChar).Value = item.WorkGroup;
                        int t = comm.ExecuteNonQuery();
                    }
                }
            }
            result = $"Data save successful, iteration {recorditeration} times";

            return new ActionResponse<string>
            {

                Data = result
            };
        }
        private List<OrderInfo> CheckOP(List<OrderInfo> Data)
        {
            string DateFormat = "yyyy-MM-dd HH:mm";
            for (int i = 1; i < Data.Count; i++)
            {
                if (DateTime.Compare(Convert.ToDateTime(Data[i - 1].EndTime), Convert.ToDateTime(Data[i].StartTime)) > 0)
                {
                    //TimeSpan TS = Convert.ToDateTime(Data[i].EndTime) - Convert.ToDateTime(Data[i].StartTime);
                    TimeSpan TS = new TimeSpan(0, 0, (int)(Data[i].HumanOpTime + (Data[i].OrderQTY * Data[i].MachOpTime)), 0, 0);
                    Data[i].StartTime = restTime.hoildaycheck(Convert.ToDateTime(Data[i - 1].EndTime)).ToString(DateFormat);
                    Data[i].EndTime = restTime.restTimecheck(Convert.ToDateTime(Data[i].StartTime), TS).ToString(DateFormat);//(Convert.ToDateTime(Data[i].StartTime) + TS).ToString(DateFormat);
                }
            }
            return Data;
        }

        /// <summary>
        /// 插單優先排程法 (O)*
        /// </summary>
        /// <returns></returns>
        [HttpPost("DispatchWork")]
        public ActionResponse<string> DispatchWork([FromBody] List<DispatchWorkScheduleRequest> request)
        {
            int iteration = _ConnectStr.iteration;
            int chromosomesCount = _ConnectStr.chromosomesCount;
            List<string> Orderlist = request.Distinct(x => x.OrderID).Select(x => x.OrderID).ToList();
            DateTime activatetime = DateTime.Now;

            //存可用設備
            var devices = new List<Device>();
            string SqlStr = @"SELECT * FROM  Device";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    //取得機台列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                devices.Add(new Device
                                {
                                    ID = int.Parse(SqlData["ID"].ToString()),
                                    MachineName = SqlData["MachineName"].ToString(),
                                    Remark = SqlData["Remark"].ToString(),
                                    GroupName = SqlData["GroupName"].ToString(),
                                });
                            }
                        }
                    }
                }
            }

            var result = string.Empty;

            //原始排程(未開工)
            var originSchedule = new List<Chromsome>();
            var sqlStr = $@"SELECT a.OrderID, a.OPID, a.Range,a.OrderQTY, Optime = HumanOpTime+MachOpTime* a.OrderQTY, a.StartTime, a.EndTime, a.AssignDate, a.WorkGroup, a.MAKTX
                            FROM Assignment a left join WIP as wip
                            on a.OrderID=wip.OrderID and a.OPID=wip.OPID
                            left join WipRegisterLog w
                            on w.WorkOrderID = a.OrderID and w.OPID=a.OPID
                            where w.WorkOrderID is NULL and (wip.WIPEvent!=3 or wip.WIPEvent is NULL) and Scheduled = 1 and a.StartTime is not null
                            order by a.WorkGroup, a.StartTime";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    using (var sqlData = comm.ExecuteReader())
                    {
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                originSchedule.Add(new Chromsome
                                {
                                    OrderID = sqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToInt32(sqlData["OPID"]),
                                    Range = Convert.ToInt32(sqlData["Range"]),
                                    WorkGroup = sqlData["WorkGroup"].ToString(),
                                    StartTime = Convert.ToDateTime(sqlData["StartTime"].ToString()),
                                    EndTime = Convert.ToDateTime(sqlData["EndTime"].ToString()),
                                    Duration = new TimeSpan(0, (int)Convert.ToDouble(sqlData["Optime"].ToString()), 0),
                                    AssignDate = Convert.ToDateTime(sqlData["AssignDate"].ToString()),
                                    Maktx = sqlData["MAKTX"].ToString().Trim()
                                });

                            }
                        }
                    }
                }
            }

            //加工或暫停中製程(不可挪動)
            var InProcessSchedule = new List<Chromsome>();
            sqlStr = $@"SELECT a.OrderID, a.OPID, a.Range,a.OrderQTY, Optime = HumanOpTime+MachOpTime* a.OrderQTY, a.StartTime, a.EndTime, a.AssignDate, a.WorkGroup, a.MAKTX
                            FROM Assignment a left join WIP as wip
                            on a.OrderID=wip.OrderID and a.OPID=wip.OPID
                            left join WipRegisterLog w
                            on w.WorkOrderID = a.OrderID and w.OPID=a.OPID
                            where w.WorkOrderID is not NULL and wip.WIPEvent in (1,2) and Scheduled = 1 and a.StartTime is not null
                            order by a.WorkGroup, a.StartTime";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (var comm = new SqlCommand(sqlStr, conn))
                {
                    using (var sqlData = comm.ExecuteReader())
                    {
                        if (sqlData.HasRows)
                        {
                            while (sqlData.Read())
                            {
                                InProcessSchedule.Add(new Chromsome
                                {
                                    OrderID = sqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToInt32(sqlData["OPID"]),
                                    Range = Convert.ToInt32(sqlData["Range"]),
                                    WorkGroup = sqlData["WorkGroup"].ToString(),
                                    StartTime = Convert.ToDateTime(sqlData["StartTime"].ToString()),
                                    EndTime = Convert.ToDateTime(sqlData["EndTime"].ToString()),
                                    Duration = new TimeSpan(0, (int)Convert.ToDouble(sqlData["Optime"].ToString()), 0),
                                    AssignDate = Convert.ToDateTime(sqlData["AssignDate"].ToString()),
                                    Maktx = sqlData["MAKTX"].ToString().Trim()
                                });

                            }
                        }
                    }
                }
            }


            var PMC = new DelayMthod(chromosomesCount, devices);
            //撈出可以重新排程的製程(未排程、未開工)
            var requestOrderSets = PMC.DispatchCreateDataSet(Orderlist, activatetime);
            var totalcount = originSchedule.Count();//所有重新排程的製程數量
            //重新設定插單的加工時間與機台
            foreach (var item in Orderlist)
            {
                var templist = requestOrderSets.Where(x => x.OrderID == item).ToList();
                foreach (var process in templist)
                {
                    int index = requestOrderSets.FindIndex(x => x.OrderID == process.OrderID && x.OPID == process.OPID);
                    requestOrderSets[index].StartTime = DateTime.MinValue;
                    requestOrderSets[index].EndTime = DateTime.MinValue;
                    requestOrderSets[index].WorkGroup = String.Empty;
                }
            }
            //可重新排程(未包含插單工單)
            var originsche = requestOrderSets.Where(x => !request.Exists(y => y.OrderID == x.OrderID)).ToList();
            originsche = originsche.OrderBy(x => x.WorkGroup).ThenBy(x => x.StartTime).ToList();
            //原排程各機台順序(未包含插單工單)
            var originseq = new List<LocalMachineSeq>();//原排程排序(未包含插單)
            int machineSeq = 0;
            foreach (var item in originsche)
            {
                if (originseq.Exists(x => x.WorkGroup == item.WorkGroup))
                {
                    machineSeq = originseq.Where(x => x.WorkGroup == item.WorkGroup)
                                       .Select(x => x.EachMachineSeq)
                                       .Max() + 1;
                }
                else
                {
                    machineSeq = 0;
                }
                originseq.Add(new LocalMachineSeq
                {
                    OrderID = item.OrderID,
                    OPID = item.OPID,
                    Range = item.Range,
                    WorkGroup = item.WorkGroup,
                    EachMachineSeq = machineSeq,
                    PredictTime = item.AssignDate,
                    Duration = item.Duration
                });
            }

            //確認是否有機台維修時間
            var mcbkresult = new List<DelayScheduleOP>();
            var mcbkstr = $@"
                            SELECT MachineName,BreakdownET
                            FROM DeviceBreakdownInfo
                            WHERE BreakdownET>=GETDATE() AND IsFixed=0";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(mcbkstr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                mcbkresult.Add(new DelayScheduleOP
                                {
                                    MachineName = SqlData["MachineName"].ToString().Trim(),
                                    BreakdownET = _ChangeTimeFormat(checkNoword(SqlData["BreakdownET"].ToString())),
                                });

                            }
                        }
                    }
                }
            }

            if (requestOrderSets.Count != 0)
            {
                //取得當前機台&當前工單最後時間
                var reportedMachine = new Dictionary<string, DateTime>();
                var reportedOrder = new Dictionary<string, DateTime>();
                //取得當前機台最後時間
                foreach (var item in InProcessSchedule.Distinct(x => x.WorkGroup).Select(x => x.WorkGroup))
                {
                    //取出已開工製程
                    var endTimes = InProcessSchedule.Where(x => x.WorkGroup == item && (!requestOrderSets.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID)))
                                                 .ToList();
                    if (endTimes.Count != 0)
                    {
                        var lastET = endTimes.Max(x => x.EndTime);
                        if (mcbkresult.Exists(x => x.MachineName == item))
                        {
                            var mcbkET = Convert.ToDateTime(mcbkresult.Find(x => x.MachineName == item).BreakdownET);
                            reportedMachine.Add(item, mcbkET > lastET ? mcbkET : lastET);
                        }
                        else
                        {
                            if (DateTime.Compare(lastET, activatetime) >= 0)
                            {
                                reportedMachine.Add(item, lastET);
                            }
                        }
                    }
                    else if (mcbkresult.Exists(x => x.MachineName == item))
                    {
                        var mcbkET = Convert.ToDateTime(mcbkresult.Find(x => x.MachineName == item).BreakdownET);
                        reportedMachine.Add(item, mcbkET);
                    }
                }
                //取得當前工單最後時間
                foreach (var item in InProcessSchedule.Distinct(x => x.OrderID).Select(x => x.OrderID))
                {
                    var endTimes = InProcessSchedule.Where(x => x.OrderID == item && (!requestOrderSets.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID)))
                                                 .ToList();
                    if (endTimes.Count != 0)
                    {
                        var lastET = endTimes.Max(x => x.EndTime);
                        if (DateTime.Compare(lastET, activatetime) >= 0)
                        {
                            reportedOrder.Add(item, lastET);
                        }
                    }
                }

                PMC.ReportedMachine = reportedMachine;
                PMC.ReportedOrder = reportedOrder;

                //算各工單緊急程度
                double CRvalue;
                List<ProcessCR> processlist = new List<ProcessCR>();
                foreach (var order in Orderlist)
                {
                    double totaltime = 0;
                    DateTime assigndate = requestOrderSets.Where(x => x.OrderID == order).Select(x => x.AssignDate).ToList()[0];
                    foreach (var item in requestOrderSets.Where(x => x.OrderID == order))
                    {
                        totaltime += item.Duration.TotalMinutes;
                    }
                    CRvalue = Math.Round((assigndate - activatetime).TotalMinutes / totaltime, 3);
                    processlist.Add(
                    new ProcessCR
                    {
                        OrderID = order,
                        CR = CRvalue
                    });
                }
                var CRlist = processlist.OrderBy(x => x.CR).ToList();//根據各製程CR值小到大排序

                List<Device> CanUseDevices = new List<Device>();
                //匹配各製程與機台
                foreach (var order in CRlist)
                {
                    foreach (var process in requestOrderSets.Where(x => x.OrderID == order.OrderID).OrderBy(x => x.Range))
                    {
                        
                        var tempschedule = PMC.Scheduled(originseq);//未包含插單製程重新排程
                        var MakeSpan = tempschedule.Select(x => x.EndTime).Max();//當前排程makespan
                        CanUseDevices = PMC.getCanUseDevice(process.OrderID, process.OPID.ToString(), process.Maktx);//可用機台
                        CanUseDevices = CanUseDevices.Distinct(x => x.Remark).ToList();

                        if (CanUseDevices.Count > 0)
                        {
                            double[] Fitness = new double[CanUseDevices.Count()];//各機台合適度
                            DateTime[] Begin = new DateTime[CanUseDevices.Count()];//各機台可排開始時間
                            var premachinetime = new DateTime();
                            var preordertime = new DateTime();
                            for (int i = 0; i < CanUseDevices.Count(); i++)
                            {
                                var preorder = tempschedule.Where(x => x.OrderID == process.OrderID).OrderByDescending(x => x.EndTime).Select(x => x.EndTime).ToList();//前一道製程完工時間
                                if (PMC.ReportedMachine.Keys.Contains(CanUseDevices[i].Remark))
                                {
                                    premachinetime = PMC.ReportedMachine[CanUseDevices[i].Remark];
                                }
                                else
                                {
                                    premachinetime = activatetime;
                                }
                                if (preorder.Count() > 0)
                                {
                                    preordertime = (preorder[0] < activatetime ? activatetime : preorder[0]); //同工單前製程時間

                                }
                                else
                                {
                                    preordertime = activatetime;
                                }
                                //比較同機台前製程&同工單前製程來決定可插入的開始時間
                                var begin = DateTime.Compare(preordertime, premachinetime) > 0 ? preordertime : premachinetime;

                                #region 用makespan算剩餘時間
                                ////剩餘時間計算
                                //var remaintime = (MakeSpan - begin).TotalMinutes - tempschedule.Where(x => x.WorkGroup == CanUseDevices[i].Remark && x.StartTime >= begin).Sum(x => x.Duration.TotalMinutes);
                                //var firstprocess = tempschedule.Where(x => x.WorkGroup == CanUseDevices[i].Remark && x.EndTime > begin && x.StartTime < begin).OrderBy(x => x.EndTime).ToList();
                                //if (firstprocess.Count > 0)
                                //{
                                //    remaintime -= (firstprocess[0].EndTime - begin).TotalMinutes;
                                //}
                                ////機台適應度=剩餘時間-加工時間(=生產數量/加工完成頻率)
                                //Fitness[i] = remaintime - (process.Duration.TotalMinutes * process.PartCount);
                                //Begin[i] = begin;
                                #endregion

                                #region 用各機台區間算剩餘時間
                                //剩餘時間計算
                                var intervalendtime = begin.AddMinutes(process.Duration.TotalMinutes);//區間結束時間
                                //區間扣除在內製程(開工&完工)的時間
                                var remaintime = (intervalendtime - begin).TotalMinutes - tempschedule.Where(x => x.WorkGroup == CanUseDevices[i].Remark && x.StartTime >= begin && x.EndTime <= intervalendtime).Sum(x => x.Duration.TotalMinutes);
                                //扣除開工時間在區間外，完工時間在區間內製程時間
                                var firstprocess = tempschedule.Where(x => x.WorkGroup == CanUseDevices[i].Remark && x.EndTime > begin && x.StartTime < begin).OrderBy(x => x.EndTime).ToList();
                                if (firstprocess.Count > 0)
                                {
                                    remaintime -= (firstprocess[0].EndTime - begin).TotalMinutes;
                                }
                                //扣除開工時間在區間內，完工時間在區間外製程時間
                                var lastprocess = tempschedule.Where(x => x.WorkGroup == CanUseDevices[i].Remark && x.EndTime > intervalendtime && x.StartTime < intervalendtime).OrderBy(x => x.EndTime).ToList();
                                if (lastprocess.Count > 0)
                                {
                                    remaintime -= (intervalendtime - lastprocess[0].StartTime).TotalMinutes;
                                }
                                //機台適應度=剩餘時間-加工時間(=生產數量/加工完成頻率)
                                Fitness[i] = remaintime;
                                Begin[i] = begin;
                                #endregion
                            }

                            //挑選適應值最大的機台
                            var final_mc_idx = Array.IndexOf(Fitness, Fitness.Max());
                            var min_time = Begin[final_mc_idx];
                            //若有一個以上的最佳機台，就選可以開始時間較早的
                            if (Fitness.Distinct().Count() != Fitness.Count())
                            {
                                for (int k = 0; k < Fitness.Count(); k++)
                                {
                                    if (Fitness[k] == final_mc_idx)
                                    {
                                        if (Begin[k] < min_time)
                                        {
                                            min_time = Begin[k];
                                            final_mc_idx = k;
                                        }
                                    }
                                }
                            }

                            var select = CanUseDevices[final_mc_idx].Remark;//選擇的機台
                                                                            //同機台seq調整
                            var forward = tempschedule.Where(x => x.WorkGroup == select && x.StartTime >= Begin[final_mc_idx]).OrderBy(x => x.StartTime).ToList();
                            if (forward.Count() != 0)
                            {
                                var insertidx = originseq.Where(x => x.OrderID == forward[0].OrderID && x.OPID == forward[0].OPID).Select(x => x.EachMachineSeq).ToList();
                                var templist = originseq.Where(x => x.WorkGroup == select && x.EachMachineSeq >= insertidx[0]).ToList();
                                foreach (var operation in templist)
                                {
                                    int idx = originseq.FindIndex(x => x.OrderID == operation.OrderID && x.OPID == operation.OPID);
                                    originseq[idx].EachMachineSeq += 1;
                                }
                                //用seq插入排程
                                originseq.Add(new LocalMachineSeq
                                {
                                    OrderID = process.OrderID,
                                    OPID = process.OPID,
                                    Range = process.Range,
                                    WorkGroup = select,
                                    EachMachineSeq = insertidx[0],
                                    PredictTime = process.AssignDate,
                                    Duration = process.Duration
                                });

                            }
                            else if (forward.Count() == 0)//若沒有後續製程，有可能排為該機台第一道或最後一道製程
                            {
                                //用seq插入排程
                                originseq.Add(new LocalMachineSeq
                                {
                                    OrderID = process.OrderID,
                                    OPID = process.OPID,
                                    Range = process.Range,
                                    WorkGroup = select,
                                    EachMachineSeq = originseq.Where(x => x.WorkGroup == select).Count(),
                                    PredictTime = process.AssignDate,
                                    Duration = process.Duration
                                });
                            }
                        }

                    }
                }

                var finalschedule = new List<Chromsome>();
                finalschedule = PMC.Scheduled(originseq);

                ////進行排程擾動
                ////double T = 0.1; //初始溫度
                //double Rt = 0.99; //降溫速率
                //Dictionary<double, DateTime> totalresult = new Dictionary<double, DateTime>();
                //var tempseq = originseq.ToList();
                //double fitness1, fitness2;
                //var schedule1 = PMC.Scheduled(originseq);//插入插單後的初始排程
                //var firstresult = PMC.CalIniFitness(originsche, schedule1, activatetime);//算插入後與插入前的排程結果差異
                //fitness1 = PMC.DispatchFitness(originsche, schedule1, firstresult, activatetime);
                //double T = Math.Round(fitness1, 0); //初始溫度

                //for (int i = 0; i < 1000; i++)//迭代1000次
                //{
                //    var rand = new Random(Guid.NewGuid().GetHashCode());
                //    if (i == 80)
                //    {

                //    }
                //    var temporiseq = tempseq.ToList();
                //    var criticalpath = PMC.FindCriticalPath(finalschedule);
                //    var machinelist = criticalpath.Distinct(x => x.WorkGroup).Select(x => x.WorkGroup).ToList();
                //    int j = rand.Next(0, machinelist.Count); //挑一台機台
                //    var samachine = criticalpath.Where(x => x.WorkGroup == machinelist[j]).ToList();
                //    //若該機台製程數小於2或皆為同一工單就跳過
                //    if (samachine.Count() < 2 || samachine.Distinct(x => x.OrderID).Select(x => x.OrderID).ToList().Count() == 1)
                //    {
                //        continue;
                //    }
                //    int idx1 = 0, idx2 = 0;
                //    string or1 = string.Empty, or2 = string.Empty;
                //    //被挑選要交換的製程必須屬於不同工單
                //    while (or1 == or2)
                //    {
                //        var randomized = samachine.OrderBy(item => rand.Next()).ToList();//每次選擇一機台交換兩個製程順序
                //        or1 = randomized[0].OrderID;
                //        or2 = randomized[1].OrderID;
                //        idx1 = tempseq.FindIndex(x => x.OrderID == randomized[0].OrderID && x.OPID == randomized[0].OPID);
                //        idx2 = tempseq.FindIndex(x => x.OrderID == randomized[1].OrderID && x.OPID == randomized[1].OPID);
                //    }
                //    var tempnum = tempseq[idx1].EachMachineSeq;
                //    tempseq[idx1].EachMachineSeq = tempseq[idx2].EachMachineSeq;
                //    tempseq[idx2].EachMachineSeq = tempnum;

                //    var schedule2 = PMC.Scheduled(tempseq);
                //    fitness2 = PMC.DispatchFitness(originsche, schedule2, firstresult, activatetime);

                //    if (fitness2 <= fitness1)
                //    {
                //        fitness1 = fitness2;
                //        finalschedule = schedule2.ToList();
                //    }
                //    else
                //    {
                //        var r = rand.NextDouble();
                //        if (r <= Math.Exp((fitness1 - fitness2) / T))
                //        {
                //            fitness1 = fitness2;
                //            finalschedule = schedule2.ToList();
                //        }
                //        else
                //        {
                //            tempseq = temporiseq.ToList(); //若維持原排程解，則seq不變
                //        }

                //    }
                //    T = T * Rt;
                //}

                //未更動的製程(已開工、暫停中)加入finalschedule
                var nochange = InProcessSchedule.FindAll(x => !finalschedule.Exists(y => y.OrderID == x.OrderID && y.OPID == x.OPID));
                foreach (var item in nochange)
                {
                    finalschedule.Add(item);
                }

                #region 總排程調整
                var OutsourcingList = PMC.getOutsourcings();
                foreach (var item in originSchedule)
                {
                    var orderList = finalschedule.Distinct(x => x.OrderID)
                              .Select(x => x.OrderID)
                              .ToList();

                    for (int k = 0; k < 2; k++)
                    {
                        foreach (var one_order in orderList)
                        {
                            //挑選同工單製程
                            var itt = finalschedule.Where(x => x.OrderID == one_order)
                                             .OrderBy(x => x.Range)
                                             .ToList();

                            for (int i = 1; i < itt.Count; i++)
                            {
                                int idx;

                                //調整同工單製程
                                if (DateTime.Compare(Convert.ToDateTime(itt[i - 1].EndTime), Convert.ToDateTime(itt[i].StartTime)) > 0)
                                {
                                    idx = finalschedule.FindIndex(x => x.OrderID == itt[i].OrderID && x.OPID == itt[i].OPID);
                                    finalschedule[idx].StartTime = itt[i - 1].EndTime;
                                    finalschedule[idx].EndTime = itt[i - 1].EndTime + itt[i].Duration;
                                    itt[i].StartTime = itt[i - 1].EndTime;
                                    itt[i].EndTime = itt[i - 1].EndTime + itt[i].Duration;
                                }
                                //若委外再調整同機台製程
                                if (OutsourcingList.Exists(x => x.remark == itt[i].WorkGroup))
                                {
                                    if (OutsourcingList.Where(x => x.remark == itt[i].WorkGroup).First().isOutsource == "0")
                                    {
                                        //調整同機台製程
                                        if (finalschedule.Exists(x => itt[i].WorkGroup == x.WorkGroup))
                                        {
                                            var sequence = finalschedule.Where(x => x.WorkGroup == itt[i].WorkGroup)
                                                                 .OrderBy(x => x.StartTime)
                                                                 .ToList();
                                            for (int j = 1; j < sequence.Count; j++)
                                            {
                                                if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                                {
                                                    idx = finalschedule.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                                    finalschedule[idx].StartTime = sequence[j - 1].EndTime;
                                                    finalschedule[idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                                    sequence[j].StartTime = sequence[j - 1].EndTime;
                                                    sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                #endregion

                #region 第二次調整(避免前後製程中間有空閒時間)
                var TempJob = finalschedule.OrderBy(x => x.StartTime).ToList();
                for (int k = 0; k < 2; k++)
                {
                    foreach (var item in TempJob)
                    {
                        if (requestOrderSets.Exists(x => x.OrderID == item.OrderID && x.OPID == item.OPID))
                        {
                            //同機台前面製程
                            var samewg = TempJob.FindAll(x => x.WorkGroup == item.WorkGroup && Convert.ToDateTime(x.StartTime) < Convert.ToDateTime(item.StartTime)).OrderByDescending(x => x.StartTime).ToList();
                            //同工單前面製程 
                            var sameod = TempJob.FindAll(x => x.OrderID == item.OrderID && Convert.ToInt32(x.Range) < Convert.ToInt32(item.Range)).OrderByDescending(x => x.Range).ToList();

                            TimeSpan TS = item.Duration;
                            if (samewg.Count() > 0 || sameod.Count() > 0)
                            {
                                //比較同工單前製程與同機台前製程結束時間
                                if (samewg.Count() > 0 && sameod.Count() > 0)
                                {
                                    if (Convert.ToDateTime(samewg[0].EndTime) > Convert.ToDateTime(sameod[0].EndTime))
                                    {
                                        //提前時間不能早於插單開始時間
                                        item.StartTime = activatetime >= samewg[0].EndTime ? activatetime : samewg[0].EndTime;
                                        item.EndTime = item.StartTime + TS;

                                    }
                                    else if (Convert.ToDateTime(samewg[0].EndTime) <= Convert.ToDateTime(sameod[0].EndTime))
                                    {
                                        item.StartTime = activatetime >= sameod[0].EndTime ? activatetime : sameod[0].EndTime;
                                        item.EndTime = item.StartTime + TS;
                                    }

                                }
                                else if (samewg.Count() == 0 && sameod.Count() > 0)
                                {
                                    item.StartTime = activatetime >= sameod[0].EndTime ? activatetime : sameod[0].EndTime;
                                    item.EndTime = item.StartTime + TS;
                                }
                                else if (sameod.Count() == 0 && samewg.Count() > 0)
                                {
                                    item.StartTime = activatetime >= samewg[0].EndTime ? activatetime : samewg[0].EndTime;
                                    item.EndTime = item.StartTime + TS;
                                }
                                else
                                {
                                    if (DateTime.Compare(Convert.ToDateTime(activatetime), Convert.ToDateTime(item.StartTime)) < 0)
                                    {
                                        item.StartTime = activatetime;
                                        item.EndTime = Convert.ToDateTime(item.StartTime) + TS;
                                    }
                                }



                                //更新Tep
                                var idx = finalschedule.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                                finalschedule[idx].StartTime = item.StartTime;
                                finalschedule[idx].EndTime = item.EndTime;
                            }
                        }
                    }
                }
                #endregion

                clearntemp($"DELETE FROM AssignmentTemp4");
                SqlStr = $@"INSERT INTO AssignmentTemp4 ([OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime],[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate],[Parent],[SAP_WorkGroup]
                              ,[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice],[IsSkip],[MAKTX]
                             ,[PRIORITY],[Note],[Important])
                           SELECT [OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime],[StartTime]
                              ,[EndTime],[WorkGroup],[Operator],[AssignDate],[Parent],[SAP_WorkGroup]
                              ,[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice],[IsSkip],[MAKTX]
                              ,[PRIORITY],[Note],[Important]
                              FROM Assignment";
                using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                    {

                        if (conn.State != ConnectionState.Open)
                            conn.Open();
                        int t = comm.ExecuteNonQuery();
                    }
                }
                int changecount = 0;
                foreach (var item in finalschedule)
                {
                    // 插單製程icon設為【預覽待確認】
                    if (request.Exists(x => x.OrderID == item.OrderID))
                    {
                        SqlStr = $@" update AssignmentTemp4
                                 set StartTime = @StartTime, EndTime = @EndTime, WorkGroup = @WorkGroup, Scheduled = 3
                                 where OrderID = @OrderID and OPID =  @OPID";
                        changecount += 1;
                    }
                    else
                    {
                        //非加工中、暫停中製程
                        if(originSchedule.Exists(x => x.OrderID == item.OrderID && x.OPID == item.OPID))
                        {
                            var idx = originSchedule.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                            //時間有調整製程icon設為【受影響】
                            if (originSchedule[idx].StartTime != item.StartTime || originSchedule[idx].WorkGroup != item.WorkGroup)
                            {
                                SqlStr = $@" update AssignmentTemp4
                                 set StartTime = @StartTime, EndTime = @EndTime, WorkGroup = @WorkGroup, Scheduled = 2
                                 where OrderID = @OrderID and OPID =  @OPID";
                                changecount += 1;
                            }
                            else
                            {
                                SqlStr = $@" update AssignmentTemp4
                                 set StartTime = @StartTime, EndTime = @EndTime, WorkGroup = @WorkGroup, Scheduled = 1
                                 where OrderID = @OrderID and OPID =  @OPID";
                            }
                        }
                        //加工中、暫停中製程
                        else
                        {
                            SqlStr = $@" update AssignmentTemp4
                                 set StartTime = @StartTime, EndTime = @EndTime, WorkGroup = @WorkGroup, Scheduled = 1
                                 where OrderID = @OrderID and OPID =  @OPID";
                        }
                    }

                    using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                    {
                        using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                        {
                            if (conn.State != ConnectionState.Open)
                                conn.Open();
                            comm.Parameters.Add("@OrderID", SqlDbType.VarChar).Value = item.OrderID;
                            comm.Parameters.Add("@OPID", SqlDbType.VarChar).Value = item.OPID;
                            comm.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = item.StartTime.ToString(_timeFormat);
                            comm.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = item.EndTime.ToString(_timeFormat);
                            comm.Parameters.Add("@WorkGroup", SqlDbType.NVarChar).Value = item.WorkGroup;
                            int t = comm.ExecuteNonQuery();
                        }
                    }
                }

                //result = "Data save successful";
                result = $"Total variation is {changecount.ToString()} times，change rate is {Math.Round(decimal.ToDouble(changecount) / decimal.ToDouble(totalcount), 2)}";
            }
            else
            {
                throw new Exception("No data available in the interval");
            }

            return new ActionResponse<string>
            {
                Data = result
            };
        }

        //休息時間、休假日計算
        RestTime restTime = new RestTime();

        private void ScheduleAlgorithm(EditOrderModels edit, ref List<OrderInfo> Tep)
        {
            var OutsourcingList = getOutsourcing();


            string SqlStr = $@"SELECT p.CanSync, OrderQTY, HumanOpTime, MachOpTime, AssignDate, AssignDate_PM, MAKTX, Range
                              FROM {_ConnectStr.APSDB}.dbo.Assignment
                              inner join {_ConnectStr.MRPDB}.dbo.Process as p
                              on OPID = p.ID
                              WHERE OrderID = @OrderID and OPID = @OPID";
            List<OrderInfo> MachineSeq = new List<OrderInfo>();//存同機台的所有製程
            List<OrderInfo> TempSeq = new List<OrderInfo>(); //存有調整時間的同機台製程
            var CData = new List<BasicChartWork>(); //甘特圖資訊
            OrderInfo dep = new OrderInfo //dep=edit資訊
            {
                OrderID = edit.OrderID,
                OPID = edit.OPID,
                StartTime = edit.StartTime,
                WorkGroup = edit.WorkGroup
            };
            string DateFormat = "yyyy-MM-dd HH:mm";
            int SqlCount = 0, Idx = 0;
            double HumanOpTime = 0;
            double MachOpTime = 0;
            int CanSync = 0;
            using (SqlConnection Conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand Comm = new SqlCommand(SqlStr, Conn))
                {
                    if (Conn.State != ConnectionState.Open)
                        Conn.Open();
                    Comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = edit.OrderID;
                    Comm.Parameters.Add("@OPID", SqlDbType.NVarChar).Value = edit.OPID;
                    using (SqlDataReader SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                SqlCount = Convert.ToInt32(SqlData["OrderQTY"]);
                                HumanOpTime = Convert.ToDouble(SqlData["HumanOpTime"]);
                                MachOpTime = Convert.ToDouble(SqlData["MachOpTime"]);
                                CanSync = Convert.ToInt16(SqlData["CanSync"]);
                                dep.OrderQTY = SqlCount;
                                dep.AssignDate = Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd");
                                dep.AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd");
                                dep.Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim();
                                dep.Range = Convert.ToInt32(SqlData["Range"]);
                            }
                        }
                    }
                }
            }
            TimeSpan PaddingTime = new TimeSpan(0);
            TimeSpan PostTime = new TimeSpan();
            if (CanSync==0)
            {
                PostTime = new TimeSpan(0, 0, (int)(HumanOpTime + (SqlCount * MachOpTime)), 0, 0);
            }
            else
            {
                PostTime = new TimeSpan(0, 0, (int)(HumanOpTime + MachOpTime), 0, 0);
            }
            

            //PostST(預計開工時間):
            //機台報修 - 機台預計修復完成時間
            //插單優先 - 現在時間
            //DateTime PostST = restTime.hoildaycheck(Convert.ToDateTime(dep.StartTime));
            //DateTime PostET = restTime.restTimecheck(PostST, PostTime);
            DateTime PostST = Convert.ToDateTime(dep.StartTime);
            DateTime PostET = PostST + PostTime;



            Idx = Tep.FindIndex(x => x.OrderID == dep.OrderID && x.OPID == dep.OPID); //找到要插單的該製程位置
            if (Idx != -1)//若有找到
            {
                Tep.RemoveAt(Idx); //移除assignment中該製程資料
            }

            MachineSeq = Tep.FindAll(x => x.WorkGroup == dep.WorkGroup).OrderBy(x => x.StartTime).ToList();//同機台的所有製程

            //找出同機台被影響的後面第一道製程
            Idx = MachineSeq.FindIndex(x => DateTime.Compare(PostST, Convert.ToDateTime(x.StartTime)) < 0);
            //非委外機台再調整同機台時間
            if (!OutsourcingList.Exists(x => x == dep.WorkGroup))
            {
                if (Idx != -1)
                {

                    for (int i = Idx; i < MachineSeq.Count; i++)
                    {
                        if (MachineSeq[i].CanSync==0)
                        {
                            PaddingTime = new TimeSpan(0, 0, (int)(MachineSeq[i].HumanOpTime + (MachineSeq[i].OrderQTY * MachineSeq[i].MachOpTime)), 0, 0);
                        }
                        else
                        {
                            PaddingTime = new TimeSpan(0, 0, (int)(MachineSeq[i].HumanOpTime + MachineSeq[i].MachOpTime), 0, 0);
                        }

                        if (i == Idx && DateTime.Compare(PostST, Convert.ToDateTime(MachineSeq[i].StartTime)) < 0)
                        {
                            MachineSeq[i].StartTime = PostET.ToString(DateFormat);
                            MachineSeq[i].EndTime = (Convert.ToDateTime(MachineSeq[i].StartTime) + PaddingTime).ToString(DateFormat);
                            TempSeq.Add(MachineSeq[i]); //有調整機台的同機台製程
                        }
                        //非第一道影響的製程
                        else if (i != Idx)
                        {
                            //其他道製程緊跟前一道製程時間
                            MachineSeq[i].StartTime = MachineSeq[i - 1].EndTime;
                            MachineSeq[i].EndTime = (Convert.ToDateTime(MachineSeq[i].StartTime) + PaddingTime).ToString(DateFormat);
                            TempSeq.Add(MachineSeq[i]); //有調整機台的同機台製程
                        }

                        //check是否可以跳出(1.跑完數量跳出 2.無重疊跳出)
                        if ((i + 1) == MachineSeq.Count)
                        {
                            break;
                        }
                        else if (DateTime.Compare(Convert.ToDateTime(MachineSeq[i + 1].StartTime), Convert.ToDateTime(MachineSeq[i].EndTime)) >= 0) //若當前製程結束時間與下一道製程開始時間無重疊則跳出
                        {
                            break;
                        }
                    }
                }
            }
            
            dep.StartTime = PostST.ToString(DateFormat);
            dep.EndTime = PostET.ToString(DateFormat);
            Tep.Add(dep);


            //調整同機台製程時間至Tep
            for (int i = 0; i < TempSeq.Count; i++)
            {
                Idx = Tep.FindIndex(x => x.OPID == TempSeq[i].OPID && x.OrderID == TempSeq[i].OrderID);
                Tep[Idx].StartTime = TempSeq[i].StartTime;
                Tep[Idx].EndTime = TempSeq[i].EndTime;
            }


            List<OrderInfo> TData = new List<OrderInfo>();
            var TOrder = new List<string>();

            for (int j = 0; j < 2; j++)
            {
                //調整所有工單前後製程時間
                TOrder = Tep.Distinct(x => x.OrderID).Select(x => x.OrderID).ToList(); //工單列表
                foreach (var item in TOrder)
                {
                    TData = Tep.FindAll(x => x.OrderID == item).OrderBy(x => x.Range).ToList();
                    TData = CheckOP(TData); //同工單製程調整時間
                    for (int i = 0; i < TData.Count; i++) //更新Tep
                    {
                        int index = Tep.FindIndex(x => x.SeriesID == TData[i].SeriesID);
                        Tep[index].StartTime = TData[i].StartTime;
                        Tep[index].EndTime = TData[i].EndTime;
                    }
                }

                //調整所有機台製程時間
                TOrder = Tep.Distinct(x => x.WorkGroup).Select(x => x.WorkGroup).ToList();
                foreach (var item in TOrder)
                {
                    if (!OutsourcingList.Exists(x => x == item))//不是委外的製程再檢查是否有重疊
                    {
                        TData = Tep.FindAll(x => x.WorkGroup == item).OrderBy(x => x.StartTime).ToList();
                        TData = CheckOP(TData);
                        for (int i = 0; i < TData.Count; i++)
                        {
                            int index = Tep.FindIndex(x => x.SeriesID == TData[i].SeriesID);
                            Tep[index].StartTime = TData[i].StartTime;
                            Tep[index].EndTime = TData[i].EndTime;
                        }
                    }
                }
            }

            CountDelay(Tep);//計算每個製程延遲天數
        }

        private List<string> getOutsourcing()
        {
            var result = new List<string>();
            string SqlStr = $@"SELECT b.remark,a.Outsource FROM Outsourcing as a
                                inner join Device as b on a.Id=b.ID
                                where Outsource=1";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                Conn.Open();
                using (SqlCommand Comm = new SqlCommand(SqlStr, Conn))
                {
                    using (SqlDataReader SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(SqlData["remark"].ToString());
                            }
                        }
                    }
                }
            }
            return result;
        }

        private bool IsOutsourcing(string device)
        {
            string SqlStr = $@"SELECT a.ID,a.remark,b.Outsource FROM Device as a 
                            left join Outsourcing as b on a.ID=b.Id
                            where a.remark='{device}'";
            bool result = false;
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                Conn.Open();
                using (SqlCommand Comm = new SqlCommand(SqlStr, Conn))
                {
                    using (SqlDataReader SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            SqlData.Read();
                            if (SqlData["Outsource"].ToString() == "1")
                                result = true;
                        }
                    }
                }
            }
            return result;
        }

        //---------------------------------------------------------------------------------------------------
        //以下為派單排程流程相關API

        /// <summary>
        /// 派單時 選擇該製程可選之機台列表 (O)
        /// </summary>
        /// <returns>
        /// 測資:
        /// { "orderID": "1219111325",  "opid": "30"}
        /// </returns>
        [HttpPost("MachineList")]
        public ActionResponse<List<AllMachine>> MachineList([FromBody] ScheduleOneOrderRequest request)
        {
            var result = new List<AllMachine>();
            //var SqlStr = $@"SELECT * FROM {_ConnectStr.APSDB}.[dbo].[Assignment] as a
            //                left join {_ConnectStr.MRPDB}.dbo.ProcessDetial as b on a.OPID=b.ProcessID
            //                left join {_ConnectStr.APSDB}.[dbo].Device as c on b.MachineID=c.ID
            //                where a.OrderID=@OrderID and a.OPID=@OPID";

            var SqlStr = $@"IF EXISTS(select top(1) WorkGroup from Assignment where OrderID=@OrderID and OPID=@OPID  and WorkGroup is not null)
	                        BEGIN 
		                        SELECT * FROM {_ConnectStr.APSDB}.[dbo].[Device] where remark != 
		                        (select top(1) WorkGroup from Assignment where OrderID=@OrderID and OPID=@OPID)
	                        END
                        ELSE
	                        BEGIN 
		                        SELECT * FROM {_ConnectStr.APSDB}.[dbo].[Device] where remark != ''
	                        END";

            List<string> MachineID = new List<string>();
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = request.OrderID;
                    comm.Parameters.Add("@OPID", SqlDbType.VarChar).Value = request.OPID;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {

                            while (SqlData.Read())
                            {
                                MachineID.Add(SqlData["remark"].ToString());
                            }
                        }
                    }
                }
            }

            var output = new List<AllMachine>();
            output.Add(new AllMachine
            {
                Machine = MachineID
            });
            return new ActionResponse<List<AllMachine>>
            {
                Data = output
            };
        }

        //---------------------------------------------------------------------------------------------------
        //以下為機台報修相關API
        #region 機台報修相關API

        /// <summary>
        /// 取得報修機台資訊 (O)
        /// </summary>
        /// <param name="Device"></param>
        /// <returns></returns>
        [HttpGet("FixedInfo/{Device}")]
        public ActionResponse<List<FixedInfo>> FixedInfo(string Device)
        {
            if (Device.Length <= 0)
            {
                throw new Exception("未輸入機台編號");
            }

            var resutl = new List<FixedInfo>();

            var sqlStr = $@"SELECT * FROM DeviceBreakdownInfo
                            WHERE MachineName=@MachineId AND IsFixed=0";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.Add("@MachineId", SqlDbType.NVarChar).Value = Device;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        while (SqlData.Read())
                        {
                            if (SqlData.HasRows)
                            {
                                resutl.Add(new FixedInfo
                                {
                                    Device = SqlData["MachineId"].ToString(),
                                    STartTime = Convert.ToDateTime(SqlData["BreakdownST"].ToString()).ToString(_timeFormat),
                                    EndTime = Convert.ToDateTime(SqlData["BreakdownET"].ToString()).ToString(_timeFormat)
                                });
                            }
                            else
                            {
                                throw new Exception("查無資料");
                            }
                        }
                    }
                }
            }


            return new ActionResponse<List<FixedInfo>>
            {
                Data = resutl
            };
        }

        /// <summary>
        /// 設備故障排程法 (O)*
        /// </summary>
        /// <param name="request">{  "machine": "2M301",  "mode": "A",  "startTime": "2021-01-02 10:00:00",  "endTime": "2021-01-02 13:00:00"}</param>
        /// <returns></returns>
        [HttpPost("BreakdownGanttChart")]
        public ActionResponse<List<Schedulelist_MB>> BreakdownGanttChart([FromBody] MachineBreakdownRequest request)
        {
            var result = new List<Schedulelist_MB>();
            var _temp = new List<OrderInfo>();
            var tep = new List<OrderInfo>();
            var rnd = new Random();
            var breakdowninfo = new MachinebreakdowInfo();
            breakdowninfo.Machine = request.Machine;
            breakdowninfo.StartTime = request.StartTime;
            breakdowninfo.EndTime = request.EndTime;
            var db = @"AssignmentTemp";
            if (request.Mode == "A")
                db += "6";
            else
                db += "7";
            ////輩分assigment至assigmentTemp05，寫入故障機台資訊至DeviceBreakdownInfo
            var sqlStr = "";
            //var sqlStr = $@"
            //                DELETE {db}

            //                INSERT INTO {db} ([OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
            //                      ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
            //                      ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
            //                      ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important])
            //                SELECT [OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
            //                      ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
            //                      ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
            //                      ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important]
            //                FROM Assignment

            //                IF EXISTS(SELECT *  FROM DeviceBreakdownInfoTemp WHERE MachineName=@MachineName AND IsFixed=0)
            //                    BEGIN
            //                        UPDATE DeviceBreakdownInfoTemp SET IsFixed=1,Lastupdatetime=GETDATE() WHERE MachineName=@MachineName AND IsFixed=0
            //                    END
            //                ELSE    
            //                    BEGIN
            //                        INSERT INTO DeviceBreakdownInfoTemp VALUES ((select ID from Device where remark=@MachineName),@MachineName,@MB_ST,@MB_ET,0,GETDATE(),GETDATE())
            //                    END

            //                SELECT  a.*, w.WIPEvent,
            //                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0)
            //                FROM AssignmentTemp6 as a
            //                LEFT JOIN WIP as w ON (a.OrderID = w.OrderID AND a.OPID = w.OPID)
            //                ";
            //using (var Conn = new SqlConnection(_ConnectStr.Local))
            //{
            //    if (Conn.State != ConnectionState.Open)
            //        Conn.Open();
            //    using (var Comm = new SqlCommand(sqlStr, Conn))
            //    {
            //        Comm.Parameters.AddWithValue("@MachineName", request.Machine);
            //        Comm.Parameters.AddWithValue("@MB_ST", request.StartTime);
            //        Comm.Parameters.AddWithValue("@MB_ET", request.EndTime);
            //        using (var SqlData = Comm.ExecuteReader())
            //        {
            //            if (SqlData.HasRows)
            //            {
            //                //while (SqlData.Read())
            //                //{
            //                //    _temp.Add(new OrderInfo
            //                //    {
            //                //        OrderID = SqlData["OrderID"].ToString().Trim(),
            //                //        OPID = Convert.ToDouble(SqlData["OPID"]),
            //                //        StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
            //                //        OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
            //                //        EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
            //                //        WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
            //                //        AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
            //                //        AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
            //                //        Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim(),
            //                //        HumanOpTime = Convert.ToDouble(SqlData["HumanOpTime"]),
            //                //        MachOpTime = Convert.ToDouble(SqlData["MachOpTime"]),
            //                //        Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString()),
            //                //        Note = _checkNoword(SqlData["Note"].ToString()),
            //                //        Important = Convert.ToBoolean(SqlData["Important"].ToString()),
            //                //        Assign = Convert.ToBoolean(SqlData["Scheduled"].ToString()),
            //                //        PRIORITY = Convert.ToInt32(SqlData["PRIORITY"]),
            //                //        OPLTXA1 = SqlData["OPLTXA1"].ToString().Trim()
            //                //    });
            //                //}
            //            }
            //        }
            //    }
            //}

            string SqlStr = @$"SELECT *,p.CanSync
                              FROM  {_ConnectStr.APSDB}.dbo.Assignment
                              INNER JOIN {_ConnectStr.MRPDB}.dbo.Process as p
                              on OPID = p.ID
                              where Scheduled = 1";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    //取得工單列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                //讀取工單資料
                                tep.Add(new OrderInfo
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim(),
                                    CanSync = Convert.ToInt16(SqlData["CanSync"])
                                });
                            }
                        }
                    }
                }
            }

            //尋找受影響工單
            var fixStart = Convert.ToDateTime(request.StartTime);
            var fixEnd = Convert.ToDateTime(request.EndTime);
            sqlStr = @"select * from Assignment as a
                       inner join WIP as b on a.OrderID=b.OrderID and a.OPID=b.OPID
                       where a.Scheduled = 1 and a.WorkGroup =  @machine and b.WIPEvent=0
                       and ((a.StartTime > @st and  a.StartTime < @et) or (a.EndTime > @st and  a.EndTime < @et))  
                       order by a.StartTime
                       ";
            var fixOrders = new List<OrderInfo>();
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(sqlStr, Conn))
                {
                    Comm.Parameters.AddWithValue("@st", fixStart);
                    Comm.Parameters.AddWithValue("@et", fixEnd);
                    Comm.Parameters.AddWithValue("@machine", request.Machine);
                    //取得工單列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                //讀取工單資料
                                fixOrders.Add(new OrderInfo
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim()
                                });
                            }
                        }
                    }
                }
            }


            var dbTable = @"AssignmentTemp";
            if (request.Mode == "A")
                dbTable += "6";
            else
                dbTable += "7";
            sqlStr = $@"
                            DELETE {dbTable}

                            INSERT INTO  {dbTable} ([OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
                                  ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
                                  ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
                                  ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important])
                            SELECT [OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
                                  ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
                                  ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
                                  ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important]
                            FROM Assignment

                            IF EXISTS(SELECT *  FROM DeviceBreakdownInfoTemp WHERE MachineName=@MachineName AND IsFixed=0)
                                BEGIN
                                    UPDATE DeviceBreakdownInfoTemp SET IsFixed=1,Lastupdatetime=GETDATE() WHERE MachineName=@MachineName AND IsFixed=0
                                END
                            ELSE    
                                BEGIN
                                    INSERT INTO DeviceBreakdownInfoTemp VALUES ((select ID from Device where remark=@MachineName),@MachineName,@MB_ST,@MB_ET,0,GETDATE(),GETDATE())
                                END

                            SELECT  a.*, w.WIPEvent,
                            Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0)
                            FROM AssignmentTemp6 as a
                            LEFT JOIN WIP as w ON (a.OrderID = w.OrderID AND a.OPID = w.OPID)
                            ";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(sqlStr, Conn))
                {
                    Comm.Parameters.AddWithValue("@MachineName", request.Machine);
                    Comm.Parameters.AddWithValue("@MB_ST", request.StartTime);
                    Comm.Parameters.AddWithValue("@MB_ET", request.EndTime);
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                _temp.Add(new OrderInfo
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim(),
                                    HumanOpTime = Convert.ToDouble(SqlData["HumanOpTime"]),
                                    MachOpTime = Convert.ToDouble(SqlData["MachOpTime"]),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString()),
                                    Note = _checkNoword(SqlData["Note"].ToString()),
                                    Important = Convert.ToBoolean(SqlData["Important"].ToString()),
                                    //Assign = Convert.ToBoolean(SqlData["Scheduled"].ToString()),
                                    PRIORITY = Convert.ToInt32(SqlData["PRIORITY"]),
                                    OPLTXA1 = SqlData["OPLTXA1"].ToString().Trim()
                                });
                            }
                        }
                    }
                }
            }
            //}

            foreach (var item in tep.Where(x => x.WorkGroup == request.Machine && DateTime.Compare(Convert.ToDateTime(x.StartTime), fixStart) <= 0).ToList())
            {
                if (DateTime.Compare(Convert.ToDateTime(item.StartTime), fixEnd) > 0)
                {
                    //padding = Convert.ToDateTime()
                }
            }

            var works = tep.Where(x => x.WorkGroup != request.Machine)
                           .Distinct(x => x.WorkGroup)
                           .Select(x => x.WorkGroup)
                           .ToList();

            switch (request.Mode)
            {
                case "A":
                    foreach (var order in fixOrders)
                    {
                        var editOrder = new EditOrderModels
                        {
                            OrderID = order.OrderID,
                            OPID = order.OPID,
                            StartTime = request.EndTime,
                            WorkGroup = works[rnd.Next(0, works.Count)]
                        };
                        ScheduleAlgorithm(editOrder, ref tep);
                    }
                    break;
                case "B":
                    foreach (var order in fixOrders)
                    {
                        var editOrder = new EditOrderModels
                        {
                            OrderID = order.OrderID,
                            OPID = order.OPID,
                            StartTime = request.EndTime,
                            WorkGroup = request.Machine
                        };
                        ScheduleAlgorithm(editOrder, ref tep);
                    }
                    break;
                default:
                    break;
            }

            var temp = new List<Schedule>();
            SqlStr = @"SELECT * FROM  Assignment where Scheduled=1 ORDER BY OrderID, OPID ";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    //取得工單列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                //讀取工單資料
                                _temp.Add(new OrderInfo
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim()
                                });
                            }
                        }
                    }
                }
            }



            foreach (var item in tep)
            {
                if (!_temp.Exists(x => x.OrderID == item.OrderID && x.OPID == item.OPID &&
                x.WorkGroup == item.WorkGroup &&
                x.StartTime == item.StartTime && x.EndTime == item.EndTime))
                {
                    sqlStr = $@" UPDATE {db} SET WorkGroup=@WorkGroup,StartTime=@StartTime,EndTime=@EndTime,Scheduled=2
                                WHERE OrderID=@OrderID AND OPID=@OPID";
                    using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                    {

                        if (conn.State != ConnectionState.Open)
                            conn.Open();

                        using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                        {
                            comm.Parameters.Add("@WorkGroup", SqlDbType.NVarChar).Value = tep.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).WorkGroup;
                            comm.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = tep.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).StartTime;
                            comm.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = tep.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).EndTime;
                            comm.Parameters.Add("@OrderID", SqlDbType.VarChar).Value = item.OrderID;
                            comm.Parameters.Add("@OPID", SqlDbType.VarChar).Value = item.OPID;

                            int impactrow = comm.ExecuteNonQuery();
                        }
                    }
                }
            }



            result.Add(new Schedulelist_MB
            {
                mode = request.Mode == "A" ? "6" : "7",
                breakdownInfo = breakdowninfo,
                schedules = temp
            });

            return new ActionResponse<List<Schedulelist_MB>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 機台報修延後排程甘特圖預覽
        /// </summary>
        /// <returns></returns>
        [HttpGet("BasicChartBreakdownPreview")]
        public ActionResponse<List<DelayScheduleOP>> BasicChartBreakdownPreview()
        {

            var result = new List<DelayScheduleOP>();
            var SqlStr = $@"UPDATE DeviceBreakdownInfoTemp SET isFixed=1 WHERE BreakdownET<GETDATE()

                            SELECT MachineName, BreakdownST, BreakdownET
                            FROM DeviceBreakdownInfoTemp
                            WHERE BreakdownET>=GETDATE() AND IsFixed=0";

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
                                result.Add(new DelayScheduleOP
                                {
                                    MachineName = SqlData["MachineName"].ToString().Trim(),
                                    BreakdownST = _ChangeTimeFormat(checkNoword(SqlData["BreakdownST"].ToString())),
                                    BreakdownET = _ChangeTimeFormat(checkNoword(SqlData["BreakdownET"].ToString())),
                                });
                            }
                        }
                    }
                }
            }


            return new ActionResponse<List<DelayScheduleOP>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 機台報修延後排程甘特圖
        /// </summary>
        /// <returns></returns>
        [HttpGet("BasicChartBreakdown")]
        public ActionResponse<List<DelayScheduleOP>> BasicChartBreakdown()
        {

            var result = new List<DelayScheduleOP>();


            var SqlStr = $@"UPDATE DeviceBreakdownInfo SET isFixed=1 WHERE BreakdownET<GETDATE()-5

                            SELECT MachineName, BreakdownST, BreakdownET
                            FROM DeviceBreakdownInfo
                            WHERE BreakdownET>=GETDATE() AND IsFixed=0";

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
                                result.Add(new DelayScheduleOP
                                {
                                    MachineName = SqlData["MachineName"].ToString().Trim(),
                                    BreakdownST = _ChangeTimeFormat(checkNoword(SqlData["BreakdownST"].ToString())),
                                    BreakdownET = _ChangeTimeFormat(checkNoword(SqlData["BreakdownET"].ToString())),
                                });

                            }
                        }
                    }
                }
            }


            return new ActionResponse<List<DelayScheduleOP>>
            {
                Data = result
            };
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

        private string _checkNoword(string data)
        {
            if (data == "")
            {
                return "－";
            }
            else
            {
                return data;
            }
        }

        #endregion



        #region 手動排程 相關API

        /// <summary>
        /// 手動排程 (O)*
        /// </summary>
        /// <param name="request">工單編號、製程編號、機台編號、開始時間</param>
        /// <returns>
        /// {
        ///   "orderID": "1219111327",
        ///   "opid": "30",
        ///   "machine": "2M301",
        ///   "startTime": "2021-12-04 16:19:00"
        /// }
        /// </returns>
        /// <summary>
        /// 手動排程 (O)*
        /// </summary>
        /// <param name="request">工單編號、製程編號、機台編號、開始時間</param>
        /// <returns>
        /// {
        ///   "orderID": "1219111327",
        ///   "opid": "30",
        ///   "machine": "2M301",
        ///   "startTime": "2021-12-04 16:19:00"
        /// }
        /// </returns>
        [HttpPost("ManualSchedule")]
        public ActionResponse<string> ManualSchedule([FromBody] ManualScheduleRequest request)
        {
            var result = "Data save failed";
            var Tep = new List<OrderInfo>();
            var temp = new List<Schedule>();
            var _temp = new List<OrderInfo>();
            var originSchedule = new List<ScheduleDto>();
            var changeOrders = new EditOrderModels();

            string SqlStr = $@"SELECT *, p.CanSync
                              FROM  {_ConnectStr.APSDB}.dbo.Assignment as a
                              left join {_ConnectStr.APSDB}.dbo.WIP as wip
                              on a.OrderID=wip.OrderID and a.OPID=wip.OPID
                              INNER JOIN {_ConnectStr.MRPDB}.dbo.Process as p on a.OPID = p.ID
                              where (wip.WIPEvent!=3 or wip.WIPEvent is NULL) and (Scheduled = 1 and a.StartTime is not null) or (a.OrderID = @orderID and a.OPID = @OPID)
                              ORDER BY a.OrderID, a.OPID";
            //string SqlStr = $@"SELECT *, p.CanSync
            //                  FROM  {_ConnectStr.APSDB}.dbo.Assignment
            //                  INNER JOIN {_ConnectStr.MRPDB}.dbo.Process as p on OPID = p.ID
            //                  where (Scheduled = 1 or (OrderID = @orderID and OPID = @OPID))
            //                  ORDER BY OrderID, OPID";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    Comm.Parameters.AddWithValue("@orderID", request.OrderID);
                    Comm.Parameters.AddWithValue("@StartTime", request.StartTime);
                    Comm.Parameters.AddWithValue("@OPID", request.OPID);
                    //取得工單列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                //讀取工單資料
                                Tep.Add(new OrderInfo
                                {
                                    SeriesID = SqlData["SeriesID"].ToString().Trim(),
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    Range = Convert.ToInt32(SqlData["Range"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    HumanOpTime = Convert.ToDouble(SqlData["HumanOpTime"].ToString()),
                                    MachOpTime = Convert.ToDouble(SqlData["MachOpTime"].ToString()),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim(),
                                    CanSync = Convert.ToInt16(SqlData["CanSync"])
                                });
                            }
                        }
                    }
                }
            }

            if (string.IsNullOrEmpty(request.StartTime))
            {
                var st = Tep.Where(x => x.WorkGroup == request.Machine).ToList();
                if (st.Count != 0)
                {
                    request.StartTime = st.Max(x => Convert.ToDateTime(x.EndTime)).ToString("yyyy-MM-dd HH:mm");
                }
                else
                {
                    request.StartTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                }
            }

            var edit = new EditOrderModels
            {
                OrderID = request.OrderID,
                OPID = Convert.ToDouble(request.OPID),
                StartTime = request.StartTime,
                WorkGroup = request.Machine
            };

            ScheduleAlgorithm(edit, ref Tep);

            clearntemp("truncate table AssignmentTemp5");
            var sqlStr = $@"                          
                            INSERT INTO {_ConnectStr.APSDB}.dbo.AssignmentTemp5 ([SeriesID],[OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
                                  ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
                                  ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
                                  ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important])
                            SELECT [SeriesID],[OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
                                  ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
                                  ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
                                  ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important]
                            FROM {_ConnectStr.APSDB}.dbo.Assignment
                            ";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(sqlStr, Conn))
                {
                    var t = Comm.ExecuteNonQuery();
                }
            }

            SqlStr = $@"SELECT * FROM  {_ConnectStr.APSDB}.dbo.Assignment where Scheduled=1 ORDER BY OrderID, OPID ";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    //取得工單列表
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                //讀取工單資料
                                _temp.Add(new OrderInfo
                                {
                                    SeriesID = SqlData["SeriesID"].ToString(),
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString(_timeFormat),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString(_timeFormat),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim()
                                });
                            }
                        }
                    }
                }
            }



            foreach (var item in Tep)
            {
                if (!_temp.Exists(x => x.OrderID == item.OrderID &&
                x.OPID == item.OPID &&
                x.WorkGroup == item.WorkGroup &&
                x.StartTime == item.StartTime &&
                x.EndTime == item.EndTime))
                {
                    sqlStr = $@" UPDATE {_ConnectStr.APSDB}.dbo.AssignmentTemp5 SET 
                                WorkGroup=@WorkGroup,
                                StartTime=@StartTime,
                                EndTime=@EndTime,
                                Scheduled=2
                                WHERE OrderID=@OrderID AND OPID=@OPID";
                    using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
                    {

                        if (conn.State != ConnectionState.Open)
                            conn.Open();

                        using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                        {
                            comm.Parameters.Add("@WorkGroup", SqlDbType.NVarChar).Value = Tep.Find(x => x.SeriesID == item.SeriesID).WorkGroup;
                            comm.Parameters.Add("@StartTime", SqlDbType.DateTime).Value = Tep.Find(x => x.SeriesID == item.SeriesID).StartTime;
                            comm.Parameters.Add("@EndTime", SqlDbType.DateTime).Value = Tep.Find(x => x.SeriesID == item.SeriesID).EndTime;
                            comm.Parameters.Add("@OrderID", SqlDbType.VarChar).Value = item.OrderID;
                            comm.Parameters.Add("@OPID", SqlDbType.VarChar).Value = item.OPID;

                            int impactrow = comm.ExecuteNonQuery();
                            if (impactrow > 0)
                            {
                                result = "Data save Successful";
                            }
                            else
                            {
                                throw new Exception("Date save failed");
                            }

                        }
                    }
                }
            }
            return new ActionResponse<string>
            {
                Data = result
            };
        }
        #endregion


        /// <summary>
        /// 取得機台狀態
        /// </summary>
        /// <returns></returns>
        [HttpGet("MachieStatus")]
        public ActionResponse<List<MachineState>> MachieStatus()
        {
            var result = new List<MachineState>();


            string[] state = { "NORMAL", "ALARM", "OFF" };
            Random r = new Random();
            var sqlStr = $@"select bb.[MachineName],bb.[remark] from (
                            select DISTINCT b.WorkGroup from Device as a
                            inner join Assignment as b on a.remark=b.WorkGroup) as aa
                            inner join Device as bb on aa.WorkGroup=bb.remark
                            inner join Outsourcing as cc on bb.ID=cc.Id"
                            //+"where cc.Outsource=0"
                            ;
            sqlStr = $@"select * from Device where ID<=48 order by ID";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(sqlStr, Conn))
                {
                    using (var SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new MachineState
                                {
                                    machineName = SqlData["MachineName"].ToString().Trim(),
                                    status = state[0],
                                    displayName = SqlData["remark"].ToString().Trim(),
                                });
                            }

                        }
                    }
                }
            }

            if (result.Count == 0)
            {
                sqlStr = $@"select * from Device";
                using (var Conn = new SqlConnection(_ConnectStr.Local))
                {
                    if (Conn.State != ConnectionState.Open)
                        Conn.Open();
                    using (var Comm = new SqlCommand(sqlStr, Conn))
                    {
                        using (var SqlData = Comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                while (SqlData.Read())
                                {
                                    result.Add(new MachineState
                                    {
                                        machineName = SqlData["MachineName"].ToString().Trim(),
                                        status = state[0],
                                        displayName = SqlData["remark"].ToString().Trim(),
                                    });
                                }

                            }
                        }
                    }
                }
            }

            //取得故障機台
            List<FixedInfo> fixedInfos = getbreakdownmachine();

            //台設備故障
            foreach (var item in fixedInfos)
            {
                if (result.Exists(x => x.displayName == item.Device))
                {
                    result.Find(x => x.displayName == item.Device).status = state[1];
                }
            }
            return new ActionResponse<List<MachineState>> { Data = result.OrderBy(x => x.displayName).ToList() };
        }

        private List<FixedInfo> getbreakdownmachine()
        {
            var resutl = new List<FixedInfo>();

            var sqlStr = $@"SELECT * FROM DeviceBreakdownInfo
                            WHERE IsFixed=0";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        while (SqlData.Read())
                        {
                            if (SqlData.HasRows)
                            {
                                resutl.Add(new FixedInfo
                                {
                                    Device = SqlData["MachineName"].ToString().Trim(),
                                    EndTime = SqlData["BreakdownST"].ToString().Trim(),
                                    STartTime = SqlData["BreakdownET"].ToString().Trim()

                                });
                            }
                            else
                            {
                                throw new Exception("查無資料");
                            }
                        }
                    }
                }
            }

            return resutl;
        }

        private ActionResponse<List<Material>> Get()
        {
            var result = new List<Material>();
            var SqlStr = $@"select ass.OrderID, ass.OPID,　ass.MAKTX, ass.OPLTXA1,max(w.CreateTime) as LastProduceTime
                            from Assignment as ass left join WIPLog as w
                            on ass.OrderID = w.OrderID and ass.OPID=w.OPID
                            group by ass.OrderID, ass.OPID,ass.MAKTX, ass.OPLTXA1
                            order by LastProduceTime";
            var cpk = new string[] { "合格", "不合格" };

            Random rnd = new Random(Guid.NewGuid().GetHashCode());


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
                                result.Add(new Material
                                {
                                    No = SqlData["MAKTX"].ToString().Trim(),
                                    ProcessNo = SqlData["OPID"].ToString().Trim(),
                                    CPKEvaluation = cpk[rnd.Next(cpk.Length)],
                                    ProcessName = SqlData["OPLTXA1"].ToString().Trim(),
                                    LastProduction = string.IsNullOrEmpty(SqlData["LastProduceTime"].ToString()) ? "N/A" : DateTime.Parse(SqlData["LastProduceTime"].ToString()).ToString(_timeFormat)
                                });
                            }
                        }
                    }
                }
            }
            return new ActionResponse<List<Material>>
            {
                Data = result
            };
        }
    }


    public class MachineState
    {
        /// <summary>
        /// 機台名稱
        /// </summary>
        public string machineName { get; set; }
        /// <summary>
        /// 機台狀態【NORMAL:正常 ALARM:異常 OFF:關機】
        /// </summary>
        public string status { get; set; }
        /// <summary>
        /// 機台顯示名稱
        /// </summary>
        public string displayName { get; set; }
    }


    public class FixedInfo
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 開始維修時間
        /// </summary>
        public string STartTime { get; set; }
        /// <summary>
        /// 完成維修時間
        /// </summary>
        public string EndTime { get; set; }
    }






    public class ScheduleCode
    {
        /// <summary>
        /// 排程結果代號
        /// </summary>
        public string mode { get; set; }
        /// <summary>
        /// 排程結果名稱
        /// </summary>
        public string ScheduleMethod { get; set; }
    }

    public class CommuterTime
    {
        /// <summary>
        /// 上班時間
        /// </summary>
        public string StartTime { get; set; }
        /// <summary>
        /// 下班時間
        /// </summary>
        public string CloseTime { get; set; }
    }



    public class DisplayRange
    {
        public string startTime { get; set; }
        public string endTime { get; set; }
    }



}

