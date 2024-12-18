using Microsoft.AspNetCore.Mvc;
using PMCDash.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using static PMCDash.Services.AccountService;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace PMCDash.Controllers.Part2
{
    [Route("api/[controller]")]
    [ApiController]
    public class WorkIidleResonReportController : BaseApiController
    {

        ConnectStr _ConnectStr = new ConnectStr();
        //資料庫連線
        //private readonly string _ConnectStr.Local = @"Data Source = 127.0.0.1; Initial Catalog = DPI; User ID = MES2014; Password = PMCMES;";
        //時間格式
        private readonly string _timeFormat = "yyyy-MM-dd HH:mm:ss";

        private UserData UserInfo()
        {
            UserData user = new UserData();

            string sqlStr = @$"SELECT aa.user_id,aa.user_account,aa.user_name,aa.usergroup_id,bb.usergroup_name,bb.DeviceGroup FROM (
                                SELECT user_id, user_account, user_name, usergroup_id
                                FROM {_ConnectStr.AccuntDB}.[dbo].[User])as aa
                                LEFT JOIN  {_ConnectStr.AccuntDB}.[dbo].[Units] as bb
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
                            user.EmpolyeeAccount = sqldata["user_account"].ToString();
                            user.EmpolyeeName = sqldata["user_name"].ToString();
                            user.GroupId = Convert.ToInt64(sqldata["usergroup_id"]);
                            user.GroupName = sqldata["usergroup_name"].ToString();
                            user.DeviceGroupId = sqldata["DeviceGroup"].ToString();
                        }
                    }
                }
            }
            return user;
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

        /// <summary>
        /// 取得待機紀錄清單
        /// </summary>
        [HttpPost("Workidlereoprtlist")]
        public List<idleList> Workidlereoprtlist()
        {
            var result = new List<idleList>();
            UserData userData = UserInfo();
            string SqlStr = @$"SELECT TOP (30) a.*,b.item,b.Category,c.[user_name] FROM {_ConnectStr.APSDB}.[dbo].[IdleReasonBinding] as a
                                LEFT JOIN {_ConnectStr.APSDB}.[dbo].[IdleResult] as b on a.ReasonCode=b.ID
                                LEFT JOIN {_ConnectStr.AccuntDB}.[dbo].[User] as c on a.StaffID=c.[user_id]
                                WHERE a.StaffID=@StaffID
                                ORDER BY a.ID DESC";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@StaffID", SqlDbType.VarChar).Value = userData.User_Id;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new idleList
                                {
                                    //idletype = changeIdletypecode(SqlData["Category"].ToString()),
                                    //idlereson_id = SqlData["ReasonCode"].ToString(),
                                    //idlereson_name = SqlData["item"].ToString(),
                                    //idle_start_time = checkNoword(SqlData["StartTime"].ToString()),
                                    //idle_end_time = checkNoword(SqlData["EndTime"].ToString()),
                                    //idle_first_report_time = checkNoword(DateTime.Now.ToString(_timeFormat)),
                                    //idle_scend_report_time = checkNoword("")
                                    orderID = changeIdletypecode(SqlData["OrderID"].ToString().Trim()),
                                    opId = changeIdletypecode(SqlData["OPID"].ToString().Trim()),
                                    idletype = changeIdletypecode(SqlData["Category"].ToString().Trim()),
                                    idleDevice = SqlData["Device"].ToString().Trim(),
                                    idlePerson = SqlData["user_name"].ToString().Trim(),
                                    idlereson_id = SqlData["ReasonCode"].ToString().Trim(),
                                    idlereson_name = SqlData["item"].ToString().Trim(),
                                    idle_start_time = checkNoword(SqlData["StartTime"].ToString().Trim()),
                                    idle_end_time = checkNoword(SqlData["EndTime"].ToString().Trim()),
                                    idleDuration = deleyday(SqlData["StartTime"].ToString().Trim(), SqlData["EndTime"].ToString().Trim())
                                });
                            }
                        }
                    }
                }
            }

            //result = result.OrderBy(x => x.idle_start_time).ToList();

            return result;
        }

        private string deleyday(string satrttime, string endtime)
        {
            string result = "N/A";
            DateTime sttDate, ettDate;
            if (DateTime.TryParse(satrttime, out sttDate) && DateTime.TryParse(endtime, out ettDate))
            {
                var ts = new TimeSpan(ettDate.Ticks - sttDate.Ticks);
                result = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00") + ":" + ts.Seconds.ToString("00");
                if (ts.Ticks < 0) result = "00:00:00";
            }
            else
            {
                result = "N/A";
            }
            return result;
        }

        private string changeIdletypecode(string code)
        {
            string result = code;
            if (code != null)
            {
                switch (code)
                {
                    case "0":
                        result = "N/A";
                        break;
                    case "1":
                        result = "人員";
                        break;
                    case "2":
                        result = "機台";
                        break;
                }
                return result;
            }
            else
            {
                result = "N/A";
            }
            return result;
        }

        /// <summary>
        /// 取得人員/機台待機原因清單
        /// </summary>
        /// <returns>"type":1(人員)、2(機台)</returns>
        [HttpPost("Idlereasonlist")]
        public List<idlereason> Idlereasonlist(string Category)
        {
            var result = new List<idlereason>();
            string SqlStr = @$"SELECT * FROM {_ConnectStr.APSDB}.[dbo].[IdleResult]
                                WHERE Category=@Category
                                ORDER BY id DESC";
            //如果沒指定就先顯示人員待機原因
            if (Category == null || !Category.All(char.IsDigit))
            {
                Category = "1";
            }
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@Category", SqlDbType.Int).Value = Category;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new idlereason
                                {
                                    id = SqlData["ID"].ToString(),
                                    type = SqlData["Category"].ToString(),
                                    typeName = changeIdletypecode(SqlData["Category"].ToString()),
                                    name = checkNoword(SqlData["item"].ToString()),
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }






        public class idlereason
        {
            /// <summary>
            /// 待機原因標號
            /// </summary>
            public string id { get; set; }
            /// <summary>
            /// 待機原因種類【1:人員、2:機台】
            /// </summary>
            public string type { get; set; }
            /// <summary>
            /// 待機原因名稱
            /// </summary>
            public string typeName { get; set; }
            /// <summary>
            /// 待機原因名稱
            /// </summary>
            public string name { get; set; }

            //public string Id { get; set; }           
            //public string Type { get; set; }            
            //public string idleReasonTitle { get; set; }
        }

        //public class idlereportdata
        //{
        //    public string OrderId { get; set; }
        //    public string OPID { get; set; }
        //    public string idletype { get; set; }
        //    public string Starttime { get; set; }
        //    public string Endtime { get; set; }
        //    public string ReasonCode { get; set; }
        //    public string ReasonName { get; set; }
        //    public string StaffID { get; set; }
        //    public string Updated { get; set; }
        //}

        //public class StartIdleReportRequest
        //{
        //    /// <summary>
        //    /// 工單編號
        //    /// </summary>
        //    public string OrderId { get; set; }
        //    /// <summary>
        //    /// 製程編號
        //    /// </summary>
        //    public string OPID { get; set; }
        //    /// <summary>
        //    /// 設備編號
        //    /// </summary>
        //    public string Device { get; set; }
        //    /// <summary>
        //    /// 待機原因編號
        //    /// </summary>
        //    public string ReasonCode { get; set; }
        //}

        //public class EndIdleReportRequest
        //{
        //    /// <summary>
        //    /// 待機原因編號
        //    /// </summary>
        //    public string ReasonCode { get; set; }
        //    /// <summary>
        //    /// 待機原因開始時間
        //    /// </summary>
        //    public DateTime StartTime { get; set; }
        //}
    }
}
