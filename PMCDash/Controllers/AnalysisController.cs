using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PMCDash.Services;
using PMCDash.Models;
using PMCDash.Models.Part2;
using Microsoft.AspNetCore.Authorization;
using System.Data.SqlTypes;

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AnalysisController : BaseApiController// ControllerBase
    {
        //DPI
        //soco
        //SkyMarsDB


        ConnectStr _ConnectStr = new ConnectStr();
        //private readonly string _ConnectStr.Local = @"Data Source = 127.0.0.1; Initial Catalog = DPI; User ID = MES2014; Password = PMCMES;";
        //private readonly string _ConnectStr.Local = @"Data Source = 192.168.0.156; Initial Catalog = DPI ;User ID = MES2014; Password = PMCMES;";
        private readonly string _timeFormat = "yyyy-MM-dd HH:mm:ss";
        private readonly string _timeFormat_ToMin = "yyyy-MM-dd HH:mm";
        ///// <summary>
        ///// 甘特圖標籤【0:交期優先-推薦權重、1:交期優先-自訂權重、2:機台優先、3:插單憂先、4:手動排程、5:派單排程、6:設備故障排程、7:生產延遲排程】
        ///// </summary>
        //private static List<string> TabList = new List<string> { };

        //private List<OrderInfo> Tep = new List<OrderInfo>();//工單列表

        [NonAction]
        private Tuple<List<string>, List<string>> OrderDefault()
        {
            var OrderID = new List<string>();
            var Order = Tuple.Create(OrderID, OrderID);
            return Order;
        }



        //---------------------------------------------------------------------------------------------------
        //以下為機台待機原因/時間區塊相關API
        #region 機台待機原因/時間區塊相關API
        /// <summary>
        /// 取得【單日待機時間區塊圖】圖表資訊 (O)*
        /// </summary>
        /// <returns>"type":1(人員)、2(機台)</returns>
        [HttpGet("Machine/IdleReason")]
        public ActionResponse<List<IdleChartInfo>> IdleReason()
        {
            var result = new List<IdleChartInfo>();
            var PersonResult = new List<PersonIdle>();
            var MachineResult = new List<MachineIdle>();
            var MachineStatus = new List<analysisMachineStatus>();
            MachineStatus.Add(new analysisMachineStatus { code = "OFF", DisplayName = "關機" });
            MachineStatus.Add(new analysisMachineStatus { code = "RUN", DisplayName = "運轉" });
            MachineStatus.Add(new analysisMachineStatus { code = "IDLE", DisplayName = "待機" });
            MachineStatus.Add(new analysisMachineStatus { code = "ALARM", DisplayName = "警報" });
            var IdleType = new List<analysisIdleType>();
            IdleType.Add(new analysisIdleType { code = "1", DisplayName = "人員" });
            IdleType.Add(new analysisIdleType { code = "2", DisplayName = "機台" });


            string SqlStr = @"SELECT * FROM  [IdleResult]　where exist=1　ORDER BY Category,ID DESC";
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
                                if (SqlData["Category"].ToString().Trim() == "1")
                                {
                                    PersonResult.Add(new PersonIdle
                                    {
                                        code = SqlData["ID"].ToString(),
                                        DisplayName = SqlData["item"].ToString()
                                    });
                                }
                                else if (SqlData["Category"].ToString().Trim() == "2")
                                {
                                    MachineResult.Add(new MachineIdle
                                    {
                                        code = SqlData["ID"].ToString(),
                                        DisplayName = SqlData["item"].ToString()
                                    });
                                }
                            }
                        }
                    }
                }
            }

            result.Add(new IdleChartInfo
            {
                IdleType = IdleType,
                machinestatus = MachineStatus,
                MachineIdles = MachineResult.OrderBy(x => x.code).ToList(),
                PersonIdles = PersonResult.OrderBy(x => x.code).ToList()
            });

            return new ActionResponse<List<IdleChartInfo>>
            {
                Data = result
            };
        }


        /// <summary>
        /// 取得【單日待機時間區塊圖】待機資料 (O)*
        /// </summary>
        /// <param name="request">{"device": "2M301","date": "2022-02-16"}</param>
        /// <returns>
        /// </returns>
        [HttpPost("Machine/IdleReport")]
        public ActionResponse<List<idleData>> IdleReport([FromBody] DeviceIdeleAnalysisRequest request)
        {
            string st = ""; string et = "";
            if (request.Device == "" || request.Date == ""
                || request.Device == "string" || request.Date == "string"
                || request.Device == null || request.Date == null)
            {
                return new ActionResponse<List<idleData>>
                {
                    Data = null,
                };
            }
            else
            {
                DateTime dtDate;

                if (DateTime.TryParse(request.Date, out dtDate))
                {
                    st = dtDate.ToString("yyyy-MM-dd") + " 08:00:00";
                    et = dtDate.ToString("yyyy-MM-dd") + " 17:00:00";
                }
                else
                {
                    return new ActionResponse<List<idleData>>
                    {
                        Data = null,
                    };
                }
            }

            var result = new List<idleData>();
            string SqlStr = @$"SELECT TOP (100) a.*,b.item,b.Category,c.[user_name] 
                                FROM  [IdleReasonBinding] as a
                                LEFT JOIN  [IdleResult] as b on a.ReasonCode=b.ID
                                LEFT JOIN {_ConnectStr.AccuntDB}.[dbo].[User] as c on a.StaffID=c.[user_id]
                                WHERE　a.Device=@device and StartTime>@st and EndTime<@et";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@device", SqlDbType.VarChar).Value = request.Device;
                    comm.Parameters.Add("@st", SqlDbType.VarChar).Value = st;
                    comm.Parameters.Add("@et", SqlDbType.VarChar).Value = et;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                result.Add(new idleData
                                {
                                    idleTypeId = SqlData["Category"].ToString().Trim(),
                                    idleReasonId = SqlData["ReasonCode"].ToString().Trim(),
                                    idleStartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"].ToString()).ToString(_timeFormat),
                                    idleEndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"].ToString()).ToString(_timeFormat)
                                });
                            }
                        }
                    }
                }
            }

            result = result.OrderBy(x => x.idleStartTime).ToList();
            return new ActionResponse<List<idleData>>
            {
                Data = result
            };
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
        /// 取得【單日待機時間區塊圖】機台狀態資料 (O)*
        /// </summary>
        /// <param name="request">{"device": "2M301","date": "2022-02-16"}</param>
        /// <returns>
        /// </returns>
        [HttpPost("Machine/Gantt")]
        public ActionResponse<DailyDeviceStatus> Gantt([FromBody] DeviceIdeleAnalysisRequest request)//DateTime day)
        {
            string st = ""; string et = "";
            if (request.Device == "" || request.Date == ""
                || request.Device == "string" || request.Date == "string"
                || request.Device == null || request.Date == null)
            {
                return new ActionResponse<DailyDeviceStatus>
                {
                    Data = null,
                };
            }
            else
            {
                DateTime dtDate;

                if (DateTime.TryParse(request.Date, out dtDate))
                {
                    st = dtDate.ToString("yyyy-MM-dd") + " 00:00:00";
                    et = dtDate.AddDays(1).ToString("yyyy-MM-dd") + " 00:00:00";
                }
                else
                {
                    return new ActionResponse<DailyDeviceStatus>
                    {
                        Data = null,
                    };
                }
            }

            var result = new DailyDeviceStatus();
            var timeinterval = new DeviceTimeInterval();
            timeinterval.ST = "8:00";
            timeinterval.ET = "17:00";
            result.Interval = timeinterval;


            var temp = new List<devicestatus>();

            string SqlStr = @$"
                            SELECT *
                              FROM {_ConnectStr.SkyMarsDB}.[dbo].[UtilizationLog]
                              where MachineName=@device and StartDateTime>@st and EndDateTime<@et";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@device", SqlDbType.VarChar).Value = request.Device;
                    comm.Parameters.Add("@st", SqlDbType.VarChar).Value = st;
                    comm.Parameters.Add("@et", SqlDbType.VarChar).Value = et;

                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                temp.Add(new devicestatus
                                {
                                    Devicesattusid = SqlData["DeviceStatus"].ToString(),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartDateTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartDateTime"].ToString()).ToString(_timeFormat_ToMin),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndDateTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndDateTime"].ToString()).ToString(_timeFormat_ToMin)
                                });
                            }
                        }
                    }
                }
            }
            temp = temp.OrderBy(x => x.StartTime).ToList();
            result.devicestatus = temp;
            return new ActionResponse<DailyDeviceStatus>
            {
                Data = result
            };
        }

        private string checkMachineStatusCode(string data)
        {
            string result = "OFF";
            switch (data)
            {
                case "0":
                    result = "OFF";
                    break;
                case "1":
                    result = "RUN";
                    break;
                case "2":
                    result = "IDLE";
                    break;
                case "3":
                    result = "ALARM";
                    break;
            }
            return result;
        }

        string[] idlereason01 = { "待人員條機", "待人員開機", };
        string[] idlereason02 = { "","機器故障", "機台維修", "機台維修", "機台清理", "停機換線", "跳電", "待料", "零件上下料", "上料設置",
        "首件送驗","製程中檢驗","治具安裝","螺絲壓板鎖副","刀具鎖壞更換","測試用待機原因",};

        /// <summary>
        /// 取得【區間待機原因長條圖】圖表資訊 (O)*
        /// </summary>
        /// <returns>"type":1(人員)、2(機台)</returns>
        [HttpGet("Machine/IdleTimeSpandChartInfo")]
        public ActionResponse<List<TimeSpandChartInfo>> IdleTimeSpandChartInfo()
        {
            var result = new List<TimeSpandChartInfo>();
            var PersonResult = new List<PersonIdle>();
            var MachineResult = new List<MachineIdle>();

            string SqlStr = @"SELECT * FROM  [IdleResult]　where exist=1　ORDER BY Category,ID DESC";
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
                                if (SqlData["Category"].ToString().Trim() == "1")
                                {
                                    PersonResult.Add(new PersonIdle
                                    {
                                        code = SqlData["ID"].ToString(),
                                        DisplayName = SqlData["item"].ToString()
                                    });
                                }
                                else if (SqlData["Category"].ToString().Trim() == "2")
                                {
                                    MachineResult.Add(new MachineIdle
                                    {
                                        code = SqlData["ID"].ToString(),
                                        DisplayName = SqlData["item"].ToString()
                                    });
                                }
                            }
                        }
                    }
                }
            }

            result.Add(new TimeSpandChartInfo
            {
                MachineIdles = MachineResult.OrderBy(x => x.code).ToList(),
                PersonIdles = PersonResult.OrderBy(x => x.code).ToList()
            });

            return new ActionResponse<List<TimeSpandChartInfo>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 取得【區間待機原因長條圖】機台狀態資料 (O)*
        /// </summary>
        /// <param name="request">{"device": "2M301","stDate": "2022-02-14 00:00:00","etDate": "2022-02-25 00:00:00"}</param>
        /// <returns>
        /// </returns>
        [HttpPost("Machine/IdleTimeSpandChart")]
        public ActionResponse<List<IdleChart>> IdleTimeSpandChart([FromBody] DeviceIdeleRangeRequest request)//DateTime day)
        {
            Random rnd = new Random();
            if (request.Device == "" || request.stDate == "" || request.etDate == ""
                || request.Device == "string" || request.stDate == "string" || request.etDate == "string"
                || request.Device == null || request.stDate == null || request.etDate == null)
            {
                return new ActionResponse<List<IdleChart>>
                {
                    Data = null,
                };
            }
            else
            {
                DateTime dtDate;

                if (DateTime.TryParse(request.stDate, out dtDate))
                {
                    if (DateTime.TryParse(request.etDate, out dtDate))
                    {

                    }
                    else
                    {
                        return new ActionResponse<List<IdleChart>>
                        {
                            Data = null,
                        };
                    }
                }
                else
                {
                    return new ActionResponse<List<IdleChart>>
                    {
                        Data = null,
                    };
                }
            }

            var result = new List<IdleChart>();
            var temp = new List<idleData>();

            string SqlStr = @$"SELECT TOP (100) a.*,b.item,b.Category,c.[user_name] 
                                FROM  [IdleReasonBinding] as a
                                LEFT JOIN  [IdleResult] as b on a.ReasonCode=b.ID
                                LEFT JOIN {_ConnectStr.AccuntDB}.[dbo].[User] as c on a.StaffID=c.[user_id]
                                WHERE　a.Device=@device and StartTime>@st and EndTime<@et";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(SqlStr, conn))
                {
                    comm.Parameters.Add("@device", SqlDbType.VarChar).Value = request.Device;
                    comm.Parameters.Add("@st", SqlDbType.VarChar).Value = request.stDate;
                    comm.Parameters.Add("@et", SqlDbType.VarChar).Value = request.etDate;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                temp.Add(new idleData
                                {
                                    idleTypeId = SqlData["Category"].ToString().Trim(),
                                    idleReasonId = SqlData["ReasonCode"].ToString().Trim(),
                                    idleStartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"].ToString()).ToString(_timeFormat),
                                    idleEndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"].ToString()).ToString(_timeFormat)
                                });
                            }
                        }
                    }
                }
            }

            var codelist_M = new List<string> { "64", "65", "75", "80", "83", "84", "85", "87", "88", "89", "90",};
            var codelist_P = new List<string> { "57", "62", "79", "81", "82", "86", "91", "92", "93", "94", "95"};
            var temp01 = temp.Distinct(x => x.idleReasonId).ToList();

            foreach (var i in temp01)
            {
                if (i.idleTypeId == "1")
                {
                    result.Add(new IdleChart
                    {
                        idleReasonId = i.idleReasonId.ToString(),
                        idleTypeId = "1",
                        count = rnd.Next(60, 150),
                        SpandTime = rnd.Next(16, 170)
                    });
                }
                else if (i.idleTypeId == "2")
                {
                    result.Add(new IdleChart
                    {
                        idleReasonId = i.idleReasonId.ToString(),
                        idleTypeId = "2",
                        count = rnd.Next(60, 150),
                        SpandTime = rnd.Next(16, 170)
                    });
                }
            }

            foreach (var item in codelist_M)
            {
                if (result.Exists(x => x.idleReasonId == item))
                {
                    continue;
                }
                else
                {
                    result.Add(new IdleChart
                    {
                        idleReasonId = item,
                        idleTypeId = "2",
                        count = rnd.Next(60, 150),
                        SpandTime = rnd.Next(16, 170)
                    });
                }
            }

            foreach (var item in codelist_P)
            {
                if (result.Exists(x => x.idleReasonId == item))
                {
                    continue;
                }
                else
                {
                    result.Add(new IdleChart
                    {
                        idleReasonId = item,
                        idleTypeId = "1",
                        count = rnd.Next(60, 150),
                        SpandTime = rnd.Next(16, 170)
                    });
                }
            }

            result = result.OrderBy(x => x.idleReasonId).ThenBy(x => x.idleTypeId).ToList();

            return new ActionResponse<List<IdleChart>> { Data = result };
        }

        #endregion
    }


    public class DeviceIdeleAnalysisRequest
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 待機原因查詢-日期
        /// </summary>
        public string Date { get; set; }
    }

    public class DeviceIdeleRangeRequest
    {
        /// <summary>
        /// 機台編號
        /// </summary>
        public string Device { get; set; }
        /// <summary>
        /// 待機原因查詢-起始時間
        /// </summary>
        public string stDate { get; set; }
        /// <summary>
        /// 待機原因查詢-結束時間
        /// </summary>
        public string etDate { get; set; }
    }

    public class IdleChart
    {
        /// <summary>
        /// 待機編號
        /// </summary>
        public string idleReasonId { get; set; }
        /// <summary>
        /// 待機原因種類【1:人員:PERSONAL 、2:機台:MACHINE】
        /// </summary>
        public string idleTypeId { get; set; }
        /// <summary>
        /// 待機次數
        /// </summary>
        public int count { get; set; }
        /// <summary>
        /// 待機時間
        /// </summary>
        public int SpandTime { get; set; }
    }

    public class idleData
    {
        /// <summary>
        /// 待機種類【1:人員、2:機台】
        /// </summary>
        public string idleTypeId { get; set; }
        /// <summary>
        /// 待機原因編號
        /// </summary>
        public string idleReasonId { get; set; }
        /// <summary>
        /// 待機開始時間
        /// </summary>
        public string idleStartTime { get; set; }
        /// <summary>
        /// 待機結束時間
        /// </summary>
        public string idleEndTime { get; set; }
    }


    public class IdleChartInfo
    {
        /// <summary>
        /// 待機種類標籤
        /// </summary>
        public List<analysisIdleType> IdleType { get; set; }

        /// <summary>
        /// 機台狀態標籤
        /// </summary>
        public List<analysisMachineStatus> machinestatus { get; set; }

        /// <summary>
        /// 人員待機原因標籤
        /// </summary>
        public List<PersonIdle> PersonIdles { get; set; }

        /// <summary>
        /// 機台待機原因標籤
        /// </summary>
        public List<MachineIdle> MachineIdles { get; set; }

    }

    public class analysisIdleType
    {
        /// <summary>
        /// 待機種類代號
        /// </summary>
        public string code { get; set; }

        /// <summary>
        /// 待機種類名稱
        /// </summary>
        public string DisplayName { get; set; }
    }
    public class analysisMachineStatus
    {
        /// <summary>
        /// 機台狀態代號
        /// </summary>
        public string code { get; set; }

        /// <summary>
        /// 機台狀態
        /// </summary>
        public string DisplayName { get; set; }
    }
    public class PersonIdle
    {
        /// <summary>
        /// 人員待機代號
        /// </summary>
        public string code { get; set; }

        /// <summary>
        /// 人員待機原因
        /// </summary>
        public string DisplayName { get; set; }

    }
    public class MachineIdle
    {
        /// <summary>
        /// 機台待機原因代號
        /// </summary>
        public string code { get; set; }

        /// <summary>
        /// 機台待機原因
        /// </summary>
        public string DisplayName { get; set; }
    }


    public class TimeSpandChartInfo
    {
        /// <summary>
        /// 人員待機原因標籤
        /// </summary>
        public List<PersonIdle> PersonIdles { get; set; }
        /// <summary>
        /// 機台待機原因標籤
        /// </summary>
        public List<MachineIdle> MachineIdles { get; set; }

    }


}

