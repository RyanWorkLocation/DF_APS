using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using PMCDash.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace PMCDash.Controllers.Part2
{
    [Route("api/[controller]")]
    [ApiController]
    public class ValuesController : ControllerBase
    {
        private readonly string _connectStr = @"Data Source = 127.0.0.1; Initial Catalog = DPI;User ID = MES2014; Password = PMCMES;";

        [HttpPost("Edit")]
        public ActionResponse<List<OrderInfo>> Edit(EditOrderModels edit)
        {
            var result = new List<OrderInfo>();
            var Tep = getTep("Assignment");
            var Cdata = getCdata(Tep);
            if (string.IsNullOrEmpty(edit.StartTime))
            {
                var st = Tep.Where(x => x.WorkGroup == edit.WorkGroup).ToList();
                if (st.Count != 0)
                {
                    edit.StartTime = st.Max(x => Convert.ToDateTime(x.EndTime)).ToString("yyyy-MM-dd HH:mm");
                }
                else
                {
                    edit.StartTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm");
                }
            }
            ScheduleAlgorithm(edit, ref Tep, Cdata);

            return new ActionResponse<List<OrderInfo>> { Data = result };
        }

        /// <summary>
        /// ["2Y501"]
        /// </summary>
        /// <param name="machine">["2Y501"]</param>
        /// <returns></returns>
        [HttpPost("DispatchAlarm")]
        public ActionResponse<List<OrderInfo>> DispatchAlarm(List<string> machine)//調度警報
        {
            var result = new List<OrderInfo>();
            var Tep = getTep("Assignment");
            if (machine.Count == 1)
            {
                if (string.IsNullOrEmpty(machine[0]))
                {
                    //ViewBag.deviceRemarks = JsonConvert.SerializeObject(devices.Select(x => x.Remark));
                    //ViewBag.taskNames = JsonConvert.SerializeObject(devices.Select(x => x.Remark));
                    //ViewBag.tasks = JsonConvert.SerializeObject(CData);
                    //TempData["data"] = Tep;
                    //TempData["message"] = @"目前沒有故障機台，所以沒有異動";
                    throw new Exception("目前沒有故障機台，所以沒有異動");
                }
            }
            var devices = getDevices();
            var deviceAlarm = new List<Device>(devices);
            for (int i = 0; i < machine.Count; i++)
            {
                int idx = deviceAlarm.FindIndex(x => x.Remark == machine[i]);
                deviceAlarm.RemoveAt(idx);
                Tep.Where(x => x.WorkGroup == machine[i])
                   .ToList()
                   .ForEach(x => x.WorkGroup = string.Empty);
            }

            var works = Tep.FindAll(x => string.IsNullOrEmpty(x.WorkGroup) == true)
                           .OrderBy(x => Convert.ToDateTime(x.AssignDate_PM))
                           .Distinct(x => x.OrderID)
                           .ToList();
            var CData = getCdata(Tep);


            var firstChoice = deviceAlarm.Where(x => !CData.Exists(y => y.WorkGroup == x.Remark))
                                         .ToList();

            foreach (var work in works)
            {
                string deviceName;
                if (firstChoice.Count != 0)
                {
                    Random rand = new Random(Guid.NewGuid().GetHashCode());
                    deviceName = firstChoice[rand.Next(0, firstChoice.Count)].Remark;
                }
                else
                {
                    //找到甘特圖中工作量最少的機台
                    deviceName = CData.Where(x => !firstChoice.Exists(y => y.Remark.Trim() == x.WorkGroup.Trim()))//判斷CData和firstchioce沒有依樣的機台
                                      .Where(x => !machine.Exists(y => y == x.WorkGroup.Trim()))//判斷"故障機台"和firstchioce沒有依樣的機台
                                      .GroupBy(x => x.WorkGroup)
                                      .Select(x => new { count = x.Count(), deviceName = x.Key })
                                      .OrderBy(x => x.count)
                                      .ToList()[0].deviceName;
                }

                var temp = Tep.FindAll(x => x.OrderID == work.OrderID)
                              .OrderBy(x => x.OPID)
                              .ToList();

                foreach (var item in temp)
                {
                    EditOrderModels edit = new EditOrderModels
                    {
                        OrderID = item.OrderID,
                        OPID = item.OPID,
                        StartTime = DateTime.Now.ToString("yyyy-MM-dd HH:mm"),
                        WorkGroup = deviceName
                    };
                    ScheduleAlgorithm(edit, ref Tep, CData);
                }

                foreach (var item in Tep)
                {
                    if (!CData.Exists(x => x.OrderID == item.OrderID && x.OPID == item.OPID))
                    {
                        if (!(item.OrderID == work.OrderID))
                            continue;
                        CData.Add(new ChartData
                        {
                            OrderID = item.OrderID,
                            OPID = item.OPID,
                            AssignST = item.StartTime,
                            AssignET = item.EndTime,
                            WorkGroup = item.WorkGroup,
                            OPSTATE = "UnSchedule"
                        });
                    }
                }

                var TData = new List<ChartData>();
                var TOrder = CData.Distinct(x => x.OrderID)
                                  .Select(x => x.OrderID);
                for (int j = 0; j < 2; j++)
                {
                    foreach (var item in TOrder)
                    {
                        TData = CData.FindAll(x => x.OrderID == item)
                                     .OrderBy(x => x.OPID)
                                     .ToList();
                        TData = CheckOP(TData);
                        for (int i = 0; i < TData.Count; i++)
                        {
                            int index = CData.FindIndex(x => x.OrderID == TData[i].OrderID && x.OPID == TData[i].OPID);
                            CData[index].AssignST = TData[i].AssignST;
                            CData[index].AssignET = TData[i].AssignET;
                        }
                    }
                    TOrder = CData.Distinct(x => x.WorkGroup)
                                  .Select(x => x.WorkGroup)
                                  .ToList();
                    foreach (var item in TOrder)
                    {
                        TData = CData.FindAll(x => x.WorkGroup == item)
                                     .OrderBy(x => x.AssignST)
                                     .ToList();
                        TData = CheckOP(TData);
                        for (int i = 0; i < TData.Count; i++)
                        {
                            int index = CData.FindIndex(x => x.OrderID == TData[i].OrderID && x.OPID == TData[i].OPID);
                            CData[index].AssignST = TData[i].AssignST;
                            CData[index].AssignET = TData[i].AssignET;
                        }
                    }
                    for (int i = 0; i < CData.Count; i++)
                    {
                        var query = Tep.Where(x => x.OrderID == CData[i].OrderID && x.OPID == CData[i].OPID)
                                       .Select(x => x)
                                       .First();
                        query.StartTime = CData[i].AssignST;
                        query.EndTime = CData[i].AssignET;
                    }


                }
            }

            result = Tep;

            return new ActionResponse<List<OrderInfo>> { Data = result };
        }


        private void ScheduleAlgorithm(EditOrderModels edit, ref List<OrderInfo> Tep, List<ChartData> CData)
        {
            string SqlStr = @"SELECT OrderQTY, HumanOpTime, MachOpTime, AssignDate, AssignDate_PM, MAKTX
                              FROM Assignment
                              WHERE OrderID = @OrderID and OPID = @OPID";
            List<ChartData> MachineSeq = new List<ChartData>();
            List<ChartData> TempSeq = new List<ChartData>();
            OrderInfo dep = new OrderInfo
            {
                OrderID = edit.OrderID,
                OPID = edit.OPID,
                StartTime = edit.StartTime,
                WorkGroup = edit.WorkGroup
            };
            string DateFormat = "yyyy-MM-dd HH:mm";
            int SqlCount = 0, Idx = 0;
            double TotalTime = 0;
            using (SqlConnection Conn = new SqlConnection(_connectStr))
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
                                TotalTime = Convert.ToDouble(SqlData["HumanOpTime"]) + Convert.ToDouble(SqlData["MachOpTime"]);
                                dep.OrderQTY = SqlCount;
                                dep.AssignDate = Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd");
                                dep.AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd");
                                dep.Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim();
                            }
                        }
                    }
                }
            }
            TimeSpan PaddingTime = new TimeSpan(0);
            TimeSpan PostTime = new TimeSpan(0, 0, (int)(SqlCount * TotalTime), 0, 0);
            DateTime PostST = Convert.ToDateTime(dep.StartTime);
            DateTime PostET = PostST + PostTime;
            Idx = Tep.FindIndex(x => x.OrderID == dep.OrderID && x.OPID == dep.OPID);
            if (Idx != -1)
            {
                Tep.RemoveAt(Idx);
                Idx = CData.FindIndex(x => x.OrderID == dep.OrderID && x.OPID == dep.OPID);
                if (Idx != -1)
                {
                    CData.RemoveAt(Idx);
                }
            }
            MachineSeq = CData.FindAll(x => x.WorkGroup == dep.WorkGroup).OrderBy(x => x.AssignST).ToList();
            Idx = CData.FindLastIndex(x => x.OrderID == dep.OrderID && x.OPID < dep.OPID && DateTime.Compare(Convert.ToDateTime(x.AssignET), PostST) > 0);
            if (Idx != -1)
            {
                PostST = Convert.ToDateTime(CData[Idx].AssignET);
                PostET = PostST + PostTime;
            }
            for (int i = 0; i < MachineSeq.Count; i++)//找出區段 重疊或者不重疊 重疊部分若相同插入優先原則
            {
                if (DateTime.Compare(Convert.ToDateTime(MachineSeq[i].AssignET), PostST) > 0)
                {
                    if (DateTime.Compare(Convert.ToDateTime(MachineSeq[i].AssignST), PostST) > 0)
                    {
                        break;
                    }
                    else
                    {
                        if (DateTime.Compare(Convert.ToDateTime(MachineSeq[i].AssignST), PostST) == 0)
                        {
                            break;
                        }
                        else
                        {
                            PostST = Convert.ToDateTime(MachineSeq[i].AssignET);
                            PostET = PostST + PostTime;
                            break;
                        }
                    }
                }
            }
            // ET > ST
            Idx = MachineSeq.FindIndex(x => DateTime.Compare(PostST, Convert.ToDateTime(x.AssignET)) < 0);
            if (Idx != -1)
            {
                for (int i = Idx; i < MachineSeq.Count; i++)
                {
                    PaddingTime = Convert.ToDateTime(MachineSeq[i].AssignET) - Convert.ToDateTime(MachineSeq[i].AssignST);
                    if (i == Idx && DateTime.Compare(PostST, Convert.ToDateTime(MachineSeq[i].AssignST)) == 0)
                    {
                        MachineSeq[i].AssignST = PostET.ToString(DateFormat);
                        MachineSeq[i].AssignET = (Convert.ToDateTime(MachineSeq[i].AssignST) + PaddingTime).ToString(DateFormat);
                    }
                    else if (i == Idx && DateTime.Compare(PostST, Convert.ToDateTime(MachineSeq[i].AssignST)) > 0)
                    {
                        PostST = Convert.ToDateTime(MachineSeq[i].AssignET);
                        PostET = PostST + PostTime;
                    }
                    else if (i == Idx && DateTime.Compare(PostST, Convert.ToDateTime(MachineSeq[i].AssignST)) < 0)
                    {
                        //MachineSeq[i].startDate = PostET.ToString(DateFormat);
                        //MachineSeq[i].endDate = (PostET + PaddingTime).ToString(DateFormat);
                    }
                    else
                    {
                        MachineSeq[i].AssignST = MachineSeq[i - 1].AssignET;
                        MachineSeq[i].AssignET = (Convert.ToDateTime(MachineSeq[i].AssignST) + PaddingTime).ToString(DateFormat);
                    }

                    TempSeq.Add(MachineSeq[i]);
                    if ((i + 1) == MachineSeq.Count)
                    {
                        break;
                    }
                    else
                    {
                        if (DateTime.Compare(Convert.ToDateTime(MachineSeq[i].AssignET), Convert.ToDateTime(MachineSeq[i + 1].AssignST)) < 0)
                        {
                            break;
                        }
                    }
                }
            }

            ChartData Temp = new ChartData
            {
                AssignST = PostST.ToString(DateFormat),
                AssignET = PostET.ToString(DateFormat),
                WorkGroup = dep.WorkGroup,
                OPSTATE = "UnSchedule",
                OPID = dep.OPID,
                OrderID = dep.OrderID,
            };
            dep.StartTime = PostST.ToString(DateFormat);
            dep.EndTime = PostET.ToString(DateFormat);
            Tep.Add(dep);
            for (int i = 0; i < TempSeq.Count; i++)
            {
                Idx = CData.FindIndex(x => x.OrderID == TempSeq[i].OrderID && x.OPID > TempSeq[i].OPID);
                if (Idx != -1)
                {
                    if (DateTime.Compare(Convert.ToDateTime(TempSeq[i].AssignET), Convert.ToDateTime(CData[Idx].AssignST)) > 0)
                    {
                        PaddingTime = Convert.ToDateTime(CData[Idx].AssignET) - Convert.ToDateTime(CData[Idx].AssignST);
                        CData[Idx].AssignST = TempSeq[i].AssignET;
                        CData[Idx].AssignET = (Convert.ToDateTime(CData[Idx].AssignST) + PaddingTime).ToString(DateFormat);
                    }
                }
            }

            for (int i = 0; i < MachineSeq.Count; i++)
            {
                Idx = CData.FindIndex(x => x.OPID == MachineSeq[i].OPID && x.OrderID == MachineSeq[i].OrderID);
                CData[Idx].AssignST = MachineSeq[i].AssignST;
                CData[Idx].AssignET = MachineSeq[i].AssignET;
            }
            CData.Add(Temp);

            List<ChartData> TData = new List<ChartData>();
            var TOrder = CData.Distinct(x => x.OrderID).Select(x => x.OrderID);
            for (int j = 0; j < 2; j++)
            {
                foreach (var item in TOrder)
                {
                    TData = CData.FindAll(x => x.OrderID == item).OrderBy(x => x.OPID).ToList();
                    TData = CheckOP(TData);
                    for (int i = 0; i < TData.Count; i++)
                    {
                        int index = CData.FindIndex(x => x.OrderID == TData[i].OrderID && x.OPID == TData[i].OPID);
                        CData[index].AssignST = TData[i].AssignST;
                        CData[index].AssignET = TData[i].AssignET;
                    }
                }
                TOrder = CData.Distinct(x => x.WorkGroup).Select(x => x.WorkGroup).ToList();
                foreach (var item in TOrder)
                {
                    TData = CData.FindAll(x => x.WorkGroup == item).OrderBy(x => x.AssignST).ToList();
                    TData = CheckOP(TData);
                    for (int i = 0; i < TData.Count; i++)
                    {
                        int index = CData.FindIndex(x => x.OrderID == TData[i].OrderID && x.OPID == TData[i].OPID);
                        CData[index].AssignST = TData[i].AssignST;
                        CData[index].AssignET = TData[i].AssignET;
                    }
                }
            }

            // 更新表格內容資料 TEP 與 CDATA 相對應
            for (int i = 0; i < CData.Count; i++)
            {
                var query = Tep.Where(x => x.OrderID == CData[i].OrderID && x.OPID == CData[i].OPID)
                        .Select(x => x).First();
                query.StartTime = CData[i].AssignST;
                query.EndTime = CData[i].AssignET;
            }
            CountDelay(Tep);
        }

        private List<OrderInfo> getTep(string TableName)
        {
            var result = new List<OrderInfo>();

            List<string> CustomerLocation_List = new List<string> { "台北", "桃園", "新竹", "苗栗", "台中", "台南", "高雄" };
            List<string> CustomerName_List = new List<string> { "太平洋", "力齊", "亞迪斯", "永利安", "伯佳", "登榮", "亨利興", "磐石", "鑫隆" };

            string SqlStr = $@"SELECT  a.*, w.WIPEvent, CPK,
                            Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0)
                            FROM {TableName} as a
                            LEFT JOIN WIP as w ON (a.OrderID = w.OrderID AND a.OPID = w.OPID)";
            using (var Conn = new SqlConnection(_connectStr))
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
                            Random rnd = new Random();
                            int i = 0;
                            while (SqlData.Read())
                            {
                                var CustomerLocation_choose = CustomerLocation_List[rnd.Next(CustomerLocation_List.Count)];
                                var CustomerName_choose = CustomerName_List[rnd.Next(CustomerName_List.Count)];

                                result.Add(new OrderInfo
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"]),
                                    StartTime = string.IsNullOrEmpty(SqlData["StartTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["StartTime"]).ToString("yyyy-MM-dd HH:mm"),
                                    OrderQTY = Convert.ToInt32(SqlData["OrderQTY"]),
                                    EndTime = string.IsNullOrEmpty(SqlData["EndTime"].ToString()) ? "" : Convert.ToDateTime(SqlData["EndTime"]).ToString("yyyy-MM-dd HH:mm"),
                                    WorkGroup = string.IsNullOrEmpty(SqlData["WorkGroup"].ToString()) ? "" : SqlData["WorkGroup"].ToString(),
                                    AssignDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate"]).ToString("yyyy-MM-dd"),
                                    AssignDate_PM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString()) ? "" : Convert.ToDateTime(SqlData["AssignDate_PM"]).ToString("yyyy-MM-dd"),
                                    Maktx = SqlData["MAKTX"].ToString().Split('(')[0].Trim(),
                                    HumanOpTime = Convert.ToDouble(SqlData["HumanOpTime"]),
                                    MachOpTime = Convert.ToDouble(SqlData["MachOpTime"]),
                                    Progress = ChangeProgressIntFormat(SqlData["Progress"].ToString()),
                                    Note = string.IsNullOrEmpty(SqlData["Note"].ToString()) ? "-" : SqlData["Note"].ToString(),
                                    CustomerName = CustomerName_choose,
                                    CustomerLocation = CustomerLocation_choose,
                                    Important = Convert.ToBoolean(SqlData["Important"].ToString()),
                                    Assign = Convert.ToBoolean(SqlData["Scheduled"].ToString()),
                                    PRIORITY = Convert.ToInt32(SqlData["PRIORITY"]),
                                    OPLTXA1 = SqlData["OPLTXA1"].ToString().Trim()
                                });
                            }
                        }
                    }
                }
            }

            return result;
        }
        private List<ChartData> getCdata(List<OrderInfo> Tep)
        {
            var result = new List<ChartData>();
            foreach (var item in Tep)
            {
                var temp = result.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                if (temp is null)
                {
                    result.Add(new ChartData
                    {
                        AssignST = item.StartTime,
                        AssignET = item.EndTime,
                        OPID = item.OPID,
                        OrderID = item.OrderID,
                        OPSTATE = "UNSTARTED",
                        WorkGroup = item.WorkGroup
                    });
                }
                else
                {
                    result.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).AssignST = item.StartTime;
                    result.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).AssignET = item.EndTime;
                    result.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).WorkGroup = item.WorkGroup;
                    result.Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID).OPSTATE = "UNSTARTED";
                }
            }
            return result;
        }
        private List<Device> getDevices()
        {
            var result = new List<Device>();
            string SqlStr = @"SELECT * FROM  [Device]";
            using (var Conn = new SqlConnection(_connectStr))
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
                                result.Add(new Device
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
            return result;
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
        private string CheckOPState(string item)
        {
            string resulr = "N/A";
            switch (item.Trim())
            {
                case "0":
                    resulr = "UNSTARTED";
                    break;
                case "1":
                    resulr = "PROCESSING";
                    break;
                case "2":
                    resulr = "UNSTARTED";
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
        private List<ChartData> CheckOP(List<ChartData> Data)
        {
            string DateFormat = "yyyy-MM-dd HH:mm";
            for (int i = 1; i < Data.Count; i++)
            {
                if (DateTime.Compare(Convert.ToDateTime(Data[i - 1].AssignET), Convert.ToDateTime(Data[i].AssignST)) > 0)
                {
                    TimeSpan TS = Convert.ToDateTime(Data[i].AssignET) - Convert.ToDateTime(Data[i].AssignST);
                    Data[i].AssignST = Data[i - 1].AssignET;
                    Data[i].AssignET = (Convert.ToDateTime(Data[i].AssignST) + TS).ToString(DateFormat);
                }
            }
            return Data;
        }
        private void CountDelay(List<OrderInfo> Tep)
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
        /// 設備故障排程法 (O)*
        /// </summary>
        /// <param name="resquest">{  "machine": "2M301",  "mode": "A",  "startTime": "2021-01-02 10:00:00",  "endTime": "2021-01-02 13:00:00"}</param>
        /// <returns></returns>
        [HttpPost("BreakdownGanttChart")]
        public ActionResponse<List<Schedulelist_MB>> BreakdownGanttChart([FromBody] MachineBreakdownRequest resquest)
        {
            if (resquest.Machine == "" || resquest.StartTime == "" || resquest.EndTime == "" || resquest.Mode == ""
                || resquest.Machine == "string" || resquest.StartTime == "string" || resquest.EndTime == "string" || resquest.Mode == "string"
                || resquest.Machine == null || resquest.StartTime == null || resquest.EndTime == null || resquest.Mode == null)
            {
                throw new Exception("目前沒有故障機台，所以沒有異動");
            }


            var result = new List<Schedulelist_MB>();
            var _temp = new List<OrderInfo>();
            var breakdowninfo = new MachinebreakdowInfo();
            breakdowninfo.Machine = resquest.Machine;
            breakdowninfo.StartTime = resquest.StartTime;
            breakdowninfo.EndTime = resquest.EndTime;

            //輩分assigment至assigmentTemp05
            var sqlStr = $@"
                            DELETE AssignmentTemp6

                            INSERT INTO AssignmentTemp6 ([OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
                                  ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
                                  ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
                                  ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important])
                            SELECT [OrderID],[OPID],[OPLTXA1],[MachOpTime],[HumanOpTime]
                                  ,[StartTime],[EndTime],[WorkGroup],[Operator],[AssignDate]
                                  ,[Parent],[SAP_WorkGroup],[OrderQTY],[Scheduled],[AssignDate_PM],[ShipAdvice]
                                  ,[IsSkip],[MAKTX],[PRIORITY],[Note],[Important]
                            FROM Assignment
                            ";
            using (var Conn = new SqlConnection(_connectStr))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();
                using (var Comm = new SqlCommand(sqlStr, Conn))
                {
                    int impactrow = Comm.ExecuteNonQuery();
                    if (impactrow > 0)
                    {
                        //result = "Update Successful";
                    }
                    else
                    {
                        //result = "Can not update";
                    }

                }
            }




            //result.Add(new Schedulelist_MB
            //{
            //    mode = "6",
            //    breakdownInfo = breakdowninfo,
            //    schedules = temp.OrderBy(x => x.WorkGroup).ThenBy(x => x.StartTime).ToList()
            //});

            return new ActionResponse<List<Schedulelist_MB>>
            {
                Data = result
            };
        }


        // GET: api/<ValuesController>
        [HttpGet]
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/<ValuesController>/5
        [HttpGet("{id}")]
        public string Get(int id)
        {
            return "value";
        }

        //Hosted web API REST Service base url  
        string Baseurl = "https://lab2.ezerp.tw/";
        string SetAccessToken = "Bearer 767t3Bj1pbR4xR7xcfxxg-nNCFLsc78D_u5XBQWMWxrt0YrVM7QF4vQ5i2hR1yD0hJXZxjDn0EJe8Yl4fxNHBJMy7kry4soDjcH4NgSvvj5YtzQBOc-x9oZlG11esrGLe9r61JJQBfY3ZcE9FUrdkGkn7AlUqDPgwXKDj7uZIkeX-tGr7RcsKptxegYzXrjBblmhk8QSDGsula3Wvku0Bea_IVSl32P7v5jpoKnsTSTRMLukzviHehL09mli1h6aogBsO04yrQnM-ScZvVc2qXiJWUldBA55G2XnPfDL0lQfXGM9x9MA8SNlltQpmesXVBvAbpN_C4pM40hkqxvMR5OKJr0CWvn_15AegU7FTymrIsC9TJGEHCc0q45rl-QpE0Ui_4h5MObUB-SIbcjOcTahsftQe0mA-p8gDD3mbpZQSqOzgR2dyXPxydljUuvOXf_t802bHBIE_fM56uJeSTiffI9rwDOszLNb1nx5tkKH8GlcqGgp-9qRQuj8BnzQ3mwxg8LWQKI_EYwBQIN5uQ8ZbXb2FD4nHUGidXqaHIFwWtF7";

        [HttpGet("GetProduct")]
        public ActionResponse<List<JobOrder>> GetProductAsync()
        {
           var temp =  RequestWebAPI("{}", "api/JobOrders/Query");
            List<JobOrder> result = JsonConvert.DeserializeObject<List<JobOrder>>(temp);
            return new ActionResponse<List<JobOrder>> { Data = result };
        }

        private string RequestWebAPI(string sendData, string url)
        {
            try
            {
                string backMsg = "";
                WebClient client = new WebClient();
                client.Headers.Add("Authorization", SetAccessToken);
                client.Headers.Add("Content-Type", "application/json;charset=utf-8");
                backMsg = client.UploadString(Baseurl+url, "POST", sendData);
                return backMsg;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        // POST api/<ValuesController>
        [HttpPost]
        public void Post([FromBody] string value)
        {
        }

        // PUT api/<ValuesController>/5
        [HttpPut("{id}")]
        public void Put(int id, [FromBody] string value)
        {
        }

        // DELETE api/<ValuesController>/5
        [HttpDelete("{id}")]
        public void Delete(int id)
        {
        }
    }

    public class Employee
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string City { get; set; }

    }

    public class EditOrderModels
    {
        public string OrderID { get; set; }
        public double OPID { get; set; }
        public string StartTime { get; set; }
        public string WorkGroup { get; set; }
    }

    public class ChartData
    {
        public string AssignST { get; set; }
        public string AssignET { get; set; }
        public string WorkGroup { get; set; }
        public string OPSTATE { get; set; }
        public double OPID { get; set; }
        public string OrderID { get; set; }
    }


    public class Product
    {
        public int Id { get; set; }
        public string Number { get; set; }
        public string Name { get; set; }
    }

    public class Process
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Number { get; set; }
    }

    public class DrawMaterial
    {
        public object Id { get; set; }
        public object Number { get; set; }
    }

    public class JobOrder
    {
        public int Id { get; set; }
        public string Number { get; set; }
        public int Type { get; set; }
        public int Status { get; set; }
        public double Quantity { get; set; }
        public double RequiredQuantity { get; set; }
        public DateTime LeadTime { get; set; }
        public DateTime EstimatedDate { get; set; }
        public Product Product { get; set; }
        public Process Process { get; set; }
        public DrawMaterial DrawMaterial { get; set; }
        public DateTime CreateDateTime { get; set; }
        public DateTime UpdateDateTime { get; set; }
        public string CreateEmployeeStr { get; set; }
        public string UpdateEmployeeStr { get; set; }
        public string OrderId { get; set; }
        public string OrderNumber { get; set; }
        public double OrderRequiredQuantity { get; set; }
        public DateTime OrderDeliveryDate { get; set; }
    }
}
