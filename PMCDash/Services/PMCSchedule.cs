using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Collections.Concurrent;
using PMCDash.Models;
using PMCDash.Models.Part2;

namespace PMCDash.Services
{
    internal class DelayMthod : IMAthModel
    {
        ConnectStr _ConnectStr = new ConnectStr();
        private readonly string _timeFormat = "yyyy-MM-dd HH:mm:ss";
        public int Chromvalue { get; set; }
        private List<Device> Devices { get; set; }

        private DateTime PresetStartTime = DateTime.Now;
        public Dictionary<string, DateTime> ReportedMachine { get; set; }
        public Dictionary<string, DateTime> ReportedOrder { get; set; }

        public DelayMthod(int chromValue, List<Device> devices)
        {
            Chromvalue = chromValue;
            Devices = new List<Device>(devices);
        }

        public List<GaSchedule> CreateDataSet(string st, string et)
        {
            TimeSpan actDuration = new TimeSpan();
            string sqlStr = @$"SELECT a.SeriesID, a.OrderID, a.OPID,p.CanSync, a.Range, a.OrderQTY, a.HumanOpTime, a.MachOpTime, a.AssignDate, a.AssignDate_PM,a.MAKTX,wip.WIPEvent,a.Scheduled
                                 FROM {_ConnectStr.APSDB}.dbo.Assignment a left join {_ConnectStr.APSDB}.dbo.WIP as wip on a.SeriesID=wip.SeriesID
                                 Left Join {_ConnectStr.APSDB}.dbo.WipRegisterLog w on w.WorkOrderID = a.OrderID and w.OPID=a.OPID
                                 inner join {_ConnectStr.MRPDB}.dbo.Process as p on a.OPID=p.ID
                                 where w.WorkOrderID is NULL and (wip.WIPEvent=0 or wip.WIPEvent is NULL) and (a.AssignDate between '{st}' and '{et}')
                                 order by a.OrderID, a.Range";
            var result = new List<GaSchedule>();
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand Comm = new SqlCommand(sqlStr, Conn))
                {
                    if (Conn.State != ConnectionState.Open)
                        Conn.Open();
                    using (SqlDataReader SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                if(Convert.ToInt16(SqlData["CanSync"])==0)
                                {
                                    actDuration = new TimeSpan(0,
                                    (string.IsNullOrEmpty(SqlData["HumanOpTime"].ToString()) ? 1 : Convert.ToInt32(SqlData["HumanOpTime"])
                                    + (string.IsNullOrEmpty(SqlData["MachOpTime"].ToString()) ? 1 : Convert.ToInt32(SqlData["MachOpTime"]) * int.Parse(SqlData["OrderQTY"].ToString()))
                                    )
                                    , 0);
                                }
                                else
                                {
                                    actDuration = new TimeSpan(0,
                                    (string.IsNullOrEmpty(SqlData["HumanOpTime"].ToString()) ? 1 : Convert.ToInt32(SqlData["HumanOpTime"])
                                    + (string.IsNullOrEmpty(SqlData["MachOpTime"].ToString()) ? 1 : Convert.ToInt32(SqlData["MachOpTime"]))
                                    )
                                    , 0);
                                }
                                

                                result.Add(new GaSchedule
                                {
                                    SeriesID = SqlData["SeriesID"].ToString(),
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = Convert.ToDouble(SqlData["OPID"].ToString()),
                                    Range = int.Parse(SqlData["Range"].ToString().Trim()),
                                    Duration = actDuration,
                                    PartCount = int.Parse(SqlData["OrderQTY"].ToString()),
                                    Maktx = string.IsNullOrEmpty(SqlData["MAKTX"].ToString()) ? "N/A" : SqlData["MAKTX"].ToString(),
                                    Assigndate = Convert.ToDateTime(SqlData["AssignDate_PM"]),
                                    Scheduled = int.Parse(SqlData["Scheduled"].ToString()) //05.24 Ryan新增
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        public List<LocalMachineSeq> CreateSequence(List<GaSchedule> data)
        {
            var result = new List<LocalMachineSeq>();
            var i = new Random(Guid.NewGuid().GetHashCode());
            int j = 0;
            int machineSeq = 0;
            var rnd = new Random(Guid.NewGuid().GetHashCode());
            var eachorderlist = data.GroupBy(x => x.OrderID);
            //打亂製程排序，避免依照工單順序排
            var randomizedOrderid = data.Select(x=>x.OrderID).OrderBy(item => rnd.Next()).ToList();
            var rangeDictionary = data
                .GroupBy(x => x.OrderID)
                .ToDictionary(
                    key => key.Key,
                    value => value.Select(x => x.Range).OrderBy(x => x).ToList()
                );
            // 記錄每個 OrderID 的出現次數，並依序分配 Range 值
            var rangeTracker = new Dictionary<string, int>();
            var randomresult = new List<GaSchedule>();
            foreach (var item in randomizedOrderid)
            {
                if (!rangeTracker.ContainsKey(item))
                {
                    //var minrange = data.Where(x => x.OrderID == item).Select(y=> y.Range).Min();
                    rangeTracker[item] = 0; // 初始為欲排程
                }
                randomresult.Add(data.Find(x => x.OrderID == item && x.Range == rangeDictionary[item][rangeTracker[item]]));

                rangeTracker[item]++; // 增加該 OrderID 的 Range
            }



            //取得各機台資訊
            var OutsourcingList = getOutsourcings();

            //取得所有製程的替代機台
            var ProcessDetial = getProcessDetial();

            foreach (var item in randomresult)
            {
                //取得該工單可以分發的機台列表，若MRP Table內沒有相關資料可能會找不到可用機台，應該要回傳錯誤訊息，此次排成失敗
                var CanUseDevices = ProcessDetial.Where(x => x.ProcessID == item.OPID.ToString()).ToList();
                if (CanUseDevices.Count != 0)
                {
                    //隨機分派一台機台
                    j = rnd.Next(0, CanUseDevices.Count);
                    //機台名稱
                    if(OutsourcingList.Exists(x=>x.remark== CanUseDevices[j].remark))
                    {
                        if (OutsourcingList.Where(x=>x.remark== CanUseDevices[j].remark).First().isOutsource=="0") //非委外機台
                        {
                            if (result.Exists(x => x.WorkGroup == CanUseDevices[j].remark))
                            {
                                machineSeq = result.Where(x => x.WorkGroup == CanUseDevices[j].remark)
                                                   .Select(x => x.EachMachineSeq)
                                                   .Max() + 1;
                            }

                            else
                            {
                                machineSeq = 0;
                            }
                        }
                        else
                        {
                            machineSeq = 0;
                        }


                        result.Add(new LocalMachineSeq
                        {
                            SeriesID = item.SeriesID,
                            OrderID = item.OrderID,
                            OPID = item.OPID,
                            Range = item.Range,
                            Duration = item.Duration,
                            PredictTime = item.Assigndate,
                            Maktx = item.Maktx,
                            PartCount = item.PartCount,
                            WorkGroup = CanUseDevices[j].remark,
                            EachMachineSeq = machineSeq,
                        });
                    }
                }
            }
            return result;
        }

        private List<MRP.ProcessDetial> getProcessDetial()
        {
            List<MRP.ProcessDetial> result = new List<MRP.ProcessDetial>();
            string SqlStr = "";
            SqlStr = $@"
                        SELECT a.*,b.remark 
                          FROM {_ConnectStr.MRPDB}.[dbo].[ProcessDetial] as a
                          left join {_ConnectStr.APSDB}.dbo.Device as b on a.MachineID=b.ID
                          order by a.ProcessID,b.ID
                        ";
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
                                result.Add(new MRP.ProcessDetial
                                {
                                    ID = int.Parse(SqlData["ID"].ToString()),
                                    ProcessID = string.IsNullOrEmpty(SqlData["ProcessID"].ToString()) ? "" : SqlData["ProcessID"].ToString(),
                                    MachineID = string.IsNullOrEmpty(SqlData["MachineID"].ToString()) ? "" : SqlData["MachineID"].ToString(),
                                    remark = string.IsNullOrEmpty(SqlData["remark"].ToString()) ? "" : SqlData["remark"].ToString(),
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        public List<Device> getCanUseDevice(string OrderID, string OPID, string MAKTX)
        {
            List<Device> devices = new List<Device>();
            string SqlStr = @$"SELECT dd.* FROM Assignment as aa
                                inner join (SELECT a.Number,a.Name,a.RoutingID,b.ProcessRang,c.ID,c.ProcessNo,c.ProcessName FROM {_ConnectStr.MRPDB}.dbo.Part as a
                                inner join {_ConnectStr.MRPDB}.dbo.RoutingDetail as b on a.RoutingID=b.RoutingId
                                inner join {_ConnectStr.MRPDB}.dbo.Process as c on b.ProcessId=c.ID
                                where a.Number= (select top(1) MAKTX from Assignment where OrderID=@OrderID and OPID=@OPID) ) as bb on aa.MAKTX=bb.Number and aa.OPID=bb.ID
                                left join {_ConnectStr.MRPDB}.dbo.ProcessDetial as cc on bb.ID=cc.ProcessID
                                inner join Device as dd on cc.MachineID=dd.ID
                                where aa.OrderID=@OrderID and aa.OPID=@OPID";

            SqlStr = $@"select c.*
                        from Assignment as a
                        left join  {_ConnectStr.MRPDB}.dbo.ProcessDetial as b  on a.OPID=b.ProcessID
                        left join Device as c on b.MachineID=c.ID
                        where a.OrderID=@OrderID and a.OPID=@OPID";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();

                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    Comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    Comm.Parameters.Add(("@OPID"), SqlDbType.Float).Value = OPID;
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

            if (devices.Count == 0) ;

            return devices;
        }

        public List<Chromsome> Scheduled(List<LocalMachineSeq> firstSchedule)
        {
            var OutsourcingList = getOutsourcings();

            var result = new List<Chromsome>();
            int Idx = 0;
            DateTime preserve_Now = DateTime.Now;
            DateTime PostST = new DateTime();
            DateTime PostET = new DateTime();
            var SortSchedule = firstSchedule.OrderBy(x => x.EachMachineSeq).ToList();//依據seq順序排每一台機台

            // 建立快取，讓查找更快，並直接儲存對應的物件
            var orderIndexCache = new Dictionary<string, Chromsome>(); 

            var machineIndexCache = new Dictionary<string, Chromsome>();


            for (int i = 0; i < SortSchedule.Count; i++)
            {

                #region 原本的查找方式
                //目前排程結果已有同機台製程
                if (result.Exists(x => x.WorkGroup == SortSchedule[i].WorkGroup) && OutsourcingList.Exists(x=>x.remark== SortSchedule[i].WorkGroup))
                {
                    //是同步加工機台=>只需考量前製程完成時間
                    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "1")//該機台為委外機台
                    {
                        //目前排程是否有同工單
                        if(result.Exists(x=>x.OrderID== SortSchedule[i].OrderID))
                        {
                            int idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                            if (idx >= 0)
                                PostST = result[idx].EndTime;
                            else
                                PostST = preserve_Now;
                            //Idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                            //PostST = result[Idx].EndTime;
                        }
                        //如果目前排程沒有，則確認原排程是否有同工單製程
                        else if(ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            
                            //有的話，比較當前時間
                            PostST = preserve_Now > ReportedOrder[SortSchedule[i].OrderID] ? DateTime.Now : ReportedOrder[SortSchedule[i].OrderID];
                        }
                        else
                        {
                            //沒有的話，直接使用當前時間
                            PostST = preserve_Now;
                        }
                        
                    }
                    //非同步加工機台=>考量同機台前製程&同工單前製程
                    else
                    {
                        var Idx_order = 0;
                        var Idx_machine = 0;
                        //目前排程是否有同工單
                        if (result.Exists(x => x.OrderID == SortSchedule[i].OrderID))
                        {
                            //比較同機台和同工單最早可開工時間
                            // 找到最後一筆相同 OrderID 的索引
                            Idx_order = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                            Idx_machine = result.FindLastIndex(x => x.WorkGroup == SortSchedule[i].WorkGroup);

                            // 安全地取得 EndTime，若找不到則預設為 DateTime.MinValue
                            DateTime EndTime_order = (Idx_order != -1) ? result[Idx_order].EndTime : preserve_Now;
                            DateTime EndTime_machine = (Idx_machine != -1) ? result[Idx_machine].EndTime : preserve_Now;

                            // 取較晚的時間
                            PostST = (EndTime_order >= EndTime_machine) ? EndTime_order : EndTime_machine;

                            //Idx_order = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                            //Idx_machine = result.FindLastIndex(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                            //PostST = result[Idx_order].EndTime >= result[Idx_machine].EndTime ? result[Idx_order].EndTime : result[Idx_machine].EndTime;
                        }
                        //如果目前排程沒有，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //有的話，比較當前時間&同機台時間&同工單時間
                            DateTime time1 = ReportedOrder[SortSchedule[i].OrderID];
                            var foundMachine = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                            DateTime time2 = foundMachine != null ? foundMachine.EndTime : preserve_Now;
                            //DateTime time2 = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup).EndTime;

                            PostST = time1 >= time2 ? time1 : time2;
                        }
                        else
                        {
                            //沒有的話，直接使用同機台時間
                            var foundMachine = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                            PostST = foundMachine != null ? foundMachine.EndTime : preserve_Now;
                            //PostST = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup).EndTime;
                        }
                    }
                }
                //目前排程結果沒有同機台製程
                else
                {
                    //是同步加工機台 => 只需比較前製程完成時間和當前時間
                    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "1")//該機台為委外機台
                    {
                        //如果目前排程有同工單，則直接使用前製程時間
                        if (result.Exists(x => x.OrderID == SortSchedule[i].OrderID))
                        {
                            int idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                            if (idx >= 0)
                                PostST = result[idx].EndTime;
                            else
                                PostST = preserve_Now;
                            //Idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                            //PostST = result[Idx].EndTime;
                        }
                        //如果目前排程沒有，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //有的話，比較當前時間
                            PostST = preserve_Now > ReportedOrder[SortSchedule[i].OrderID] ? preserve_Now : ReportedOrder[SortSchedule[i].OrderID];
                        }
                        else
                        {
                            //沒有的話，直接使用當前時間
                            PostST = preserve_Now;
                        }
                    }
                    //非同步加工機台=>考量同機台前製程&同工單前製程
                    else
                    {
                        //如果目前排程有同工單
                        if (result.Exists(x => x.OrderID == SortSchedule[i].OrderID))
                        {
                            //如果原排程有同機台製程
                            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                            {
                                //比較同機台和同工單最早可開工時間
                                DateTime time1 = result.Find(x => x.OrderID == SortSchedule[i].OrderID).EndTime;
                                DateTime time2 = ReportedMachine[SortSchedule[i].OrderID];
                                PostST = time1 >= time2 ? time1 : time2;
                            }
                            //如果原排程沒有同機台製程
                            else
                            {
                                //直接使用前製程時間
                                PostST = result.Find(x => x.OrderID == SortSchedule[i].OrderID).EndTime;
                            } 
                        }
                        //如果目前排程沒有同工單，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //如果原排程有同機台製程
                            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                            {
                                //比較同機台、同工單、當前時間
                                DateTime time1 = preserve_Now;
                                DateTime time2 = ReportedOrder[SortSchedule[i].OrderID];
                                DateTime time3 = ReportedMachine[SortSchedule[i].OrderID];
                                if (time1 >= time2 && time1 >= time3)
                                    PostST = time1;
                                else if (time2 >= time1 && time2 >= time3)
                                    PostST = time2;
                                else
                                    PostST = time3;
                            }
                            //如果原排程沒有同機台製程，比較原排程同工單、當前時間
                            else
                            {
                                DateTime time1 = preserve_Now;
                                DateTime time2 = ReportedOrder[SortSchedule[i].OrderID];
                                PostST = time1 >= time2 ? time1 : time2;
                            }
                        }
                        else
                        {
                            //如果原排程有同機台，比較原排程同機台&當前時間
                            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                            {
                                DateTime time1 = preserve_Now;
                                DateTime time2 = ReportedMachine[SortSchedule[i].OrderID];
                                PostST = time1 >= time2 ? time1 : time2;
                            }
                            //如果原排程有沒同機台，直接用當前時間
                            else
                            {
                                PostST = preserve_Now;
                            }
                        }
                    }
                }
                #endregion


                #region 新的快取查找方式
                // 建立快取，讓查找更快，並直接儲存對應的物件
                orderIndexCache = result
                    .GroupBy(item => item.OrderID)
                    .ToDictionary(g => g.Key, g => g.OrderByDescending(x => x.EndTime).First());

                machineIndexCache = result
                    .GroupBy(item => item.WorkGroup)
                    .ToDictionary(g => g.Key, g => g.OrderByDescending(x => x.EndTime).First());


                //目前排程結果已有同機台製程
                if (machineIndexCache.ContainsKey(SortSchedule[i].WorkGroup) && OutsourcingList.Exists(x => x.remark == SortSchedule[i].WorkGroup))
                {
                    //是同步加工機台=>只需考量前製程完成時間
                    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "1")//該機台為委外機台
                    {
                        //目前排程是否有同工單
                        if (orderIndexCache.ContainsKey(SortSchedule[i].OrderID))
                        {
                            PostST = orderIndexCache[SortSchedule[i].OrderID].EndTime;
                        }
                        //如果目前排程沒有，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //有的話，比較當前時間
                            PostST = preserve_Now > ReportedOrder[SortSchedule[i].OrderID] ? DateTime.Now : ReportedOrder[SortSchedule[i].OrderID];
                        }
                        else
                        {
                            //沒有的話，直接使用當前時間
                            PostST = preserve_Now;
                        }
                    }
                    //非同步加工機台=>考量同機台前製程&同工單前製程
                    else
                    {
                        //目前排程是否有同工單
                        if (orderIndexCache.ContainsKey(SortSchedule[i].OrderID))
                        {
                            //比較同機台和同工單最早可開工時間
                            DateTime EndTime_order = orderIndexCache[SortSchedule[i].OrderID].EndTime;
                            DateTime EndTime_machine = machineIndexCache[SortSchedule[i].WorkGroup].EndTime;

                            // 取較晚的時間
                            PostST = (EndTime_order >= EndTime_machine) ? EndTime_order : EndTime_machine;
                        }
                        //如果目前排程沒有，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //有的話，比較當前時間&同機台時間&同工單時間
                            DateTime time1 = ReportedOrder[SortSchedule[i].OrderID];
                            DateTime time2 = machineIndexCache.ContainsKey(SortSchedule[i].WorkGroup)
                                ? machineIndexCache[SortSchedule[i].WorkGroup].EndTime
                                : preserve_Now;

                            PostST = time1 >= time2 ? time1 : time2;
                        }
                        else
                        {
                            //沒有的話，直接使用同機台時間
                            PostST = machineIndexCache.ContainsKey(SortSchedule[i].WorkGroup)
                                ? machineIndexCache[SortSchedule[i].WorkGroup].EndTime
                                : preserve_Now;
                        }
                    }
                }
                //目前排程結果沒有同機台製程
                else
                {
                    //是同步加工機台 => 只需比較前製程完成時間和當前時間
                    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "1")//該機台為委外機台
                    {
                        //如果目前排程有同工單，則直接使用前製程時間
                        if (orderIndexCache.ContainsKey(SortSchedule[i].OrderID))
                        {
                            PostST = orderIndexCache[SortSchedule[i].OrderID].EndTime;
                        }
                        //如果目前排程沒有，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //有的話，比較當前時間
                            PostST = preserve_Now > ReportedOrder[SortSchedule[i].OrderID] ? preserve_Now : ReportedOrder[SortSchedule[i].OrderID];
                        }
                        else
                        {
                            //沒有的話，直接使用當前時間
                            PostST = preserve_Now;
                        }
                    }
                    //非同步加工機台=>考量同機台前製程&同工單前製程
                    else
                    {
                        //如果目前排程有同工單
                        if (orderIndexCache.ContainsKey(SortSchedule[i].OrderID))
                        {
                            //如果原排程有同機台製程
                            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                            {
                                //比較同機台和同工單最早可開工時間
                                DateTime time1 = orderIndexCache[SortSchedule[i].OrderID].EndTime;
                                DateTime time2 = ReportedMachine[SortSchedule[i].OrderID];
                                PostST = time1 >= time2 ? time1 : time2;
                            }
                            //如果原排程沒有同機台製程
                            else
                            {
                                //直接使用前製程時間
                                PostST = orderIndexCache[SortSchedule[i].OrderID].EndTime;
                            }
                        }
                        //如果目前排程沒有同工單，則確認原排程是否有同工單製程
                        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            //如果原排程有同機台製程
                            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                            {
                                //比較同機台、同工單、當前時間
                                DateTime time1 = preserve_Now;
                                DateTime time2 = ReportedOrder[SortSchedule[i].OrderID];
                                DateTime time3 = ReportedMachine[SortSchedule[i].OrderID];
                                if (time1 >= time2 && time1 >= time3)
                                    PostST = time1;
                                else if (time2 >= time1 && time2 >= time3)
                                    PostST = time2;
                                else
                                    PostST = time3;
                            }
                            //如果原排程沒有同機台製程，比較原排程同工單、當前時間
                            else
                            {
                                DateTime time1 = preserve_Now;
                                DateTime time2 = ReportedOrder[SortSchedule[i].OrderID];
                                PostST = time1 >= time2 ? time1 : time2;
                            }
                        }
                        else
                        {
                            //如果原排程有同機台，比較原排程同機台&當前時間
                            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                            {
                                DateTime time1 = preserve_Now;
                                DateTime time2 = ReportedMachine[SortSchedule[i].OrderID];
                                PostST = time1 >= time2 ? time1 : time2;
                            }
                            //如果原排程有沒同機台，直接用當前時間
                            else
                            {
                                PostST = preserve_Now;
                            }
                        }
                    }
                }
                #endregion

                //補償休息時間
                //PostET = restTimecheck(PostST, ii.Duration);

                PostET = PostST + SortSchedule[i].Duration;

                result.Add(new Chromsome
                {
                    SeriesID = SortSchedule[i].SeriesID,
                    OrderID = SortSchedule[i].OrderID,
                    OPID = SortSchedule[i].OPID,
                    Range = SortSchedule[i].Range,
                    StartTime = PostST,
                    EndTime = PostET,
                    WorkGroup = SortSchedule[i].WorkGroup,
                    AssignDate = SortSchedule[i].PredictTime,
                    PartCount = SortSchedule[i].PartCount,
                    Maktx = SortSchedule[i].Maktx,
                    Duration = SortSchedule[i].Duration,
                    EachMachineSeq = SortSchedule[i].EachMachineSeq
                });
            }

            //篩選本次排程工單類別
            //var orderList = firstSchedule.OrderBy(x => x.EachMachineSeq).Select(x => x.OrderID)
            //                      .Distinct()
            //                      .ToList();
            var orderList = result.Distinct(x => x.OrderID)
                                  .Select(x => x.OrderID)
                                  .ToList();

            for (int k = 0; k < 2; k++)
            {
                foreach (var one_order in orderList)
                {

                    //挑選同工單製程
                    var temp = result.Where(x => x.OrderID == one_order)
                                     .OrderBy(x => x.Range)
                                     .ToList();

                    for (int i = 1; i < temp.Count; i++)
                    {
                        int idx;

                        //調整同工單製程
                        if (DateTime.Compare(Convert.ToDateTime(temp[i - 1].EndTime), Convert.ToDateTime(temp[i].StartTime)) > 0)
                        {
                            idx = result.FindIndex(x => x.OrderID == temp[i].OrderID && x.OPID == temp[i].OPID);
                            result[idx].StartTime = temp[i - 1].EndTime;
                            result[idx].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                            temp[i].StartTime = temp[i - 1].EndTime;
                            temp[i].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                        }
                        
                        if(OutsourcingList.Exists(x => x.remark == temp[i].WorkGroup))
                        {
                            //如果不是同步機台再調整同機台時間
                            if (OutsourcingList.Where(x => x.remark == temp[i].WorkGroup).First().isOutsource == "0")
                            {
                                //調整同機台製程
                                if (result.Exists(x => temp[i].WorkGroup == x.WorkGroup))
                                {
                                    var sequence = result.Where(x => x.WorkGroup == temp[i].WorkGroup)
                                                         .OrderBy(x => x.StartTime)
                                                         .ToList();
                                    for (int j = 1; j < sequence.Count; j++)
                                    {
                                        if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                        {
                                            idx = result.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                            result[idx].StartTime = sequence[j - 1].EndTime;
                                            result[idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
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

            CountDelay(result);
            return result;
        }

        private List<Chromsome> CheckOP(List<Chromsome> Data)
        {
            for (int i = 1; i < Data.Count; i++)
            {
                if (DateTime.Compare(Convert.ToDateTime(Data[i - 1].EndTime), Convert.ToDateTime(Data[i].StartTime)) != 0)
                {
                    TimeSpan TS = Convert.ToDateTime(Data[i].EndTime) - Convert.ToDateTime(Data[i].StartTime);
                    Data[i].StartTime = Data[i - 1].EndTime;
                    Data[i].EndTime = Convert.ToDateTime(Data[i].StartTime) + TS;
                }
            }
            return Data;
        }

        private int[] RandomNumberSeq(int count)
        {
            int[] arr = new int[count];
            for (int j = 0; j < count; j++)
            {
                arr[j] = j;
            }

            int[] arr2 = new int[count];
            int i = 0;
            int index;
            do
            {
                Random random = new Random();
                int r = count - i;
                index = random.Next(r);
                arr2[i++] = arr[index];
                arr[index] = arr[r - 1];
            } while (1 < count);

            return arr2;
        }
        //打亂List陣列
        public List<T> RandomSortList<T>(List<T> ListT)
        {

            System.Random random = new System.Random();
            List<T> newList = new List<T>();
            foreach (T item in ListT)
            {
                newList.Insert(random.Next(newList.Count), item);
            }
            return newList;
        }

        public List<MRP.Outsource> getOutsourcings()
        {
            string SqlStr = $@"SELECT a.*,b.Outsource
                              FROM {_ConnectStr.APSDB}.[dbo].Device as a
                              left join {_ConnectStr.APSDB}.[dbo].Outsourcing as b on a.ID=b.Id";
            List<MRP.Outsource> result = new List<MRP.Outsource>(); ;
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
                                result.Add(new MRP.Outsource
                                {
                                    ID = int.Parse(SqlData["ID"].ToString()),
                                    remark = SqlData["remark"].ToString(),
                                    isOutsource = SqlData["Outsource"].ToString(),
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        private bool IsOutsourcing_order(string orderid, string opid)
        {
            string SqlStr = $@"select d.* from Assignment as a
                            inner join {_ConnectStr.MRPDB}.dbo.ProcessDetial as b on a.OPID=b.ProcessID
                            inner join Device as c on b.MachineID=c.ID
                            inner join Outsourcing as d on c.ID=d.Id
                            where a.OrderID='{orderid}' and a.OPID={opid}";
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

        #region
        //获取当前周几

        private string _strWorkingDayAM = "08:00";//工作时间上午08:00
        private string _strWorkingDayPM = "17:00";
        private string _strRestDay = "6,7";//周几休息日 周六周日为 6,7

        private TimeSpan dspWorkingDayAM;//工作时间上午08:00
        private TimeSpan dspWorkingDayPM;

        private string m_GetWeekNow(DateTime date)
        {
            string strWeek = date.DayOfWeek.ToString();
            switch (strWeek)
            {
                case "Monday":
                    return "1";
                case "Tuesday":
                    return "2";
                case "Wednesday":
                    return "3";
                case "Thursday":
                    return "4";
                case "Friday":
                    return "5";
                case "Saturday":
                    return "6";
                case "Sunday":
                    return "7";
            }
            return "0";
        }


        /// <summary>
        /// 判断是否在工作日内
        /// </summary>
        /// <returns></returns>
        private bool m_IsWorkingDay(DateTime startTime)
        {
            string strWeekNow = this.m_GetWeekNow(startTime);//当前周几
            string[] RestDay = _strRestDay.Split(',');
            if (RestDay.Contains(strWeekNow))
            {
                return false;
            }
            //判断当前时间是否在工作时间段内

            dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;

            TimeSpan dspNow = startTime.TimeOfDay;
            if (dspNow > dspWorkingDayAM && dspNow < dspWorkingDayPM)
            {
                return true;
            }
            return false;
        }
        //初始化默认值
        private void m_InitWorkingDay()
        {
            dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;

        }
        #endregion

        private DateTime restTimecheck(DateTime PostST, TimeSpan Duration)
        {
            if (Duration > new TimeSpan(1, 00, 00, 00))
            {
                var days = Duration.TotalDays;
                TimeSpan resttime = new TimeSpan((int)(16 * days), 00, 00);
                Duration = Duration.Subtract(resttime);
            }
            const int hoursPerDay = 9;
            const int startHour = 8;
            // Don't start counting hours until start time is during working hours
            if (PostST.TimeOfDay.TotalHours > startHour + hoursPerDay)
                PostST = PostST.Date.AddDays(1).AddHours(startHour);
            if (PostST.TimeOfDay.TotalHours < startHour)
                PostST = PostST.Date.AddHours(startHour);
            if (PostST.DayOfWeek == DayOfWeek.Saturday)
                PostST.AddDays(2);
            else if (PostST.DayOfWeek == DayOfWeek.Sunday)
                PostST.AddDays(1);
            // Calculate how much working time already passed on the first day
            TimeSpan firstDayOffset = PostST.TimeOfDay.Subtract(TimeSpan.FromHours(startHour));
            // Calculate number of whole days to add
            var aaa = Duration.Add(firstDayOffset).TotalHours;
            int wholeDays = (int)(Duration.Add(firstDayOffset).TotalHours / hoursPerDay);
            // How many hours off the specified offset does this many whole days consume?
            TimeSpan wholeDaysHours = TimeSpan.FromHours(wholeDays * hoursPerDay);
            // Calculate the final time of day based on the number of whole days spanned and the specified offset
            TimeSpan remainder = Duration - wholeDaysHours;
            // How far into the week is the starting date?
            int weekOffset = ((int)(PostST.DayOfWeek + 7) - (int)DayOfWeek.Monday) % 7;
            // How many weekends are spanned?
            int weekends = (int)((wholeDays + weekOffset) / 5);
            // Calculate the final result using all the above calculated values
            return PostST.AddDays(wholeDays + weekends * 2).Add(remainder);
        }

        //原始函式
        //public void EvaluationFitness(ref Dictionary<int, List<Chromsome>> ChromosomeList, ref int noImprovementCount)
        //{
        //    var fitness_idx_value = new List<Evafitnessvalue>();
        //    var opt_ChromosomeList = new Dictionary<int, List<Chromsome>>();

        //    for (int i = 0; i < ChromosomeList.Count; i++)
        //    {
        //        int sumDelay = ChromosomeList[i].Sum(x => x.Delay);
        //        fitness_idx_value.Add(new Evafitnessvalue(i, sumDelay));
        //    }
        //    //計算適應度後排序，由小到大
        //    fitness_idx_value.Sort((x, y) => { return x.Fitness.CompareTo(y.Fitness); });
        //    //挑出前50%的染色體解答
        //    int chromosomeCount = Chromvalue / 2;
        //    for (int i = 0; i < chromosomeCount; i++)
        //    {

        //        opt_ChromosomeList.Add(i, ChromosomeList[fitness_idx_value[i].Idx].OrderBy(x => x.WorkGroup).ThenBy(x => x.StartTime).Select(x => x.Clone() as Chromsome)
        //                                                                          .ToList());
        //    }
        //    var random = new Random(Guid.NewGuid().GetHashCode());
        //    var crossoverResultList = new Dictionary<int, List<Chromsome>>();

        //    var crossoverList = new List<List<Chromsome>>();
        //    var crossoverTemp = new List<List<Chromsome>>();
        //    // opt_ChromosomeList 是前50%的母體資料 選兩個來做交換
        //    for (int i = 0; i < chromosomeCount; i++)
        //    {
        //        int randomNum = random.Next(0, chromosomeCount);
        //        crossoverList.Add(opt_ChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
        //        crossoverTemp.Add(opt_ChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
        //    }

        //    for (int childItem = 0; childItem < chromosomeCount; childItem++)
        //    {
        //        //crossover
        //        int cutLine = random.Next(1, crossoverList[0].Count);
        //        if (childItem < chromosomeCount - 1)
        //        {
        //            var swapData = crossoverList[childItem + 1].GetRange(cutLine, crossoverList[childItem + 1].Count - cutLine);
        //            crossoverTemp[childItem].RemoveRange(cutLine, crossoverList[childItem + 1].Count - cutLine);
        //            crossoverTemp[childItem].AddRange(new List<Chromsome>(swapData));

        //            swapData = crossoverList[childItem].GetRange(cutLine, crossoverList[childItem].Count - cutLine);
        //            crossoverTemp[childItem + 1].RemoveRange(cutLine, crossoverList[childItem].Count - cutLine);
        //            crossoverTemp[childItem + 1].AddRange(new List<Chromsome>(swapData));

        //            crossoverResultList.Add(2 * childItem, crossoverTemp[childItem]);
        //            crossoverResultList.Add(2 * childItem + 1, crossoverTemp[childItem + 1]);
        //        }
        //        else
        //        {
        //            var swapData = crossoverList[0].GetRange(cutLine, crossoverList[0].Count - cutLine);
        //            crossoverTemp[childItem].RemoveRange(cutLine, crossoverList[0].Count - cutLine);
        //            crossoverTemp[childItem].AddRange(new List<Chromsome>(swapData));

        //            swapData = crossoverList[childItem].GetRange(cutLine, crossoverList[childItem].Count - cutLine);
        //            crossoverTemp[0].RemoveRange(cutLine, crossoverList[childItem].Count - cutLine);
        //            crossoverTemp[0].AddRange(new List<Chromsome>(swapData));

        //            crossoverResultList.Add(2 * childItem, crossoverTemp[childItem]);
        //            crossoverResultList.Add(2 * childItem + 1, crossoverTemp[0]);
        //        }
        //    }
        //    InspectJobOper(crossoverResultList, ref ChromosomeList, fitness_idx_value.GetRange(0, crossoverList.Count), ref noImprovementCount);
        //}

        //2024.05.27 修改
        public void EvaluationFitness(ref Dictionary<int, List<Chromsome>> ChromosomeList, ref int noImprovementCount)
        {
            var fitness_idx_value = new List<Evafitnessvalue>();
            var localOptChromosomeList = new Dictionary<int, List<Chromsome>>(); // 使用局部变量存储结果
            var localChromosomeList = new Dictionary<int, List<Chromsome>>(ChromosomeList); // 创建局部变量

            // 使用并行计算计算适应度
            Parallel.ForEach(localChromosomeList, kvp =>
            {
                int sumDelay = kvp.Value.Sum(x => x.Delay);
                lock (fitness_idx_value)
                {
                    fitness_idx_value.Add(new Evafitnessvalue(kvp.Key, sumDelay));
                }
            });

            // 计算适应度后排序，由小到大
            fitness_idx_value.Sort((x, y) => x.Fitness.CompareTo(y.Fitness));

            // 挑出前50%的染色体解答
            int chromosomeCount = Chromvalue / 2;

            Parallel.For(0, chromosomeCount, i =>
            {
                var idx = fitness_idx_value[i].Idx;
                var orderedChromosomes = localChromosomeList[idx]
                    .OrderBy(x => x.WorkGroup)
                    .ThenBy(x => x.StartTime)
                    .Select(x => x.Clone() as Chromsome)
                    .ToList();

                lock (localOptChromosomeList)
                {
                    localOptChromosomeList[i] = orderedChromosomes;
                }
            });

            // 将局部变量更新到 ChromosomeList
            foreach (var item in localOptChromosomeList)
            {
                ChromosomeList[item.Key] = item.Value;
            }

            var random = new Random(Guid.NewGuid().GetHashCode());
            var crossoverResultList = new ConcurrentDictionary<int, List<Chromsome>>();

            var crossoverList = new List<List<Chromsome>>();
            var crossoverTemp = new List<List<Chromsome>>();

            // opt_ChromosomeList 是前50%的母体资料 选两个来做交换
            for (int i = 0; i < chromosomeCount; i++)
            {
                int randomNum = random.Next(0, chromosomeCount);
                crossoverList.Add(localOptChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
                crossoverTemp.Add(localOptChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
            }

            Parallel.For(0, chromosomeCount, childItem =>
            {
                // crossover
                int cutLine = random.Next(1, crossoverList[0].Count);

                if (childItem < chromosomeCount - 1)
                {
                    if (cutLine < crossoverList[childItem + 1].Count)
                    {
                        int length1 = crossoverList[childItem + 1].Count - cutLine;
                        if (length1 > 0)
                        {
                            var swapData1 = crossoverList[childItem + 1].GetRange(cutLine, length1);

                            lock (crossoverTemp)
                            {
                                crossoverTemp[childItem].RemoveRange(cutLine, length1);
                                crossoverTemp[childItem].AddRange(swapData1);
                            }

                            int length2 = crossoverList[childItem].Count - cutLine;
                            if (length2 > 0)
                            {
                                var swapData2 = crossoverList[childItem].GetRange(cutLine, length2);

                                lock (crossoverTemp)
                                {
                                    crossoverTemp[childItem + 1].RemoveRange(cutLine, length2);
                                    crossoverTemp[childItem + 1].AddRange(swapData2);
                                }
                            }
                        }
                    }

                    crossoverResultList[2 * childItem] = crossoverTemp[childItem];
                    crossoverResultList[2 * childItem + 1] = crossoverTemp[childItem + 1];
                }
                else
                {
                    if (cutLine < crossoverList[0].Count)
                    {
                        int length1 = crossoverList[0].Count - cutLine;
                        if (length1 > 0)
                        {
                            var swapData1 = crossoverList[0].GetRange(cutLine, length1);

                            lock (crossoverTemp)
                            {
                                crossoverTemp[childItem].RemoveRange(cutLine, length1);
                                crossoverTemp[childItem].AddRange(swapData1);
                            }

                            int length2 = crossoverList[childItem].Count - cutLine;
                            if (length2 > 0)
                            {
                                var swapData2 = crossoverList[childItem].GetRange(cutLine, length2);

                                lock (crossoverTemp)
                                {
                                    crossoverTemp[0].RemoveRange(cutLine, length2);
                                    crossoverTemp[0].AddRange(swapData2);
                                }
                            }
                        }
                    }

                    crossoverResultList[2 * childItem] = crossoverTemp[childItem];
                    crossoverResultList[2 * childItem + 1] = crossoverTemp[0];
                }
            });

            InspectJobOper(crossoverResultList, ref ChromosomeList, fitness_idx_value.GetRange(0, crossoverList.Count), ref noImprovementCount);
        }




        public void Mutation(List<Chromsome> scheduledData)
        {
            List<Chromsome> Datas = scheduledData.Select(x => x.Clone() as Chromsome).ToList();

            //倒序Chromsome內容(根據完工時間倒序排列)
            var temp2 = scheduledData.OrderByDescending(x => Convert.ToDateTime(x.EndTime))
                                     .Select(x => new { x.OrderID, x.OPID })
                                     .ToList();

            //取Chromsome最後一筆工單OrderID、OPID
            string keyOrderID = temp2[0].OrderID;
            double keyOPID = temp2[0].OPID;

            //找Chromsome內最早開工的時間
            DateTime minStartTime = scheduledData.Min(x => x.StartTime);

            //取得KeyOrder工單製程列表
            var data2 = scheduledData.Where(x => x.OrderID == keyOrderID /*&& x.OPID < keyOPID*/)
                                     .OrderBy(x => x.OPID)
                                     .ToList();

            //取得Chromsome最後一道製程資料
            var addData = scheduledData.Find(x => x.OrderID == keyOrderID && x.OPID == keyOPID);

            List<Chromsome> critpath = new List<Chromsome>();

            critpath = this.FindCriticalPath(scheduledData);



            Random random = new Random(Guid.NewGuid().GetHashCode());
            if (critpath.Count > 2)
            {
                int[] randomnums = { random.Next(0, critpath.Count), random.Next(0, critpath.Count) };
                while (randomnums[0] == randomnums[1])
                {
                    randomnums[1] = random.Next(0, critpath.Count);
                }
                if (critpath.Count > 2 && randomnums[0] != randomnums[1])
                {
                    int idx = scheduledData.FindIndex(x => x.OrderID == critpath[randomnums[0]].OrderID && x.OPID == critpath[randomnums[0]].OPID);
                    int idx2 = scheduledData.FindIndex(x => x.OrderID == critpath[randomnums[1]].OrderID && x.OPID == critpath[randomnums[1]].OPID);
                    var swap = Datas[idx];
                    var swap2 = Datas[idx2];
                    var orderList = new List<string>(scheduledData.Distinct(x => x.OrderID)
                                                                  .Select(x => x.OrderID)
                                                                  .ToList());
                    //製程互換
                    scheduledData[idx] = swap2.Clone() as Chromsome;
                    scheduledData[idx2] = swap.Clone() as Chromsome;

                    var duration1 = swap.EndTime - swap.StartTime;
                    var duration2 = swap2.EndTime - swap2.StartTime;

                    //更新互換後的機台和開始時間
                    scheduledData[idx].StartTime = swap.StartTime;
                    scheduledData[idx].WorkGroup = swap.WorkGroup;
                    scheduledData[idx].EndTime = swap.StartTime.Add(duration2);
                    scheduledData[idx2].WorkGroup = swap2.WorkGroup;
                    scheduledData[idx2].StartTime = swap2.StartTime;

                    scheduledData[idx2].EndTime = swap2.StartTime.Add(duration1);

                    var check = scheduledData.Distinct(x => x.WorkGroup)
                                             .Select(x => x.WorkGroup)
                                             .ToList();

                    //調整時間避免重疊
                    for (int k = 0; k < 2; k++)
                    {
                        foreach (var one_order in orderList)
                        {
                            //挑選同工單製程
                            var temp = scheduledData.Where(x => x.OrderID == one_order)
                                             .OrderBy(x => x.Range)
                                             .ToList();


                            #region 判斷是否為下班日or六日
                            //if (!m_IsWorkingDay(startTime))
                            //{
                            //    if (startTime > DateTime.Parse(startTime.ToShortDateString() + " 08:00"))
                            //    {
                            //        if (startTime.DayOfWeek == DayOfWeek.Saturday)
                            //            startTime = DateTime.Parse(startTime.AddDays(2).ToShortDateString() + " 08:00");
                            //        else if (startTime.DayOfWeek == DayOfWeek.Friday)
                            //            startTime = DateTime.Parse(startTime.AddDays(3).ToShortDateString() + " 08:00");
                            //        else
                            //            startTime = DateTime.Parse(startTime.AddDays(1).ToShortDateString() + " 08:00");
                            //    }
                            //    else
                            //        startTime = DateTime.Parse(startTime.ToShortDateString() + " 08:00");
                            //}
                            //else
                            //{
                            //    var s = startTime.DayOfWeek;
                            //}
                            #endregion

                            #region 有多道製程時
                            for (int i = 1; i < temp.Count; i++)
                            {
                                int indx=0;
                                //調整同工單製程
                                if (DateTime.Compare(Convert.ToDateTime(temp[i - 1].EndTime), Convert.ToDateTime(temp[i].StartTime)) > 0)
                                {
                                    indx = scheduledData.FindIndex(x => x.OrderID == temp[i].OrderID && x.OPID == temp[i].OPID);
                                    scheduledData[indx].StartTime = temp[i - 1].EndTime;
                                    scheduledData[indx].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                                    temp[i].StartTime = temp[i - 1].EndTime;
                                    temp[i].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                                }
                                //調整同機台製程
                                if (scheduledData.Exists(x => temp[i].WorkGroup == x.WorkGroup))
                                {
                                    var sequence = scheduledData.Where(x => x.WorkGroup == temp[i].WorkGroup)
                                                         .OrderBy(x => x.StartTime)
                                                         .ToList();
                                    for (int j = 1; j < sequence.Count; j++)
                                    {
                                        if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                        {
                                            indx = scheduledData.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                            scheduledData[indx].StartTime = sequence[j - 1].EndTime;
                                            scheduledData[indx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                            sequence[j].StartTime = sequence[j - 1].EndTime;
                                            sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 只有單一道製程時

                            ////調整同機台製程
                            //for (int i = 0; i < temp.Count; i++)
                            //{
                            //    if (scheduledData.Exists(x => temp[i].WorkGroup == x.WorkGroup))
                            //    {
                            //        var sequence = scheduledData.Where(x => x.WorkGroup == temp[i].WorkGroup)
                            //                             .OrderBy(x => x.StartTime)
                            //                             .ToList();
                            //        for (int j = 1; j < sequence.Count; j++)
                            //        {
                            //            if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                            //            {
                            //                int Idx = scheduledData.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                            //                scheduledData[Idx].StartTime = sequence[j - 1].EndTime;
                            //                scheduledData[Idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                            //                sequence[j].StartTime = sequence[j - 1].EndTime;
                            //                sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                            //            }
                            //        }
                            //    }
                            //}

                            #endregion
                        }
                    }
                }
                CountDelay(scheduledData);
            }

            void findLastTime(Chromsome temp, List<Chromsome> data)
            {
                var t1 = data.FindLast(x => x.WorkGroup == temp.WorkGroup && DateTime.Compare(x.EndTime, temp.StartTime) <= 0);

                DateTime OPET = data.FindLast(x => x.OrderID == temp.OrderID && x.OPID == temp.OPID).EndTime;

                if (t1 is null)
                {
                    temp.EndTime = temp.StartTime + temp.Duration;
                }
                else
                {
                    if (!(t1 is null) && DateTime.Compare(OPET, t1.EndTime) >= 0)
                    {
                        temp.StartTime = OPET;
                        temp.EndTime = OPET + temp.Duration;
                    }
                    else
                    {
                        if (DateTime.Compare(t1.StartTime, temp.StartTime) != 0)
                        {
                            temp.StartTime = t1.StartTime;
                            temp.EndTime = t1.StartTime + temp.Duration;
                        }
                        else
                        {
                            temp.EndTime = temp.EndTime = temp.StartTime + temp.Duration;
                        }
                    }
                }
            }
        }

        public void InspectJobOper_A(Dictionary<int, List<Chromsome>> crossoverResultList, ref Dictionary<int, List<Chromsome>> ChromosomeList, List<Evafitnessvalue> fitness_idx_value)
        {
            int total = ChromosomeList[0].Count;//製程總筆數
            for (int i = 0; i < crossoverResultList.Count; i++)
            {
                //比對自己有沒有重複指派工單的狀況，挑出未重複的至distinct_2
                var results = new List<Tuple<string, double>>();
                for (int j = 0; j < crossoverResultList[i].Count; j++)
                {
                    results.Add(Tuple.Create(crossoverResultList[i][j].OrderID, crossoverResultList[i][j].OPID));
                }
                List<Tuple<string, double>> distinct_2 = results.Distinct().ToList();
                var distinct_1 = new List<Tuple<string, double, string, TimeSpan, DateTime>>();
                //distinct_1把遺失的工單工序加回來
                if (distinct_2.Count != total)
                {
                    foreach (var item in ChromosomeList[i])
                    {
                        if (!distinct_2.Exists(x => x.Item1 == item.OrderID && x.Item2 == item.OPID))
                        {
                            distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, item.WorkGroup, item.Duration, item.AssignDate));
                            continue;
                        }
                        var query = crossoverResultList[i].Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                        distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, query.WorkGroup, query.Duration, query.AssignDate));
                    }
                }
                else
                {
                    distinct_1 = crossoverResultList[i].Select(x => Tuple.Create(x.OrderID, x.OPID, x.WorkGroup, x.Duration, x.AssignDate))
                                                       .ToList();
                }

                //每個機台的製程順序
                List<LocalMachineSeq> MachineSeq = new List<LocalMachineSeq>();
                for (int machinenameseq = 0; machinenameseq < Devices.Count; machinenameseq++)
                {
                    int seq = 0;
                    var ordersOnMachine = distinct_1.Where(x => x.Item3 == Devices[machinenameseq].Remark);
                    foreach (var item in ordersOnMachine)
                    {
                        MachineSeq.Add(new LocalMachineSeq
                        {
                            OrderID = item.Item1,
                            OPID = item.Item2,
                            WorkGroup = item.Item3,
                            Duration = item.Item4,
                            PredictTime = item.Item5,
                            EachMachineSeq = seq,
                        });
                        seq++;
                    }
                    if (MachineSeq.Count == distinct_1.Count)
                        break;
                }

                var tempOrder = Scheduled(MachineSeq);

                ////mutation(低於突變率才突變)
                //Random random = new Random(Guid.NewGuid().GetHashCode());
                //double prob = random.Next();
                //if(prob<0.05)
                //{
                //    Mutation(ref tempOrder);

                //}
                // 多一個比較sumdelay
                int sum = tempOrder.Sum(x => x.Delay);
                if (fitness_idx_value.Exists(x => x.Fitness > sum)) //判斷突變之後是否有更好的解
                {
                    //找到第一筆適應度較大的染色體，以突變後之染色體替換
                    int index = fitness_idx_value.FindIndex(x => x.Fitness > sum);
                    ChromosomeList.Remove(fitness_idx_value[index].Idx);
                    ChromosomeList.Add(fitness_idx_value[index].Idx, tempOrder.Select(x => (Chromsome)x.Clone())
                                                                              .ToList());
                    Debug.WriteLine($"delay is {fitness_idx_value[0].Fitness}");
                }
            }
        }

        //2024.05.27 修改
        public void InspectJobOper(ConcurrentDictionary<int, List<Chromsome>> crossoverResultList, ref Dictionary<int, List<Chromsome>> ChromosomeList, List<Evafitnessvalue> fitness_idx_value, ref int noImprovementCount)
        {
            bool HasImproved = false;
            var parallelOptions = new ParallelOptions() { MaxDegreeOfParallelism = Environment.ProcessorCount };
            var localChromosomeList = new Dictionary<int, List<Chromsome>>(ChromosomeList); // 创建局部变量
            var updatedChromosomes = new ConcurrentDictionary<int, List<Chromsome>>(); // 使用 ConcurrentDictionary

            Parallel.For(0, crossoverResultList.Count, parallelOptions, i =>
            {
                int total = localChromosomeList[i].Count; // 正确工单制程数
                var results = new List<Tuple<string, double>>();
                for (int j = 0; j < crossoverResultList[i].Count; j++)
                {
                    results.Add(Tuple.Create(crossoverResultList[i][j].OrderID, crossoverResultList[i][j].OPID));
                }
                List<Tuple<string, double>> distinct_2 = results.Distinct().ToList();
                var distinct_1 = new List<Tuple<string, double, string, TimeSpan, DateTime, int, string>>();
                // 把遗失的工单工序加回来
                if (distinct_2.Count != total)
                {
                    foreach (var item in localChromosomeList[i])
                    {
                        if (!distinct_2.Exists(x => x.Item1 == item.OrderID && x.Item2 == item.OPID))
                        {
                            distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, item.WorkGroup, item.Duration, item.AssignDate, item.Range, item.SeriesID));
                            continue;
                        }
                        var query = crossoverResultList[i].Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                        distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, query.WorkGroup, query.Duration, query.AssignDate, query.Range, item.SeriesID));
                    }
                }
                else
                {
                    distinct_1 = crossoverResultList[i].Select(x => Tuple.Create(x.OrderID, x.OPID, x.WorkGroup, x.Duration, x.AssignDate, x.Range, x.SeriesID))
                                                       .ToList();
                }
                // 重新给定机台排序
                List<LocalMachineSeq> MachineSeq = new List<LocalMachineSeq>();
                // 获取各机台是否为委外机台
                var OutsourcingList = getOutsourcings();
                for (int machinenameseq = 0; machinenameseq < Devices.Count; machinenameseq++)
                {
                    int seq = 0;
                    // 排序以OPID排，避免同工单后制程放在前制程前面
                    var ordersOnMachine = distinct_1.Where(x => x.Item3 == Devices[machinenameseq].Remark).OrderBy(x => x.Item2);
                    foreach (var item in ordersOnMachine)
                    {
                        if (OutsourcingList.Exists(x => x.remark == Devices[machinenameseq].Remark))
                        {
                            if (OutsourcingList.Where(x => x.remark == Devices[machinenameseq].Remark).First().isOutsource == "1")
                            {
                                MachineSeq.Add(new LocalMachineSeq
                                {
                                    OrderID = item.Item1,
                                    OPID = item.Item2,
                                    WorkGroup = item.Item3,
                                    Duration = item.Item4,
                                    PredictTime = item.Item5,
                                    PartCount = item.Item6,
                                    Range = item.Item6,
                                    EachMachineSeq = 0,
                                });
                            }
                            else
                            {
                                MachineSeq.Add(new LocalMachineSeq
                                {
                                    SeriesID = item.Item7,
                                    OrderID = item.Item1,
                                    OPID = item.Item2,
                                    WorkGroup = item.Item3,
                                    Duration = item.Item4,
                                    PredictTime = item.Item5,
                                    PartCount = item.Item6,
                                    Range = item.Item6,
                                    EachMachineSeq = seq,
                                });
                                seq++;
                            }
                        }
                    }
                    if (MachineSeq.Count == distinct_1.Count)
                    {
                        break;
                    }
                }
                // 重新排程
                var tempOrder = Scheduled(MachineSeq);
                // 多一个比较sumdelay
                int sum = tempOrder.Sum(x => x.Delay);
                if (fitness_idx_value.Exists(x => x.Fitness > sum)) // 判断突变之后是否有更好的解
                {
                    int index = fitness_idx_value.FindIndex(x => x.Fitness > sum);
                    updatedChromosomes[fitness_idx_value[index].Idx] = tempOrder.Select(x => (Chromsome)x.Clone()).ToList();
                    HasImproved = true;
                }
            });

            // 更新 ChromosomeList
            foreach (var kvp in updatedChromosomes)
            {
                ChromosomeList[kvp.Key] = kvp.Value;
            }

            if (!HasImproved)
            {
                noImprovementCount++;
            }
        }



        ////原始函式
        //public void InspectJobOper(Dictionary<int, List<Chromsome>> crossoverResultList, ref Dictionary<int, List<Chromsome>> ChromosomeList, List<Evafitnessvalue> fitness_idx_value,ref int noImprovementCount)
        //{
        //    bool HasImproved = false;
        //    for (int i = 0; i < crossoverResultList.Count; i++)
        //    {
        //        int total = ChromosomeList[i].Count;//正確工單製程數
        //        var results = new List<Tuple<string, double>>();
        //        for (int j = 0; j < crossoverResultList[i].Count; j++)
        //        {
        //            results.Add(Tuple.Create(crossoverResultList[i][j].OrderID, crossoverResultList[i][j].OPID));
        //        }
        //        List<Tuple<string, double>> distinct_2 = results.Distinct().ToList();
        //        var distinct_1 = new List<Tuple<string, double, string, TimeSpan, DateTime, int, string>>();
        //        //把遺失的工單工序加回來
        //        if (distinct_2.Count != total)
        //        {
        //            foreach (var item in ChromosomeList[i])
        //            {
        //                if (!distinct_2.Exists(x => x.Item1 == item.OrderID && x.Item2 == item.OPID))
        //                {
        //                    distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, item.WorkGroup, item.Duration, item.AssignDate, item.Range, item.SeriesID));
        //                    continue;
        //                }
        //                var query = crossoverResultList[i].Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
        //                distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, query.WorkGroup, query.Duration, query.AssignDate, query.Range, item.SeriesID));
        //            }
        //        }
        //        else
        //        {
        //            distinct_1 = crossoverResultList[i].Select(x => Tuple.Create(x.OrderID, x.OPID, x.WorkGroup, x.Duration, x.AssignDate, x.Range, x.SeriesID))
        //                                               .ToList();
        //        }
        //        //重新給定機台排序
        //        List<LocalMachineSeq> MachineSeq = new List<LocalMachineSeq>();
        //        //取得各機台是否為委外機台
        //        var OutsourcingList = getOutsourcings();
        //        for (int machinenameseq = 0; machinenameseq < Devices.Count; machinenameseq++)
        //        {
        //            int seq = 0;
        //            //排序以OPID排，避免同工單後製程放在前製程前面
        //            var ordersOnMachine = distinct_1.Where(x => x.Item3 == Devices[machinenameseq].Remark).OrderBy(x => x.Item2);
        //            foreach (var item in ordersOnMachine)
        //            {
        //                if(OutsourcingList.Exists(x=>x.remark== Devices[machinenameseq].Remark))
        //                {
        //                    if (OutsourcingList.Where(x => x.remark == Devices[machinenameseq].Remark).First().isOutsource == "1")
        //                    {
        //                        MachineSeq.Add(new LocalMachineSeq
        //                        {
        //                            OrderID = item.Item1,
        //                            OPID = item.Item2,
        //                            WorkGroup = item.Item3,
        //                            Duration = item.Item4,
        //                            PredictTime = item.Item5,
        //                            PartCount = item.Item6,
        //                            Range = item.Item6,
        //                            EachMachineSeq = 0,
        //                        });
        //                    }
        //                    else
        //                    {
        //                        MachineSeq.Add(new LocalMachineSeq
        //                        {
        //                            SeriesID = item.Item7,
        //                            OrderID = item.Item1,
        //                            OPID = item.Item2,
        //                            WorkGroup = item.Item3,
        //                            Duration = item.Item4,
        //                            PredictTime = item.Item5,
        //                            PartCount = item.Item6,
        //                            Range = item.Item6,
        //                            EachMachineSeq = seq,
        //                        });
        //                        seq++;
        //                    }
        //                }



        //            }
        //            if (MachineSeq.Count == distinct_1.Count)
        //            {
        //                break;
        //            }
        //        }
        //        //重新排程
        //        var tempOrder = Scheduled(MachineSeq);
        //        ////突變mutation
        //        //Random rand = new Random(Guid.NewGuid().GetHashCode());
        //        //if(rand.NextDouble()<0.05)
        //        //{
        //        //    Mutation(tempOrder);
        //        //}

        //        // 多一個比較sumdelay
        //        int sum = tempOrder.Sum(x => x.Delay);
        //        if (fitness_idx_value.Exists(x => x.Fitness > sum)) //判斷突變之後是否有更好的解
        //        {
        //            int index = fitness_idx_value.FindIndex(x => x.Fitness > sum);
        //            ChromosomeList.Remove(fitness_idx_value[index].Idx);
        //            ChromosomeList.Add(fitness_idx_value[index].Idx, tempOrder.Select(x => (Chromsome)x.Clone())
        //                                                                      .ToList());
        //            Debug.WriteLine($"delay is {fitness_idx_value[0].Fitness}");
        //            HasImproved = true;
        //        }
        //    }
        //    if(HasImproved==false)
        //    {
        //        noImprovementCount += 1;
        //    }
        //}

        public void CountDelay(List<Chromsome> Tep)
        {
            TimeSpan temp;
            int itemDelay;
            foreach (var item in Tep)
            {
                try
                {
                    temp = item.AssignDate - item.EndTime;

                    itemDelay = (temp.TotalDays > 0) ? 0 : Math.Abs(temp.Days);
                    if (itemDelay != 0) ;
                    item.Delay = itemDelay;
                }
                catch
                {
                    continue;
                }
            }
        }


        public class OrderDevice
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
            /// 機台名稱
            /// </summary>
            public string DeviceName { get; set; }
        }

        public List<Chromsome> DispatchCreateDataSet(List<string> OrderList, DateTime activetime)
        {
            Dictionary<string, TimeSpan> inserttemp = new Dictionary<string, TimeSpan>();
            # region 修改插單預交日期=>所有製程工時加總*2
            foreach (var item in OrderList)
            {
                string tempstr = @$"SELECT a.OrderID, a.OPID, a.OrderQTY, a.HumanOpTime, a.MachOpTime, a.AssignDate, a.AssignDate_PM, a.MAKTX,wip.WIPEvent
                                FROM Assignment a left join WIP as wip
                                on a.OrderID=wip.OrderID and a.OPID=wip.OPID
                                left join WipRegisterLog w
                                on w.WorkOrderID = a.OrderID and w.OPID=a.OPID
                                where w.WorkOrderID is NULL and (wip.WIPEvent!=3 or wip.WIPEvent is NULL) and a.OrderID = @OrderID 
                                order by a.OrderID, a.Range";
                using (var Conn = new SqlConnection(_ConnectStr.Local))
                {
                    using (SqlCommand Comm = new SqlCommand(tempstr, Conn))
                    {
                        if (Conn.State != ConnectionState.Open)
                            Conn.Open();
                        Comm.Parameters.Add("@OrderID", SqlDbType.VarChar).Value = item;
                        using (SqlDataReader SqlData = Comm.ExecuteReader())
                        {
                            if (SqlData.HasRows)
                            {
                                while (SqlData.Read())
                                {
                                    if (inserttemp.ContainsKey(item))
                                    {
                                        inserttemp[item] += new TimeSpan(0, (int)(Convert.ToDouble(SqlData["HumanOpTime"]) + Convert.ToInt32(SqlData["OrderQTY"]) *
                                                                    Convert.ToDouble(SqlData["MachOpTime"])) * 2, 0);
                                    }
                                    else
                                    {
                                        inserttemp.Add(item, new TimeSpan(0, (int)(Convert.ToDouble(SqlData["HumanOpTime"]) + Convert.ToInt32(SqlData["OrderQTY"]) *
                                                                    Convert.ToDouble(SqlData["MachOpTime"])) * 2, 0));
                                    }

                                }
                            }
                        }
                    }
                }
            }
            #endregion
            TimeSpan actDuration = new TimeSpan();
            //撈出原排程未開工&未排程之製程(排程時間可調整)
            string sqlStr = @$"SELECT a.OrderID, a.OPID, p.CanSync, a.Range,a.OrderQTY, a.HumanOpTime, a.MachOpTime, a.OrderQTY, a.StartTime, a.EndTime, a.AssignDate, a.WorkGroup, a.MAKTX
                            FROM {_ConnectStr.APSDB}.dbo.Assignment a left join {_ConnectStr.APSDB}.dbo.WIP as wip
                            on a.OrderID=wip.OrderID and a.OPID=wip.OPID
                            left join {_ConnectStr.APSDB}.dbo.WipRegisterLog w
                            on w.WorkOrderID = a.OrderID and w.OPID=a.OPID
                            inner join {_ConnectStr.MRPDB}.dbo.Process as p 
                            on a.OPID=p.ID
                            where w.WorkOrderID is NULL and (wip.WIPEvent=0 or wip.WIPEvent is NULL) and (a.StartTime >=@activetime or a.StartTime is null)
                            order by a.WorkGroup, a.StartTime";
            var result = new List<Chromsome>();
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                using (SqlCommand Comm = new SqlCommand(sqlStr, Conn))
                {
                    if (Conn.State != ConnectionState.Open)
                        Conn.Open();
                    Comm.Parameters.Add("@activetime", SqlDbType.DateTime).Value = activetime;
                    using (SqlDataReader SqlData = Comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            while (SqlData.Read())
                            {
                                if (Convert.ToInt16(SqlData["CanSync"]) == 0)
                                {
                                    actDuration = new TimeSpan(0, (int)(Convert.ToDouble(SqlData["HumanOpTime"]) + (Convert.ToInt32(SqlData["OrderQTY"]) *
                                                          Convert.ToDouble(SqlData["MachOpTime"]))), 0);
                                }
                                else
                                {
                                    actDuration = new TimeSpan(0, (int)(Convert.ToDouble(SqlData["HumanOpTime"]) +
                                                          Convert.ToDouble(SqlData["MachOpTime"])), 0);
                                }
                                //若該工單為插單，修改預交日期與派工機台
                                if (OrderList.Exists(x => x == SqlData["OrderID"].ToString().Trim()))
                                {
                                    result.Add(new Chromsome
                                    {
                                        OrderID = SqlData["OrderID"].ToString().Trim(),
                                        OPID = Convert.ToDouble(SqlData["OPID"].ToString()),
                                        Range = Convert.ToInt32(SqlData["Range"]),
                                        //Duration = new TimeSpan(0, Convert.ToInt32(SqlData["Optime"]), 0),
                                        Duration = actDuration,
                                        AssignDate = DateTime.Now + inserttemp[SqlData["OrderID"].ToString().Trim()],
                                        Maktx = SqlData["MAKTX"].ToString().Trim(),
                                        WorkGroup = string.Empty,
                                        PartCount = Convert.ToInt32(SqlData["OrderQTY"].ToString())

                                    }); ;
                                }
                                //若該工單未在原排程但也不為插單就不放入資料集當中
                                else if (!OrderList.Exists(x => x == SqlData["OrderID"].ToString().Trim()) && SqlData["StartTime"].ToString() == "")
                                {
                                    continue;
                                }
                                //為原排程之工單
                                else
                                {
                                    result.Add(new Chromsome
                                    {
                                        OrderID = SqlData["OrderID"].ToString().Trim(),
                                        OPID = Convert.ToDouble(SqlData["OPID"].ToString()),
                                        Range = Convert.ToInt32(SqlData["Range"]),
                                        Duration = actDuration,
                                        StartTime = Convert.ToDateTime(SqlData["StartTime"].ToString()),
                                        EndTime = Convert.ToDateTime(SqlData["EndTime"].ToString()),
                                        AssignDate = Convert.ToDateTime(SqlData["AssignDate"].ToString()),
                                        Maktx = SqlData["MAKTX"].ToString().Trim(),
                                        WorkGroup = SqlData["WorkGroup"].ToString().Trim(),
                                        PartCount = Convert.ToInt32(SqlData["OrderQTY"].ToString())
                                    });
                                }

                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 計算插單排程適應度
        /// </summary>
        /// <param name="oridata">原排程</param>
        /// <param name="newdata">新排程</param>
        /// <returns></returns>
        public double DispatchFitness(List<Chromsome> oridata, List<Chromsome> newdata)
        {
            var orderidlist = newdata.Select(x => x.OrderID).Distinct().ToList();
            TimeSpan temp;
            double itemDelay;
            double totaldelay = 0;
            double totaldiff = 0;
            foreach (string order in orderidlist)
            {

                var sameorder = newdata.Where(x => x.OrderID == order).OrderByDescending(x => x.EndTime).ToList();
                //算延遲時間
                temp = sameorder[0].AssignDate - sameorder[0].EndTime;
                itemDelay = (temp.TotalDays > 0) ? 0 : Math.Abs(temp.TotalDays);
                if (!oridata.Exists(x => x.OrderID == order))
                {
                    //totaldelay += 5 * (itemDelay / (sameorder[0].EndTime - sameorder[sameorder.Count - 1].StartTime).TotalDays);
                    totaldelay += 1 * itemDelay;
                }
                else
                {
                    //totaldelay += 2.5 * (itemDelay / (sameorder[0].EndTime - sameorder[sameorder.Count - 1].StartTime).TotalDays);
                    totaldelay += 2 * itemDelay;
                }


                //算異動時間
                if (oridata.Exists(x => x.OrderID == order))
                {
                    double orderdiff = 0;
                    var orisch = oridata.Where(x => x.OrderID == order).OrderByDescending(x => x.StartTime).ToList();
                    var maxtime = sameorder[0].EndTime >= orisch[0].EndTime ? sameorder[0].EndTime : orisch[0].EndTime;
                    var mintime = sameorder[sameorder.Count() - 1].StartTime <= orisch[orisch.Count() - 1].StartTime ? sameorder[sameorder.Count() - 1].StartTime : orisch[orisch.Count() - 1].StartTime;

                    //var flowtime = (sameorder[0].EndTime - orisch[0].StartTime).TotalDays;//flow time取新排程結束-舊排程開始
                    var flowtime = (maxtime - mintime).TotalDays;

                    foreach (var item in sameorder)
                    {
                        int idx = oridata.FindIndex(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                        orderdiff += Math.Abs((item.StartTime - oridata[idx].StartTime).TotalDays);
                    }
                    //totaldiff += (orderdiff / flowtime);
                    totaldiff += orderdiff;
                }

            }
            var newms = newdata.Select(x => x.EndTime).Max();
            var newlastop = newdata.Where(x => x.EndTime == newms).ToList();
            var orims = oridata.Select(x => x.EndTime).Max();
            var orilastop = oridata.Where(x => x.EndTime == orims).ToList();
            var PCvalue = totaldelay / orderidlist.Count();
            var STvalue = totaldiff / oridata.Select(x => x.OrderID).Distinct().Count();
            var newcmax = (newms > orims) ? newms : orims;
            //var D_MSvalue = (newcmax - orims).TotalDays / (newms - oridata.Select(x => x.StartTime).Min()).TotalDays;
            //var D_MSvalue = (newcmax - orims).TotalDays / (newcmax-DateTime.Now).TotalDays;
            var lastbegin = newlastop[0].StartTime <= orilastop[0].StartTime ? newlastop[0].StartTime : orilastop[0].StartTime;
            //var D_MSvalue = (newcmax - orims).TotalDays / (newcmax - lastbegin).TotalDays;
            var D_MSvalue = (newcmax - orims).TotalDays;

            var fitness = (0.3 * PCvalue + 0.4 * STvalue + 0.3 * D_MSvalue);


            return fitness;

        }

        /// <summary>
        /// 找關鍵路徑
        /// </summary>
        /// <param name="Inputschedule"></param>
        /// <returns></returns>
        public List<Chromsome> FindCriticalPath(List<Chromsome> Inputschedule)
        {
            Inputschedule = Inputschedule.OrderBy(x => x.WorkGroup).ThenByDescending(x => x.StartTime).ToList();
            var makespan = Inputschedule.Max(x => x.EndTime);
            var begin = Inputschedule.Find(x => x.EndTime == makespan);
            string orderid;
            double opid;
            var result = new List<Chromsome>();
            result.Add(begin);
            while (true)
            {
                orderid = begin.OrderID;
                opid = begin.OPID;
                var sameod = Inputschedule.Find(x => x.OrderID == orderid && x.OPID == opid - 1);
                var samewg = Inputschedule.Find(x => x.WorkGroup == begin.WorkGroup && x.EndTime <= begin.StartTime);
                if (Inputschedule.Exists(x => x.OrderID == orderid && x.OPID == opid - 1) && Inputschedule.Exists(x => x.WorkGroup == begin.WorkGroup && x.EndTime <= begin.StartTime))
                {
                    if (sameod.EndTime > samewg.EndTime)
                    {
                        result.Add(sameod);
                        begin = sameod;
                    }
                    else
                    {
                        result.Add(samewg);
                        begin = samewg;
                    }
                }
                else if (Inputschedule.Exists(x => x.OrderID == orderid && x.OPID == opid - 1))
                {
                    result.Add(sameod);
                    begin = sameod;
                }
                else if (Inputschedule.Exists(x => x.WorkGroup == begin.WorkGroup && x.EndTime <= begin.StartTime))
                {
                    result.Add(samewg);
                    begin = samewg;
                }
                else
                {
                    break;
                }
            }
            return result;
        }
    }

    internal class SetupMethod : IMAthModel
    {
        ConnectStr _ConnectStr = new ConnectStr();
        public int Chromvalue { get; set; }
        private List<Device> Devices { get; set; }
        private DateTime PresetStartTime { get; set; } = DateTime.Now;
        public Dictionary<string, DateTime> ReportedMachine { get; set; }
        public Dictionary<string, DateTime> ReportedOrder { get; set; }

        public SetupMethod(int chromvalue, List<Device> devices)
        {
            Chromvalue = chromvalue;
            Devices = new List<Device>(devices);
        }

        public List<GaSchedule> CreateDataSet()
        {
            TimeSpan actDuration = new TimeSpan();
            string SqlStr = @$"SELECT a.SeriesID, a.OrderID, a.OPID, p.CanSync, a.Range, a.OrderQTY, a.HumanOpTime, a.MachOpTime, a.AssignDate, a.AssignDate_PM,a.MAKTX,wip.WIPEvent
                                FROM {_ConnectStr.APSDB}.dbo.Assignment a left join {_ConnectStr.APSDB}.dbo.WIP as wip
                                on a.OrderID=wip.OrderID and a.OPID=wip.OPID and a.SeriesID=wip.SeriesID
                                left join {_ConnectStr.APSDB}.dbo.WipRegisterLog w
                                on w.WorkOrderID = a.OrderID and w.OPID=a.OPID
                                inner join {_ConnectStr.MRPDB}.dbo.Process as p 
                                on a.OPID=p.ID
                                where w.WorkOrderID is NULL and (wip.WIPEvent=0 or wip.WIPEvent is NULL)
                                order by a.OrderID, a.Range";
            var result = new List<GaSchedule>();
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
                                try
                                {
                                    if(Convert.ToInt16(SqlData["CanSync"])==0)
                                    {
                                        actDuration = new TimeSpan(0, (int)(Convert.ToDouble(SqlData["HumanOpTime"]) + (Convert.ToInt32(SqlData["OrderQTY"]) *
                                                              Convert.ToDouble(SqlData["MachOpTime"]))), 0);
                                    }
                                    else
                                    {
                                        actDuration = new TimeSpan(0, (int)(Convert.ToDouble(SqlData["HumanOpTime"]) + 
                                                              Convert.ToDouble(SqlData["MachOpTime"])), 0);
                                    }
                                    result.Add(new GaSchedule
                                    {
                                        SeriesID = SqlData["SeriesID"].ToString().Trim(),
                                        Range = int.Parse(SqlData["Range"].ToString().Trim()),
                                        PartCount = int.Parse(SqlData["OrderQTY"].ToString().Trim()),
                                        OrderID = SqlData["OrderID"].ToString().Trim(),
                                        OPID = Convert.ToDouble(SqlData["OPID"].ToString()),
                                        Duration = actDuration,
                                        Assigndate = Convert.ToDateTime(SqlData["AssignDate_PM"]),
                                        Maktx = SqlData["MAKTX"].ToString()
                                    });
                                }
                                catch
                                {
                                    continue;
                                }
                            }
                        }
                    }
                }
            }
            return result;
        }

        public List<LocalMachineSeq> CreateSequence(List<GaSchedule> dataSet)
        {
            Random rnd = new Random(Guid.NewGuid().GetHashCode());
            var result = new List<LocalMachineSeq>();
            var totalOrders = dataSet.Distinct(x => x.OrderID)
                                     .Select(x => x.OrderID);
            foreach (var totalOrder in totalOrders)
            {
                MakeSequence(totalOrder);
            }
            return result;

            void MakeSequence(string orderID)
            {

                var orderSequences = dataSet.FindAll(x => x.OrderID == orderID).OrderBy(x => x.OPID);

                //取得所有製程的替代機台
                var ProcessDetial = getProcessDetial();

                //取得各機台是否為委外機台
                var OutsourcingList = getOutsourcings();

                int j = 0;
                int machineSeq = 0;
                string chosmach = String.Empty;
                foreach (var orderSequence in orderSequences)
                {
                    //取得該工單可以分發的機台列表，若MRP Table內沒有相關資料可能會找不到可用機台，應該要回傳錯誤訊息，此次排成失敗
                    var CanUseDevices = ProcessDetial.Where(x => x.ProcessID == orderSequence.OPID.ToString()).ToList();

                    if (CanUseDevices.Count != 0)
                    {
                        //若該製程可用機台以有前面製程使用，則該製程也分派至該機台
                        if (result.Exists(x => x.OrderID == orderID && CanUseDevices.Exists(y => y.remark == x.WorkGroup)))
                        {
                            //指定可使用機台為重複機台
                            var temp = result.FindLast(x => x.OrderID == orderID && CanUseDevices.Exists(y => y.remark == x.WorkGroup));
                            chosmach = temp.WorkGroup;
                        }
                        else
                        {
                            j = rnd.Next(0, CanUseDevices.Count); //在可用之機台中產生機台編號
                            chosmach = CanUseDevices[j].remark;
                        }

                        if(OutsourcingList.Exists(x=>x.remark==chosmach))
                        {
                            if(OutsourcingList.Where(x => x.remark == chosmach).First().isOutsource=="0")
                            {
                                if (result.Exists(x => x.WorkGroup == chosmach))
                                {
                                    machineSeq = result.Where(x => x.WorkGroup == chosmach)
                                                     .Select(x => x.EachMachineSeq)
                                                     .Max() + 1;
                                }
                                else
                                {
                                    machineSeq = 0;
                                }
                            }
                            else
                            {
                                machineSeq = 0;
                            }
                            result.Add(new LocalMachineSeq
                            {
                                SeriesID = orderSequence.SeriesID,
                                OrderID = orderSequence.OrderID,
                                OPID = orderSequence.OPID,
                                Range = orderSequence.Range,
                                Duration = orderSequence.Duration,
                                PredictTime = orderSequence.Assigndate,
                                Maktx = orderSequence.Maktx,
                                PartCount = orderSequence.PartCount,
                                WorkGroup = chosmach,
                                EachMachineSeq = machineSeq

                            });
                        }
                    }
                }
            }
        }

        private List<MRP.ProcessDetial> getProcessDetial()
        {
            List<MRP.ProcessDetial> result = new List<MRP.ProcessDetial>();
            string SqlStr = "";
            SqlStr = $@"
                        SELECT a.*,b.remark 
                          FROM {_ConnectStr.MRPDB}.[dbo].[ProcessDetial] as a
                          left join {_ConnectStr.APSDB}.[dbo].[Device] as b on a.MachineID=b.ID
                          order by a.ProcessID,b.ID
                        ";
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
                                result.Add(new MRP.ProcessDetial
                                {
                                    ID = int.Parse(SqlData["ID"].ToString()),
                                    ProcessID = string.IsNullOrEmpty(SqlData["ProcessID"].ToString()) ? "" : SqlData["ProcessID"].ToString(),
                                    MachineID = string.IsNullOrEmpty(SqlData["MachineID"].ToString()) ? "" : SqlData["MachineID"].ToString(),
                                    remark = string.IsNullOrEmpty(SqlData["remark"].ToString()) ? "" : SqlData["remark"].ToString(),
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        private List<Device> getCanUseDevice(string OrderID, string OPID)
        {
            List<Device> devices = new List<Device>();
            string SqlStr = @$"SELECT dd.* FROM Assignment as aa
                                inner join (SELECT a.Number,a.Name,a.RoutingID,b.ProcessRang,c.ID,c.ProcessNo,c.ProcessName FROM {_ConnectStr.MRPDB}.dbo.Part as a
                                inner join {_ConnectStr.MRPDB}.dbo.RoutingDetail as b on a.RoutingID=b.RoutingId
                                inner join {_ConnectStr.MRPDB}.dbo.Process as c on b.ProcessId=c.ID
                                where a.Number= (select top(1) MAKTX from Assignment where OrderID=@OrderID and OPID=@OPID) ) as bb on aa.MAKTX=bb.Number and aa.OPID=bb.ID
                                left join {_ConnectStr.MRPDB}.dbo.ProcessDetial as cc on bb.ID=cc.ProcessID
                                inner join Device as dd on cc.MachineID=dd.ID
                                where aa.OrderID=@OrderID and aa.OPID=@OPID";
            using (var Conn = new SqlConnection(_ConnectStr.Local))
            {
                if (Conn.State != ConnectionState.Open)
                    Conn.Open();

                using (var Comm = new SqlCommand(SqlStr, Conn))
                {
                    Comm.Parameters.Add(("@OrderID"), SqlDbType.NVarChar).Value = OrderID;
                    Comm.Parameters.Add(("@OPID"), SqlDbType.Float).Value = OPID;
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

            if (devices.Count == 0) ;

            return devices;
        }

        public void EvaluationFitness(ref Dictionary<int, List<Chromsome>> ChromosomeList,ref int noImprovementCount)
        {
            

            var fitness_idx_value = new List<Evafitnessvalue>();
            var opt_ChromosomeList = new Dictionary<int, List<Chromsome>>();

            for (int i = 0; i < ChromosomeList.Count; i++)
            {
                int sumDelay = ChromosomeList[i].Sum(x => x.Delay);
                fitness_idx_value.Add(new Evafitnessvalue(i, sumDelay));
            }
            //計算適應度後排序，由小到大
            fitness_idx_value.Sort((x, y) => { return x.Fitness.CompareTo(y.Fitness); });
            //挑出前50%的染色體解答
            int chromosomeCount = Chromvalue / 2;
            for (int i = 0; i < chromosomeCount; i++)
            {
                //opt_ChromosomeList.Add(
                //    i,
                //    ChromosomeList[fitness_idx_value[i].Idx].Select(x => x.Clone() as Chromsome).ToList()
                //    );
                opt_ChromosomeList.Add(i, ChromosomeList[fitness_idx_value[i].Idx].OrderBy(x => x.WorkGroup).ThenBy(x => x.StartTime).Select(x => x.Clone() as Chromsome)
                                                                                  .ToList());
            }
            var random = new Random(Guid.NewGuid().GetHashCode());
            var crossoverResultList = new Dictionary<int, List<Chromsome>>();

            var crossoverList = new List<List<Chromsome>>();
            var crossoverTemp = new List<List<Chromsome>>();
            // opt_ChromosomeList 是前50%的母體資料 選兩個來做交換
            for (int i = 0; i < chromosomeCount; i++)
            {
                int randomNum = random.Next(0, chromosomeCount);
                crossoverList.Add(opt_ChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
                crossoverTemp.Add(opt_ChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
            }

            for (int childItem = 0; childItem < chromosomeCount; childItem++)
            {
                //crossover
                int cutLine = random.Next(1, crossoverList[0].Count);
                if (childItem < chromosomeCount - 1)
                {
                    var swapData = crossoverList[childItem + 1].GetRange(cutLine, crossoverList[childItem + 1].Count - cutLine);
                    crossoverTemp[childItem].RemoveRange(cutLine, crossoverList[childItem + 1].Count - cutLine);
                    crossoverTemp[childItem].AddRange(new List<Chromsome>(swapData));

                    swapData = crossoverList[childItem].GetRange(cutLine, crossoverList[childItem].Count - cutLine);
                    crossoverTemp[childItem + 1].RemoveRange(cutLine, crossoverList[childItem].Count - cutLine);
                    crossoverTemp[childItem + 1].AddRange(new List<Chromsome>(swapData));

                    crossoverResultList.Add(2 * childItem, crossoverTemp[childItem]);
                    crossoverResultList.Add(2 * childItem + 1, crossoverTemp[childItem + 1]);
                }
                else
                {
                    var swapData = crossoverList[0].GetRange(cutLine, crossoverList[0].Count - cutLine);
                    crossoverTemp[childItem].RemoveRange(cutLine, crossoverList[0].Count - cutLine);
                    crossoverTemp[childItem].AddRange(new List<Chromsome>(swapData));

                    swapData = crossoverList[childItem].GetRange(cutLine, crossoverList[childItem].Count - cutLine);
                    crossoverTemp[0].RemoveRange(cutLine, crossoverList[childItem].Count - cutLine);
                    crossoverTemp[0].AddRange(new List<Chromsome>(swapData));

                    crossoverResultList.Add(2 * childItem, crossoverTemp[childItem]);
                    crossoverResultList.Add(2 * childItem + 1, crossoverTemp[0]);
                }
            }
            InspectJobOper(crossoverResultList, ref ChromosomeList, fitness_idx_value.GetRange(0, crossoverList.Count), ref noImprovementCount);
        }

        public void Mutation(ref List<Chromsome> scheduledData)
        {
            List<Chromsome> Datas = scheduledData.Select(x => x.Clone() as Chromsome).ToList();

            //倒序Chromsome內容(根據完工時間倒序排列)
            var temp2 = scheduledData.OrderByDescending(x => Convert.ToDateTime(x.EndTime))
                                     .Select(x => new { x.OrderID, x.OPID })
                                     .ToList();

            //取Chromsome最後一筆工單OrderID、OPID
            string keyOrderID = temp2[0].OrderID;
            double keyOPID = temp2[0].OPID;

            //找Chromsome內最早開工的時間
            DateTime minStartTime = scheduledData.Min(x => x.StartTime);

            //取得KeyOrder工單製程列表
            var data2 = scheduledData.Where(x => x.OrderID == keyOrderID /*&& x.OPID < keyOPID*/)
                                     .OrderBy(x => x.OPID)
                                     .ToList();

            //取得Chromsome最後一道製程資料
            var addData = scheduledData.Find(x => x.OrderID == keyOrderID && x.OPID == keyOPID);

            List<Chromsome> critpath = new List<Chromsome>();

            critpath = this.FindCriticalPath(scheduledData);



            Random random = new Random(Guid.NewGuid().GetHashCode());
            if (critpath.Count > 2)
            {
                int[] randomnums = { random.Next(0, critpath.Count), random.Next(0, critpath.Count) };
                while (randomnums[0] == randomnums[1])
                {
                    randomnums[1] = random.Next(0, critpath.Count);
                }
                if (critpath.Count > 2 && randomnums[0] != randomnums[1])
                {
                    int idx = scheduledData.FindIndex(x => x.OrderID == critpath[randomnums[0]].OrderID && x.OPID == critpath[randomnums[0]].OPID);
                    int idx2 = scheduledData.FindIndex(x => x.OrderID == critpath[randomnums[1]].OrderID && x.OPID == critpath[randomnums[1]].OPID);
                    var swap = Datas[idx];
                    var swap2 = Datas[idx2];
                    var orderList = new List<string>(scheduledData.Distinct(x => x.OrderID)
                                                                  .Select(x => x.OrderID)
                                                                  .ToList());
                    //製程互換
                    scheduledData[idx] = swap2.Clone() as Chromsome;
                    scheduledData[idx2] = swap.Clone() as Chromsome;

                    var duration1 = swap.EndTime - swap.StartTime;
                    var duration2 = swap2.EndTime - swap2.StartTime;

                    //更新互換後的機台和開始時間
                    scheduledData[idx].StartTime = swap.StartTime;
                    scheduledData[idx].WorkGroup = swap.WorkGroup;
                    scheduledData[idx].EndTime = swap.StartTime.Add(duration2);
                    scheduledData[idx2].WorkGroup = swap2.WorkGroup;
                    scheduledData[idx2].StartTime = swap2.StartTime;

                    scheduledData[idx2].EndTime = swap2.StartTime.Add(duration1);

                    var check = scheduledData.Distinct(x => x.WorkGroup)
                                             .Select(x => x.WorkGroup)
                                             .ToList();

                    //調整時間避免重疊
                    for (int k = 0; k < 2; k++)
                    {
                        foreach (var one_order in orderList)
                        {
                            //挑選同工單製程
                            var temp = scheduledData.Where(x => x.OrderID == one_order)
                                             .OrderBy(x => x.Range)
                                             .ToList();


                            #region 判斷是否為下班日or六日
                            //if (!m_IsWorkingDay(startTime))
                            //{
                            //    if (startTime > DateTime.Parse(startTime.ToShortDateString() + " 08:00"))
                            //    {
                            //        if (startTime.DayOfWeek == DayOfWeek.Saturday)
                            //            startTime = DateTime.Parse(startTime.AddDays(2).ToShortDateString() + " 08:00");
                            //        else if (startTime.DayOfWeek == DayOfWeek.Friday)
                            //            startTime = DateTime.Parse(startTime.AddDays(3).ToShortDateString() + " 08:00");
                            //        else
                            //            startTime = DateTime.Parse(startTime.AddDays(1).ToShortDateString() + " 08:00");
                            //    }
                            //    else
                            //        startTime = DateTime.Parse(startTime.ToShortDateString() + " 08:00");
                            //}
                            //else
                            //{
                            //    var s = startTime.DayOfWeek;
                            //}
                            #endregion

                            #region 有多道製程時
                            for (int i = 1; i < temp.Count; i++)
                            {
                                int indx = 0;
                                //調整同工單製程
                                if (DateTime.Compare(Convert.ToDateTime(temp[i - 1].EndTime), Convert.ToDateTime(temp[i].StartTime)) > 0)
                                {
                                    indx = scheduledData.FindIndex(x => x.OrderID == temp[i].OrderID && x.OPID == temp[i].OPID);
                                    scheduledData[indx].StartTime = temp[i - 1].EndTime;
                                    scheduledData[indx].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                                    temp[i].StartTime = temp[i - 1].EndTime;
                                    temp[i].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                                }
                                //調整同機台製程
                                if (scheduledData.Exists(x => temp[i].WorkGroup == x.WorkGroup))
                                {
                                    var sequence = scheduledData.Where(x => x.WorkGroup == temp[i].WorkGroup)
                                                         .OrderBy(x => x.StartTime)
                                                         .ToList();
                                    for (int j = 1; j < sequence.Count; j++)
                                    {
                                        if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                        {
                                            indx = scheduledData.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                            scheduledData[indx].StartTime = sequence[j - 1].EndTime;
                                            scheduledData[indx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                            sequence[j].StartTime = sequence[j - 1].EndTime;
                                            sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                                        }
                                    }
                                }
                            }
                            #endregion

                            #region 只有單一道製程時

                            ////調整同機台製程
                            //for (int i = 0; i < temp.Count; i++)
                            //{
                            //    if (scheduledData.Exists(x => temp[i].WorkGroup == x.WorkGroup))
                            //    {
                            //        var sequence = scheduledData.Where(x => x.WorkGroup == temp[i].WorkGroup)
                            //                             .OrderBy(x => x.StartTime)
                            //                             .ToList();
                            //        for (int j = 1; j < sequence.Count; j++)
                            //        {
                            //            if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                            //            {
                            //                int Idx = scheduledData.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                            //                scheduledData[Idx].StartTime = sequence[j - 1].EndTime;
                            //                scheduledData[Idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                            //                sequence[j].StartTime = sequence[j - 1].EndTime;
                            //                sequence[j].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
                            //            }
                            //        }
                            //    }
                            //}

                            #endregion
                        }
                    }
                }
                CountDelay(scheduledData);
            }

            void findLastTime(Chromsome temp, List<Chromsome> data)
            {
                var t1 = data.FindLast(x => x.WorkGroup == temp.WorkGroup && DateTime.Compare(x.EndTime, temp.StartTime) <= 0);

                DateTime OPET = data.FindLast(x => x.OrderID == temp.OrderID && x.OPID == temp.OPID).EndTime;

                if (t1 is null)
                {
                    temp.EndTime = temp.StartTime + temp.Duration;
                }
                else
                {
                    if (!(t1 is null) && DateTime.Compare(OPET, t1.EndTime) >= 0)
                    {
                        temp.StartTime = OPET;
                        temp.EndTime = OPET + temp.Duration;
                    }
                    else
                    {
                        if (DateTime.Compare(t1.StartTime, temp.StartTime) != 0)
                        {
                            temp.StartTime = t1.StartTime;
                            temp.EndTime = t1.StartTime + temp.Duration;
                        }
                        else
                        {
                            temp.EndTime = temp.EndTime = temp.StartTime + temp.Duration;
                        }
                    }
                }
            }
        }

        public List<Chromsome> Scheduled(List<LocalMachineSeq> firstSchedule)
        {
            var OutsourcingList = getOutsourcings();

            var result = new List<Chromsome>();
            int Idx = 0;
            DateTime getNow = DateTime.Now;
            DateTime PostST = getNow;
            DateTime PostET = getNow;
            var SortSchedule = firstSchedule.OrderBy(x => x.EachMachineSeq).ToList();//依據seq順序排每一台機台

            for (int i = 0; i < SortSchedule.Count; i++)
            {
                Idx = 0;
                PostST = getNow;
                PostET = getNow;

                if (result.Exists(x => x.WorkGroup == SortSchedule[i].WorkGroup) && OutsourcingList.Exists(x => x.remark == SortSchedule[i].WorkGroup))
                {
                    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "0")//該機台已有排程且非委外機台
                    {
                        Idx = result.FindLastIndex(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                        PostST = result[Idx].EndTime;
                    }
                }
                else
                {
                    //比較同機台最後一道製程&同工單最後一道製程結束時間
                    if (ReportedMachine.Keys.Contains(SortSchedule[i].WorkGroup) && ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                    {
                        PostST = ReportedMachine[SortSchedule[i].WorkGroup] >= ReportedOrder[SortSchedule[i].OrderID] ? ReportedMachine[SortSchedule[i].WorkGroup] : ReportedOrder[SortSchedule[i].OrderID];
                    }
                    else if (ReportedMachine.Count > 0 && ReportedMachine.Keys.Contains(SortSchedule[i].WorkGroup))
                    {
                        PostST = ReportedMachine[SortSchedule[i].WorkGroup];
                    }
                    else if (ReportedOrder.Count > 0)
                    {
                        if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                        {
                            PostST = ReportedOrder[SortSchedule[i].OrderID];
                        }
                    }
                }

                //補償休息時間
                //PostET = restTimecheck(PostST, ii.Duration);

                PostET = PostST + SortSchedule[i].Duration;

                result.Add(new Chromsome
                {
                    SeriesID = SortSchedule[i].SeriesID,
                    OrderID = SortSchedule[i].OrderID,
                    OPID = SortSchedule[i].OPID,
                    Range = SortSchedule[i].Range,
                    StartTime = PostST,
                    EndTime = PostET,
                    WorkGroup = SortSchedule[i].WorkGroup,
                    AssignDate = SortSchedule[i].PredictTime,
                    PartCount = SortSchedule[i].PartCount,
                    Maktx = SortSchedule[i].Maktx,
                    Duration = SortSchedule[i].Duration,
                    EachMachineSeq = SortSchedule[i].EachMachineSeq
                });
            }

            //篩選本次排程工單類別
            //var orderList = firstSchedule.OrderBy(x => x.EachMachineSeq).Select(x => x.OrderID)
            //                      .Distinct()
            //                      .ToList();
            var orderList = result.Distinct(x => x.OrderID)
                                  .Select(x => x.OrderID)
                                  .ToList();

            for (int k = 0; k < 2; k++)
            {
                foreach (var one_order in orderList)
                {
                    //挑選同工單製程
                    var temp = result.Where(x => x.OrderID == one_order)
                                     .OrderBy(x => x.Range)
                                     .ToList();

                    for (int i = 1; i < temp.Count; i++)
                    {
                        int idx;

                        //調整同工單製程
                        if (DateTime.Compare(Convert.ToDateTime(temp[i - 1].EndTime), Convert.ToDateTime(temp[i].StartTime)) > 0)
                        {
                            idx = result.FindIndex(x => x.OrderID == temp[i].OrderID && x.OPID == temp[i].OPID);
                            result[idx].StartTime = temp[i - 1].EndTime;
                            result[idx].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                            temp[i].StartTime = temp[i - 1].EndTime;
                            temp[i].EndTime = temp[i - 1].EndTime + temp[i].Duration;
                        }
                        //若非超音波清洗再調整同機台製程
                        if (OutsourcingList.Exists(x => x.remark == temp[i].WorkGroup))
                        {
                            if (OutsourcingList.Where(x => x.remark == temp[i].WorkGroup).First().isOutsource == "0")
                            {
                                //調整同機台製程
                                if (result.Exists(x => temp[i].WorkGroup == x.WorkGroup))
                                {
                                    var sequence = result.Where(x => x.WorkGroup == temp[i].WorkGroup)
                                                         .OrderBy(x => x.StartTime)
                                                         .ToList();
                                    for (int j = 1; j < sequence.Count; j++)
                                    {
                                        if (DateTime.Compare(sequence[j - 1].EndTime, sequence[j].StartTime) > 0)
                                        {
                                            idx = result.FindIndex(x => x.OrderID == sequence[j].OrderID && x.OPID == sequence[j].OPID);
                                            result[idx].StartTime = sequence[j - 1].EndTime;
                                            result[idx].EndTime = sequence[j - 1].EndTime + sequence[j].Duration;
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

            CountDelay(result);
            return result;
        }

        /// <summary>
        /// 找關鍵路徑
        /// </summary>
        /// <param name="Inputschedule"></param>
        /// <returns></returns>
        public List<Chromsome> FindCriticalPath(List<Chromsome> Inputschedule)
        {
            Inputschedule = Inputschedule.OrderBy(x => x.WorkGroup).ThenByDescending(x => x.StartTime).ToList();
            var makespan = Inputschedule.Max(x => x.EndTime);
            var begin = Inputschedule.Find(x => x.EndTime == makespan);
            string orderid;
            double opid;
            var result = new List<Chromsome>();
            result.Add(begin);
            while (true)
            {
                orderid = begin.OrderID;
                opid = Convert.ToDouble(begin.OPID);
                var sameod = Inputschedule.Find(x => x.OrderID == orderid && Convert.ToDouble(x.OPID) == opid - 1);
                var samewg = Inputschedule.Find(x => x.WorkGroup == begin.WorkGroup && x.EndTime <= begin.StartTime);
                if (Inputschedule.Exists(x => x.OrderID == orderid && Convert.ToDouble(x.OPID) == opid - 1) && Inputschedule.Exists(x => x.WorkGroup == begin.WorkGroup && x.EndTime <= begin.StartTime))
                {
                    if (sameod.EndTime > samewg.EndTime)
                    {
                        result.Add(sameod);
                        begin = sameod;
                    }
                    else
                    {
                        result.Add(samewg);
                        begin = samewg;
                    }
                }
                else if (Inputschedule.Exists(x => x.OrderID == orderid && Convert.ToDouble(x.OPID) == opid - 1))
                {
                    result.Add(sameod);
                    begin = sameod;
                }
                else if (Inputschedule.Exists(x => x.WorkGroup == begin.WorkGroup && x.EndTime <= begin.StartTime))
                {
                    result.Add(samewg);
                    begin = samewg;
                }
                else
                {
                    break;
                }
            }
            return result;
        }

        //取得外包機台的資料
        public List<MRP.Outsource> getOutsourcings()
        {
            string SqlStr = $@"SELECT a.*,b.Outsource
                              FROM {_ConnectStr.APSDB}.[dbo].Device as a
                              left join {_ConnectStr.APSDB}.[dbo].Outsourcing as b on a.ID=b.Id";
            List<MRP.Outsource> result = new List<MRP.Outsource>(); ;
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
                                result.Add(new MRP.Outsource
                                {
                                    ID = int.Parse(SqlData["ID"].ToString()),
                                    remark = SqlData["remark"].ToString(),
                                    isOutsource = SqlData["Outsource"].ToString(),
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        //補償休息時間
        private DateTime restTimecheck(DateTime PostST, TimeSpan Duration)
        {
            if (Duration > new TimeSpan(1, 00, 00, 00))
            {
                var days = Duration.TotalDays;
                TimeSpan resttime = new TimeSpan((int)(16 * days), 00, 00);
                Duration = Duration.Subtract(resttime);
            }
            const int hoursPerDay = 9;
            const int startHour = 8;
            // Don't start counting hours until start time is during working hours
            if (PostST.TimeOfDay.TotalHours > startHour + hoursPerDay)
                PostST = PostST.Date.AddDays(1).AddHours(startHour);
            if (PostST.TimeOfDay.TotalHours < startHour)
                PostST = PostST.Date.AddHours(startHour);
            if (PostST.DayOfWeek == DayOfWeek.Saturday)
                PostST.AddDays(2);
            else if (PostST.DayOfWeek == DayOfWeek.Sunday)
                PostST.AddDays(1);
            // Calculate how much working time already passed on the first day
            TimeSpan firstDayOffset = PostST.TimeOfDay.Subtract(TimeSpan.FromHours(startHour));
            // Calculate number of whole days to add
            var aaa = Duration.Add(firstDayOffset).TotalHours;
            int wholeDays = (int)(Duration.Add(firstDayOffset).TotalHours / hoursPerDay);
            // How many hours off the specified offset does this many whole days consume?
            TimeSpan wholeDaysHours = TimeSpan.FromHours(wholeDays * hoursPerDay);
            // Calculate the final time of day based on the number of whole days spanned and the specified offset
            TimeSpan remainder = Duration - wholeDaysHours;
            // How far into the week is the starting date?
            int weekOffset = ((int)(PostST.DayOfWeek + 7) - (int)DayOfWeek.Monday) % 7;
            // How many weekends are spanned?
            int weekends = (int)((wholeDays + weekOffset) / 5);
            // Calculate the final result using all the above calculated values
            return PostST.AddDays(wholeDays + weekends * 2).Add(remainder);
        }

        public void InspectJobOper(Dictionary<int, List<Chromsome>> crossoverResultList, ref Dictionary<int, List<Chromsome>> ChromosomeList, List<Evafitnessvalue> fitness_idx_value, ref int noImprovementCount)
        {
            bool HasImproved = false;
            //取得各機台是否為委外機台
            var OutsourcingList = getOutsourcings();
            for (int i = 0; i < crossoverResultList.Count; i++)
            {
                int total = ChromosomeList[i].Count;//正確工單製程數
                var results = new List<Tuple<string, double>>();
                for (int j = 0; j < crossoverResultList[i].Count; j++)
                {
                    results.Add(Tuple.Create(crossoverResultList[i][j].OrderID, crossoverResultList[i][j].OPID));
                }
                List<Tuple<string, double>> distinct_2 = results.Distinct().ToList();
                var distinct_1 = new List<Tuple<string, double, string, TimeSpan, DateTime, int, string>>();
                //把遺失的工單工序加回來
                if (distinct_2.Count != total)
                {
                    foreach (var item in ChromosomeList[i])
                    {
                        if (!distinct_2.Exists(x => x.Item1 == item.OrderID && x.Item2 == item.OPID))
                        {
                            distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, item.WorkGroup, item.Duration, item.AssignDate, item.Range, item.SeriesID));
                            continue;
                        }
                        var query = crossoverResultList[i].Find(x => x.OrderID == item.OrderID && x.OPID == item.OPID);
                        distinct_1.Add(Tuple.Create(item.OrderID, item.OPID, query.WorkGroup, query.Duration, query.AssignDate, query.Range, item.SeriesID));
                    }
                }
                else
                {
                    distinct_1 = crossoverResultList[i].Select(x => Tuple.Create(x.OrderID, x.OPID, x.WorkGroup, x.Duration, x.AssignDate, x.Range, x.SeriesID))
                                                       .ToList();
                }
                //重新給定機台排序
                List<LocalMachineSeq> MachineSeq = new List<LocalMachineSeq>();
                for (int machinenameseq = 0; machinenameseq < Devices.Count; machinenameseq++)
                {
                    int seq = 0;
                    //排序以OPID排，避免同工單後製程放在前製程前面
                    var ordersOnMachine = distinct_1.Where(x => x.Item3 == Devices[machinenameseq].Remark).OrderBy(x => x.Item2);
                    foreach (var item in ordersOnMachine)
                    {
                        if(OutsourcingList.Exists(x=>x.remark== Devices[machinenameseq].Remark))
                        {
                            if(OutsourcingList.Where(x => x.remark == Devices[machinenameseq].Remark).First().isOutsource == "1")
                            {
                                MachineSeq.Add(new LocalMachineSeq
                                {
                                    OrderID = item.Item1,
                                    OPID = item.Item2,
                                    WorkGroup = item.Item3,
                                    Duration = item.Item4,
                                    PredictTime = item.Item5,
                                    PartCount = item.Item6,
                                    Range = item.Item6,
                                    EachMachineSeq = 0,
                                });
                            }
                            else
                            {
                                MachineSeq.Add(new LocalMachineSeq
                                {
                                    SeriesID = item.Item7,
                                    OrderID = item.Item1,
                                    OPID = item.Item2,
                                    WorkGroup = item.Item3,
                                    Duration = item.Item4,
                                    PredictTime = item.Item5,
                                    PartCount = item.Item6,
                                    Range = item.Item6,
                                    EachMachineSeq = seq,
                                });
                                seq++;
                            }
                        }
                    }
                    if (MachineSeq.Count == distinct_1.Count)
                    {
                        break;
                    }
                }
                //重新排程
                var tempOrder = Scheduled(MachineSeq);
                ////突變mutation
                //Random rand = new Random(Guid.NewGuid().GetHashCode());
                //if(rand.NextDouble()<0.05)
                //{
                //    Mutation(tempOrder);
                //}

                // 多一個比較sumdelay
                int sum = tempOrder.Sum(x => x.Delay);
                if (fitness_idx_value.Exists(x => x.Fitness > sum)) //判斷突變之後是否有更好的解
                {
                    int index = fitness_idx_value.FindIndex(x => x.Fitness > sum);
                    ChromosomeList.Remove(fitness_idx_value[index].Idx);
                    ChromosomeList.Add(fitness_idx_value[index].Idx, tempOrder.Select(x => (Chromsome)x.Clone())
                                                                              .ToList());
                    Debug.WriteLine($"delay is {fitness_idx_value[0].Fitness}");
                    HasImproved = true;

                }
            }
            if(HasImproved==false)
            {
                noImprovementCount += 1;
            }
        }

        public void CountDelay(List<Chromsome> Tep)
        {
            TimeSpan temp;
            int itemDelay;
            foreach (var item in Tep)
            {
                try
                {
                    temp = item.AssignDate - item.EndTime;
                    itemDelay = (temp.TotalDays > 0) ? 0 : Math.Abs(temp.Days);
                    item.Delay = itemDelay;
                }
                catch
                {
                    continue;
                }
            }
        }
    }
}

internal interface IMAthModel
{

}