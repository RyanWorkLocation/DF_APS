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

                //#region 原本的查找方式
                ////目前排程結果已有同機台製程
                //if (result.Exists(x => x.WorkGroup == SortSchedule[i].WorkGroup) && OutsourcingList.Exists(x=>x.remark== SortSchedule[i].WorkGroup))
                //{
                //    //是同步加工機台=>只需考量前製程完成時間
                //    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "1")//該機台為委外機台
                //    {
                //        //目前排程是否有同工單
                //        if(result.Exists(x=>x.OrderID== SortSchedule[i].OrderID))
                //        {
                //            int idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                //            if (idx >= 0)
                //                PostST = result[idx].EndTime;
                //            else
                //                PostST = preserve_Now;
                //            //Idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                //            //PostST = result[Idx].EndTime;
                //        }
                //        //如果目前排程沒有，則確認原排程是否有同工單製程
                //        else if(ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                //        {
                            
                //            //有的話，比較當前時間
                //            PostST = preserve_Now > ReportedOrder[SortSchedule[i].OrderID] ? DateTime.Now : ReportedOrder[SortSchedule[i].OrderID];
                //        }
                //        else
                //        {
                //            //沒有的話，直接使用當前時間
                //            PostST = preserve_Now;
                //        }
                        
                //    }
                //    //非同步加工機台=>考量同機台前製程&同工單前製程
                //    else
                //    {
                //        var Idx_order = 0;
                //        var Idx_machine = 0;
                //        //目前排程是否有同工單
                //        if (result.Exists(x => x.OrderID == SortSchedule[i].OrderID))
                //        {
                //            //比較同機台和同工單最早可開工時間
                //            // 找到最後一筆相同 OrderID 的索引
                //            Idx_order = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                //            Idx_machine = result.FindLastIndex(x => x.WorkGroup == SortSchedule[i].WorkGroup);

                //            // 安全地取得 EndTime，若找不到則預設為 DateTime.MinValue
                //            DateTime EndTime_order = (Idx_order != -1) ? result[Idx_order].EndTime : preserve_Now;
                //            DateTime EndTime_machine = (Idx_machine != -1) ? result[Idx_machine].EndTime : preserve_Now;

                //            // 取較晚的時間
                //            PostST = (EndTime_order >= EndTime_machine) ? EndTime_order : EndTime_machine;

                //            //Idx_order = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                //            //Idx_machine = result.FindLastIndex(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                //            //PostST = result[Idx_order].EndTime >= result[Idx_machine].EndTime ? result[Idx_order].EndTime : result[Idx_machine].EndTime;
                //        }
                //        //如果目前排程沒有，則確認原排程是否有同工單製程
                //        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                //        {
                //            //有的話，比較當前時間&同機台時間&同工單時間
                //            DateTime time1 = ReportedOrder[SortSchedule[i].OrderID];
                //            var foundMachine = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                //            DateTime time2 = foundMachine != null ? foundMachine.EndTime : preserve_Now;
                //            //DateTime time2 = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup).EndTime;

                //            PostST = time1 >= time2 ? time1 : time2;
                //        }
                //        else
                //        {
                //            //沒有的話，直接使用同機台時間
                //            var foundMachine = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup);
                //            PostST = foundMachine != null ? foundMachine.EndTime : preserve_Now;
                //            //PostST = result.Find(x => x.WorkGroup == SortSchedule[i].WorkGroup).EndTime;
                //        }
                //    }
                //}
                ////目前排程結果沒有同機台製程
                //else
                //{
                //    //是同步加工機台 => 只需比較前製程完成時間和當前時間
                //    if (OutsourcingList.Where(x => x.remark == SortSchedule[i].WorkGroup).First().isOutsource == "1")//該機台為委外機台
                //    {
                //        //如果目前排程有同工單，則直接使用前製程時間
                //        if (result.Exists(x => x.OrderID == SortSchedule[i].OrderID))
                //        {
                //            int idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                //            if (idx >= 0)
                //                PostST = result[idx].EndTime;
                //            else
                //                PostST = preserve_Now;
                //            //Idx = result.FindLastIndex(x => x.OrderID == SortSchedule[i].OrderID);
                //            //PostST = result[Idx].EndTime;
                //        }
                //        //如果目前排程沒有，則確認原排程是否有同工單製程
                //        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                //        {
                //            //有的話，比較當前時間
                //            PostST = preserve_Now > ReportedOrder[SortSchedule[i].OrderID] ? preserve_Now : ReportedOrder[SortSchedule[i].OrderID];
                //        }
                //        else
                //        {
                //            //沒有的話，直接使用當前時間
                //            PostST = preserve_Now;
                //        }
                //    }
                //    //非同步加工機台=>考量同機台前製程&同工單前製程
                //    else
                //    {
                //        //如果目前排程有同工單
                //        if (result.Exists(x => x.OrderID == SortSchedule[i].OrderID))
                //        {
                //            //如果原排程有同機台製程
                //            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                //            {
                //                //比較同機台和同工單最早可開工時間
                //                DateTime time1 = result.Find(x => x.OrderID == SortSchedule[i].OrderID).EndTime;
                //                DateTime time2 = ReportedMachine[SortSchedule[i].OrderID];
                //                PostST = time1 >= time2 ? time1 : time2;
                //            }
                //            //如果原排程沒有同機台製程
                //            else
                //            {
                //                //直接使用前製程時間
                //                PostST = result.Find(x => x.OrderID == SortSchedule[i].OrderID).EndTime;
                //            } 
                //        }
                //        //如果目前排程沒有同工單，則確認原排程是否有同工單製程
                //        else if (ReportedOrder.Keys.Contains(SortSchedule[i].OrderID))
                //        {
                //            //如果原排程有同機台製程
                //            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                //            {
                //                //比較同機台、同工單、當前時間
                //                DateTime time1 = preserve_Now;
                //                DateTime time2 = ReportedOrder[SortSchedule[i].OrderID];
                //                DateTime time3 = ReportedMachine[SortSchedule[i].OrderID];
                //                if (time1 >= time2 && time1 >= time3)
                //                    PostST = time1;
                //                else if (time2 >= time1 && time2 >= time3)
                //                    PostST = time2;
                //                else
                //                    PostST = time3;
                //            }
                //            //如果原排程沒有同機台製程，比較原排程同工單、當前時間
                //            else
                //            {
                //                DateTime time1 = preserve_Now;
                //                DateTime time2 = ReportedOrder[SortSchedule[i].OrderID];
                //                PostST = time1 >= time2 ? time1 : time2;
                //            }
                //        }
                //        else
                //        {
                //            //如果原排程有同機台，比較原排程同機台&當前時間
                //            if (ReportedMachine.Keys.Contains(SortSchedule[i].OrderID))
                //            {
                //                DateTime time1 = preserve_Now;
                //                DateTime time2 = ReportedMachine[SortSchedule[i].OrderID];
                //                PostST = time1 >= time2 ? time1 : time2;
                //            }
                //            //如果原排程有沒同機台，直接用當前時間
                //            else
                //            {
                //                PostST = preserve_Now;
                //            }
                //        }
                //    }
                //}
                //#endregion


                #region 新的快取查找方式
                // 建立快取，讓查找更快，並直接儲存對應的物件
                // 每次都根據cache找出最後一道製程時間
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

            //CountDelay(result);
            Delay_and_waiting(result);
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
        //public void EvaluationFitness(ref Dictionary<int, List<Chromsome>> ChromosomeList, ref int noImprovementCount)
        //{
        //    var fitness_idx_value = new List<Evafitnessvalue>();
        //    var localOptChromosomeList = new Dictionary<int, List<Chromsome>>(); // 使用局部变量存储结果
        //    var localChromosomeList = new Dictionary<int, List<Chromsome>>(ChromosomeList); // 创建局部变量

        //    // 使用并行计算计算适应度
        //    Parallel.ForEach(localChromosomeList, kvp =>
        //    {
        //        int sumDelay = kvp.Value.Sum(x => x.Delay);
        //        lock (fitness_idx_value)
        //        {
        //            fitness_idx_value.Add(new Evafitnessvalue(kvp.Key, sumDelay));
        //        }
        //    });

        //    // 计算适应度后排序，由小到大
        //    fitness_idx_value.Sort((x, y) => x.Fitness.CompareTo(y.Fitness));

        //    // 挑出前50%的染色体解答
        //    int chromosomeCount = Chromvalue / 2;

        //    Parallel.For(0, chromosomeCount, i =>
        //    {
        //        var idx = fitness_idx_value[i].Idx;
        //        var orderedChromosomes = localChromosomeList[idx]
        //            .OrderBy(x => x.WorkGroup)
        //            .ThenBy(x => x.StartTime)
        //            .Select(x => x.Clone() as Chromsome)
        //            .ToList();

        //        lock (localOptChromosomeList)
        //        {
        //            localOptChromosomeList[i] = orderedChromosomes;
        //        }
        //    });

        //    // 将局部变量更新到 ChromosomeList
        //    foreach (var item in localOptChromosomeList)
        //    {
        //        ChromosomeList[item.Key] = item.Value;
        //    }

        //    var random = new Random(Guid.NewGuid().GetHashCode());
        //    var crossoverResultList = new ConcurrentDictionary<int, List<Chromsome>>();

        //    var crossoverList = new List<List<Chromsome>>();
        //    var crossoverTemp = new List<List<Chromsome>>();

        //    // opt_ChromosomeList 是前50%的母体资料 选两个来做交换
        //    for (int i = 0; i < chromosomeCount; i++)
        //    {
        //        int randomNum = random.Next(0, chromosomeCount);
        //        crossoverList.Add(localOptChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
        //        crossoverTemp.Add(localOptChromosomeList[randomNum].Select(x => x.Clone() as Chromsome).ToList());
        //    }

        //    Parallel.For(0, chromosomeCount, childItem =>
        //    {
        //        // crossover
        //        int cutLine = random.Next(1, crossoverList[0].Count);

        //        if (childItem < chromosomeCount - 1)
        //        {
        //            if (cutLine < crossoverList[childItem + 1].Count)
        //            {
        //                int length1 = crossoverList[childItem + 1].Count - cutLine;
        //                if (length1 > 0)
        //                {
        //                    var swapData1 = crossoverList[childItem + 1].GetRange(cutLine, length1);

        //                    lock (crossoverTemp)
        //                    {
        //                        crossoverTemp[childItem].RemoveRange(cutLine, length1);
        //                        crossoverTemp[childItem].AddRange(swapData1);
        //                    }

        //                    int length2 = crossoverList[childItem].Count - cutLine;
        //                    if (length2 > 0)
        //                    {
        //                        var swapData2 = crossoverList[childItem].GetRange(cutLine, length2);

        //                        lock (crossoverTemp)
        //                        {
        //                            crossoverTemp[childItem + 1].RemoveRange(cutLine, length2);
        //                            crossoverTemp[childItem + 1].AddRange(swapData2);
        //                        }
        //                    }
        //                }
        //            }

        //            crossoverResultList[2 * childItem] = crossoverTemp[childItem];
        //            crossoverResultList[2 * childItem + 1] = crossoverTemp[childItem + 1];
        //        }
        //        else
        //        {
        //            if (cutLine < crossoverList[0].Count)
        //            {
        //                int length1 = crossoverList[0].Count - cutLine;
        //                if (length1 > 0)
        //                {
        //                    var swapData1 = crossoverList[0].GetRange(cutLine, length1);

        //                    lock (crossoverTemp)
        //                    {
        //                        crossoverTemp[childItem].RemoveRange(cutLine, length1);
        //                        crossoverTemp[childItem].AddRange(swapData1);
        //                    }

        //                    int length2 = crossoverList[childItem].Count - cutLine;
        //                    if (length2 > 0)
        //                    {
        //                        var swapData2 = crossoverList[childItem].GetRange(cutLine, length2);

        //                        lock (crossoverTemp)
        //                        {
        //                            crossoverTemp[0].RemoveRange(cutLine, length2);
        //                            crossoverTemp[0].AddRange(swapData2);
        //                        }
        //                    }
        //                }
        //            }

        //            crossoverResultList[2 * childItem] = crossoverTemp[childItem];
        //            crossoverResultList[2 * childItem + 1] = crossoverTemp[0];
        //        }
        //    });

        //    InspectJobOper(crossoverResultList, ref ChromosomeList, fitness_idx_value.GetRange(0, crossoverList.Count), ref noImprovementCount);
        //}


        public double Calculate_Combined_Fitness(List<Chromsome> soulution)
        {
            int sumDelay = 0;
            double sumWaiting = 0;
            int delayedProcessCount = 0; // 記錄有延遲的製程數
            int waitingProcessCount = 0; // 記錄有等待的製程數
            int done_too_early_orderCount = 0; // 記錄有延遲的製程數
            int lastProcessDelay = 0; // 記錄最後一道製程的早交天數

            foreach (var process in soulution)
            {
                if (process.Delay > 0) // 只統計有延遲的製程
                {
                    sumDelay += process.Delay;
                    delayedProcessCount++;
                }
                if (process.Waiting > 0)
                {
                    sumWaiting += process.Waiting;
                    waitingProcessCount++;
                }
            }

            // 找到最後一道製程
            var lastProcessList = soulution
                                .GroupBy(x => x.OrderID)
                                .Select(g => g.OrderByDescending(p => p.EndTime).First())
                                .ToList();  // 轉為 List

            // 取得第一個 OrderID 的 EndTime
            if (lastProcessList != null)
            {
                foreach (var order in lastProcessList)
                {
                    if (order.AssignDate > order.EndTime)
                    {
                        //避免交貨日期太早完成，壓縮其他工單時間(限制需小於10天，大於5天)
                        var daydiff = (order.AssignDate - order.EndTime).Days;
                        if (daydiff > 10)
                        {
                            done_too_early_orderCount += 1;
                            lastProcessDelay += (daydiff - 10);
                        }
                        else if (daydiff < 3)
                        {
                            done_too_early_orderCount += 1;
                            lastProcessDelay += (3 - daydiff);
                        }
                    }
                }

            }
            ////計算平均過度交貨冗餘或緊縮
            //double avg_overdelivery_redundancy = done_too_early_orderCount > 0 ? (double)lastProcessDelay / done_too_early_orderCount : 0;

            // 計算平均延遲
            double avgDelay = delayedProcessCount > 0 ? (double)sumDelay / delayedProcessCount : 0;

            // 計算平均等待時間
            double avgWaiting = waitingProcessCount > 0 ? (double)sumWaiting / waitingProcessCount : 0;

            // 給予權重，優先看交貨日期準的
            double finalFitness = 0.7*avgDelay + 0.3*avgWaiting;
            return finalFitness;
        }

        


        //2025.02.20 修改測試
        public void EvaluationFitness(ref Dictionary<int, List<Chromsome>> ChromosomeList, ref int noImprovementCount)
        {
            var fitness_idx_value = new List<Evafitnessvalue>();
            var localOptChromosomeList = new Dictionary<int, List<Chromsome>>(); // 使用局部變量存儲結果
            var localChromosomeList = new Dictionary<int, List<Chromsome>>(ChromosomeList); // 創建局部變量

            // 使用並行計算計算適應度
            Parallel.ForEach(localChromosomeList, kvp =>
            {

                //int sumDelay = 0;
                //double sumWaiting = 0;
                //int delayedProcessCount = 0; // 記錄有延遲的製程數
                //int waitingProcessCount = 0; // 記錄有延遲的製程數
                //int done_too_early_orderCount = 0; // 記錄有延遲的製程數
                //int lastProcessDelay = 0; // 記錄最後一道製程的延遲天數

                //foreach (var process in kvp.Value)
                //{
                //    if (process.Delay > 0) // 只統計有延遲的製程
                //    {
                //        sumDelay += process.Delay;
                //        delayedProcessCount++;
                //    }
                //    if(process.Waiting > 0)
                //    {
                //        sumWaiting += process.Waiting;
                //        waitingProcessCount++;
                //    }
                //}

                //// 找到最後一道製程
                //var lastProcessList = kvp.Value
                //                    .GroupBy(x => x.OrderID)
                //                    .Select(g => g.OrderByDescending(p => p.EndTime).First())
                //                    .ToList();  // 轉為 List

                //// 取得第一個 OrderID 的 EndTime
                //if (lastProcessList != null)
                //{
                //    foreach (var order in lastProcessList)
                //    {
                //        if (order.AssignDate > order.EndTime)
                //        {
                //            //避免交貨日期太早完成，壓縮其他工單時間(限制需小於10天，大於5天)
                //            var daydiff = (order.AssignDate - order.EndTime).Days;
                //            if (daydiff > 10)
                //            {
                //                done_too_early_orderCount += 1;
                //                lastProcessDelay += (daydiff - 10);
                //            } 
                //            else if (daydiff < 5)
                //            {
                //                done_too_early_orderCount += 1;
                //                lastProcessDelay += (5 - daydiff);
                //            }
                //        }
                //    }

                //}
                ////計算平均過度交貨冗餘或緊縮
                //double avg_overdelivery_redundancy = done_too_early_orderCount > 0 ? (double) lastProcessDelay / done_too_early_orderCount :0 ;

                //// 計算平均延遲
                //double avgDelay = delayedProcessCount > 0 ? (double)sumDelay / delayedProcessCount : 0;

                //// 計算平均等待時間
                //double avgWaiting = waitingProcessCount > 0 ? (double)sumWaiting / waitingProcessCount : 0;


                //double finalFitness = avgDelay + avgWaiting + avg_overdelivery_redundancy;

                double finalFitness = Calculate_Combined_Fitness(kvp.Value);

                lock (fitness_idx_value)
                {
                    fitness_idx_value.Add(new Evafitnessvalue(kvp.Key, finalFitness));
                }
            });

            // 計算適應度後排序，由小到大
            fitness_idx_value.Sort((x, y) => x.Fitness.CompareTo(y.Fitness));

            // 挑出前50%的染色體解答
            int chromosomeCount = Chromvalue / 2;

            Parallel.For(0, Chromvalue, i =>
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

            // 將局部變量更新到 ChromosomeList
            foreach (var item in localOptChromosomeList)
            {
                ChromosomeList[item.Key] = item.Value;
            }

            // 計算當前生成的自適應突變率
            double mutationRate = CalculateAdaptiveMutationRate(noImprovementCount, 0.05);

            var random = new Random(Guid.NewGuid().GetHashCode());
            var crossoverResultList = new ConcurrentDictionary<int, List<Chromsome>>();

            // 選擇父本進行交叉
            // 錦標賽選擇方式挑出一半的染色體
            var parentIndices = TournamentSelection(fitness_idx_value, chromosomeCount, 3);
            var parentList = new List<List<Chromsome>>();

            // 獲取選中的父本
            foreach (int idx in parentIndices)
            {
                parentList.Add(localOptChromosomeList[idx].Select(x => x.Clone() as Chromsome).ToList());
            }

            // 執行交叉和突變操作
            Parallel.For(0, chromosomeCount / 2, i =>
            {
                // 使用基於優先權的交叉運算子
                var offspring = PriorityBasedCrossover(
                    parentList[i * 2],
                    parentList[i * 2 + 1]
                );

                // 對每個後代應用突變
                SwapMutation(offspring[0], mutationRate);
                SwapMutation(offspring[1], mutationRate);

                // 保存後代
                crossoverResultList[i * 2] = offspring[0];
                crossoverResultList[i * 2 + 1] = offspring[1];
            });

            InspectJobOper(crossoverResultList, ref ChromosomeList, fitness_idx_value.GetRange(0, crossoverResultList.Count), ref noImprovementCount);
        }

        /// <summary>
        /// 使用競爭選擇法從個體群中選取指定數量的個體
        /// </summary>
        /// <param name="fitness">包含每個個體適應度值的列表</param>
        /// <param name="selectionCount">要選擇的個體數量</param>
        /// <param name="tournamentSize">每次競爭中參與比較的個體數量</param>
        /// <returns>被選中個體的索引列表</returns>
        private List<int> TournamentSelection(List<Evafitnessvalue> fitness, int selectionCount, int tournamentSize)
        {
            // 初始化存儲選中個體索引的列表
            var selected = new List<int>();

            // 初始化可選個體的索引範圍(0到fitness.Count-1)
            var availableIndices = Enumerable.Range(0, fitness.Count).ToList();

            // 創建隨機數生成器，使用GUID確保隨機性
            Random random = new Random(Guid.NewGuid().GetHashCode());

            // 循環選擇指定數量的個體，或直到沒有可用個體
            for (int i = 0; i < selectionCount && availableIndices.Count > 0; i++)
            {
                // 用於記錄本次競爭中最佳個體
                int bestIdx = -1;
                double bestFitness = double.MaxValue; // 假設適應度值越小越好
                List<int> tournamentIndices = new List<int>(); // 記錄參與本次競爭的個體索引

                // 從可用索引中隨機選取tournamentSize個個體進行比較
                // 如果可用個體數少於tournamentSize，則使用所有可用個體
                for (int j = 0; j < Math.Min(tournamentSize, availableIndices.Count); j++)
                {
                    // 從剩餘可用索引中隨機選一個位置
                    int randomPos = random.Next(0, availableIndices.Count);
                    int idx = availableIndices[randomPos];
                    tournamentIndices.Add(idx);

                    // 如果該個體的適應度更好(此處為值越小越好)，則更新最佳個體
                    if (fitness[idx].Fitness < bestFitness)
                    {
                        bestFitness = fitness[idx].Fitness;
                        bestIdx = idx;
                    }
                }

                // 將最佳個體的原始索引加入選中列表
                selected.Add(fitness[bestIdx].Idx);

                // 從可用索引中移除已選擇的個體，避免重複選擇
                availableIndices.Remove(bestIdx);
            }

            // 返回選中的個體索引列表
            return selected;
        }

        /// <summary>
        /// 基於優先權的交叉運算子
        /// </summary>
        /// <param name="parent1"></param>
        /// <param name="parent2"></param>
        /// <returns></returns>
        private List<List<Chromsome>> PriorityBasedCrossover(List<Chromsome> parent1, List<Chromsome> parent2)
        {
            Random random = new Random(Guid.NewGuid().GetHashCode());
            double alpha = random.NextDouble() * 0.4 + 0.3; // 混合權重在0.3到0.7之間

            // 1. 提取每個工序的優先權
            var priorities1 = new Dictionary<string, double>();
            var priorities2 = new Dictionary<string, double>();

            // 設定延遲和等待時間的權重
            double delayWeight = 1.5;  // 延遲權重較高
            double waitingWeight = 1.0;  // 等待時間權重

            // 用工序在染色體中的位置作為優先權
            for (int i = 0; i < parent1.Count; i++)
            {
                string key = parent1[i].OrderID + "-" + parent1[i].Range;


                // 優先度計算：延遲天數*權重 + 等待時間*權重
                // 值越小表示優先度越高（即需要更早處理）
                double priorityValue = parent1[i].Delay * delayWeight +
                                     (parent1[i].Waiting > 0 ? parent1[i].Waiting * waitingWeight : 0);
                //TimeSpan timeDifference = parent1[i].AssignDate - parent1[i].EndTime;
                //int daysDifference = timeDifference.Days;


                //if (daysDifference >= 0)
                //{
                //    // 未延遲的情況，給予最高優先級
                //    if (daysDifference > 10)
                //    {
                //        // 過早完成（但仍未延遲）
                //        priorityValue = 100 + (daysDifference - 10); // 基礎分數100，加上過早天數的懲罰
                //    }
                //    else if (daysDifference < 5)
                //    {
                //        // 接近截止日（但仍未延遲）
                //        priorityValue = 105 + (5 - daysDifference); // 基礎分數100，加上接近截止日的懲罰
                //    }
                //    else
                //    {
                //        // 理想範圍（5-10天內完成）
                //        priorityValue = 80; // 最理想情況，純基礎分數
                //    }
                //}
                //else
                //{
                //    // 已延遲的情況，優先度較低
                //    priorityValue = 200 + Math.Abs(daysDifference); // 基礎分數200（低於未延遲），加上延遲天數
                //}

                //if (daysDifference < 0)
                //{
                //    priorityValue += (double)Math.Abs(daysDifference);
                //}

                // 加入小的隨機擾動，避免陷入局部最優解
                double randomFactor = 0.98 + random.NextDouble() * 0.04;  // 0.98 到 1.02 之間
                priorityValue *= randomFactor;

                priorities1[key] = priorityValue;
            }

            for (int i = 0; i < parent2.Count; i++)
            {
                string key = parent2[i].OrderID + "-" + parent2[i].Range;
                // 使用與parent1相同的優先度計算方式
                double priorityValue = parent2[i].Delay * delayWeight +
                                     (parent2[i].Waiting > 0 ? parent2[i].Waiting * waitingWeight : 0);
                
                priorityValue = parent2[i].Delay + (parent2[i].Waiting > 0 ? parent2[i].Waiting : 0);
                //TimeSpan timeDifference = parent1[i].AssignDate - parent1[i].EndTime;
                //int daysDifference = timeDifference.Days;


                //if (daysDifference >= 0)
                //{
                //    // 未延遲的情況，給予最高優先級
                //    if (daysDifference > 10)
                //    {
                //        // 過早完成（但仍未延遲）
                //        priorityValue = 100 + (daysDifference - 10); // 基礎分數100，加上過早天數的懲罰
                //    }
                //    else if (daysDifference < 5)
                //    {
                //        // 接近截止日（但仍未延遲）
                //        priorityValue = 105 + (5 - daysDifference); // 基礎分數100，加上接近截止日的懲罰
                //    }
                //    else
                //    {
                //        // 理想範圍（5-10天內完成）
                //        priorityValue = 80; // 最理想情況，純基礎分數
                //    }
                //}
                //else
                //{
                //    // 已延遲的情況，優先度較低
                //    priorityValue = 200 + Math.Abs(daysDifference); // 基礎分數200（低於未延遲），加上延遲天數
                //}

                //if (daysDifference < 0)
                //{
                //    priorityValue += (double)Math.Abs(daysDifference);
                //}

                // 加入小的隨機擾動
                double randomFactor = 0.98 + random.NextDouble() * 0.04;
                priorityValue *= randomFactor;

                priorities2[key] = priorityValue;
            }

            // 2. 混合優先權生成兩個子代
            var childPriorities1 = new Dictionary<string, double>();
            var childPriorities2 = new Dictionary<string, double>();

            foreach (var key in priorities1.Keys)
            {
                // 確保兩個父本都有該工序
                if (priorities2.ContainsKey(key))
                {
                    //根據每個製程在父代中優先權計算子代的優先權
                    childPriorities1[key] = alpha * priorities1[key] + (1 - alpha) * priorities2[key];
                    childPriorities2[key] = (1 - alpha) * priorities1[key] + alpha * priorities2[key];
                }
                else
                {
                    childPriorities1[key] = priorities1[key];
                    childPriorities2[key] = priorities1[key];
                }
            }

            // 3. 根據混合後的優先權重新排序
            // 假設 Value 越大優先度越高
            var child1 = childPriorities1.OrderBy(x => x.Value)
                             .Select(x => {
                                 var parts = x.Key.Split('-');
                                 var orderId = parts[0];
                                 var range = parts[1];

                                 // 從 priorities1 和 priorities2 中獲取優先度
                                 double priority1 = priorities1.TryGetValue(x.Key, out double p1) ? p1 : double.MaxValue;
                                 double priority2 = priorities2.TryGetValue(x.Key, out double p2) ? p2 : double.MaxValue;

                                 // 優先度較高（Value 較小）的那方決定來源
                                 if (priority1>priority2)
                                 {
                                     return parent1.FirstOrDefault(j =>
                                         j.OrderID == orderId && j.Range.ToString() == range);
                                 }
                                 else if(priority2>priority1)
                                 {
                                     return parent2.FirstOrDefault(j =>
                                         j.OrderID == orderId && j.Range.ToString() == range);
                                 }
                                 else
                                 {
                                     // 優先度相同時隨機挑選
                                     return random.Next(2) == 0
                                         ? parent1.FirstOrDefault(j => j.OrderID == orderId && j.Range.ToString() == range)
                                         : parent2.FirstOrDefault(j => j.OrderID == orderId && j.Range.ToString() == range);
                                 }
                             })
                             .Where(j => j != null)
                             .Select(j => j.Clone() as Chromsome)
                             .ToList();

            var child2 = childPriorities2.OrderBy(x => x.Value)
                                         .Select(x => {
                                             var parts = x.Key.Split('-');
                                             var orderId = parts[0];
                                             var range = parts[1];

                                             // 從 priorities1 和 priorities2 中獲取優先度
                                             double priority1 = priorities1.TryGetValue(x.Key, out double p1) ? p1 : double.MaxValue;
                                             double priority2 = priorities2.TryGetValue(x.Key, out double p2) ? p2 : double.MaxValue;

                                 // 優先度較高（Value 較大）的那方決定來源
                                             if (priority2 > priority1)
                                             {
                                                 return parent2.FirstOrDefault(j =>
                                                     j.OrderID == orderId && j.Range.ToString() == range);
                                             }
                                             else if (priority1 > priority2)
                                             {
                                                 return parent1.FirstOrDefault(j =>
                                                     j.OrderID == orderId && j.Range.ToString() == range);
                                             }
                                             else
                                             {
                                                 // 優先度相同時隨機挑選
                                                 return random.Next(2) == 0
                                                     ? parent1.FirstOrDefault(j => j.OrderID == orderId && j.Range.ToString() == range)
                                                     : parent2.FirstOrDefault(j => j.OrderID == orderId && j.Range.ToString() == range);
                                             }
                                         })
                                         .Where(j => j != null)
                                         .Select(j => j.Clone() as Chromsome)
                                         .ToList();
            

            // 4. 確保每個子代的工序數量與父本相同
            EnsureJobCompleteness(child1, parent1);
            EnsureJobCompleteness(child2, parent2);

            // 5. 重新計算eachmachineseq
            var new_schedule_solution1 = AssignSequence(child1, new Dictionary<string, int>());
            var new_schedule_solution2 = AssignSequence(child2, new Dictionary<string, int>());


            // 6. 重新計算時間表
            //UpdateScheduleTimes(child1);
            //UpdateScheduleTimes(child2);
            var newchild1 = Scheduled(new_schedule_solution1);
            var newchild2 = Scheduled(new_schedule_solution2);

            return new List<List<Chromsome>> { newchild1, newchild2 };
        }

        //交配後分配機台順序
        private List<LocalMachineSeq> AssignSequence(List<Chromsome> child, Dictionary<string, int> eachmaclist)
        {
            List<LocalMachineSeq> new_seq = new List<LocalMachineSeq>();
            foreach (var item in child)
            {
                // 如果 WorkGroup 不存在，初始化為 0，否則遞增
                if (!eachmaclist.TryGetValue(item.WorkGroup, out int seq))
                {
                    eachmaclist[item.WorkGroup] = 0;
                }
                else
                {
                    eachmaclist[item.WorkGroup] = seq + 1;
                }

                // 分配 EachMachineSeq
                new_seq.Add(new LocalMachineSeq
                {
                    SeriesID = item.SeriesID,
                    OrderID = item.OrderID,
                    OPID = item.OPID,
                    Range = item.Range,
                    Duration = item.Duration,
                    PredictTime = item.AssignDate,
                    Maktx = item.Maktx,
                    PartCount = item.PartCount,
                    WorkGroup = item.WorkGroup,
                    EachMachineSeq = eachmaclist[item.WorkGroup],
                });

                
            }
            return new_seq;
        }

        // 確保子代包含所有必要的工序
        private void EnsureJobCompleteness(List<Chromsome> child, List<Chromsome> parent)
        {
            // 檢查是否所有工序都存在
            var childJobs = child.Select(j => j.OrderID + "-" + j.Range).ToHashSet();
            var parentJobs = parent.Select(j => j.OrderID + "-" + j.Range).ToHashSet();

            foreach (var job in parent)
            {
                string jobKey = job.OrderID + "-" + job.Range;
                if (!childJobs.Contains(jobKey))
                {
                    child.Add(job.Clone() as Chromsome);
                }
            }

            // 確保沒有重複工序
            child = child.GroupBy(j => j.OrderID + "-" + j.Range)
                        .Select(g => g.First())
                        .ToList();
        }
        // 交換突變算法
        private void SwapMutation(List<Chromsome> chromosome, double mutationRate)
        {
            Random random = new Random(Guid.NewGuid().GetHashCode());
            
            // 按機台分組
            var machineGroups = chromosome.GroupBy(x => x.WorkGroup).ToDictionary(g => g.Key, g => g.ToList());

            var schedule_after_mutation = new List<LocalMachineSeq>();

            foreach(var item in chromosome)
            {
                schedule_after_mutation.Add(new LocalMachineSeq
                {
                    SeriesID = item.SeriesID,
                    OrderID = item.OrderID,
                    OPID = item.OPID,
                    Range = item.Range,
                    Duration = item.Duration,
                    PredictTime = item.AssignDate,
                    Maktx = item.Maktx,
                    PartCount = item.PartCount,
                    WorkGroup = item.WorkGroup,
                    EachMachineSeq = item.EachMachineSeq
                });
            }
            
            foreach (var machine in machineGroups.Keys)
            {
                // 對每個機台，有mutationRate的概率執行突變
                if (random.NextDouble() <= mutationRate && machineGroups[machine].Count >= 2)
                {
                    // 獲取排序後的列表，但不立即賦值回字典
                    var sortedJobs = machineGroups[machine].OrderBy(x => x.StartTime).ToList();
                    // 在該機台上隨機選擇兩個工序交換
                    int jobCount = machineGroups[machine].Count;
                    int pos1 = random.Next(jobCount);
                    int pos2 = random.Next(jobCount);

                    // 確保選擇兩個不同位置
                    while (pos1 == pos2 && jobCount > 1)
                    {
                        pos2 = random.Next(jobCount);
                    }

                    if (pos1 != pos2)
                    {
                        // 獲取這兩個工序在染色體中的位置
                        var job1 = sortedJobs[pos1];
                        var job2 = sortedJobs[pos2];

                        //int idx1 = chromosome.IndexOf(job1);
                        //int idx2 = chromosome.IndexOf(job2);

                        //// 交換工序
                        //chromosome[idx1] = job2.Clone() as Chromsome;
                        //chromosome[idx2] = job1.Clone() as Chromsome;

                        //交換順序
                        schedule_after_mutation.Find(x => x.OrderID == job1.OrderID && x.Range == job1.Range).EachMachineSeq = pos2;
                        schedule_after_mutation.Find(x => x.OrderID == job2.OrderID && x.Range == job2.Range).EachMachineSeq = pos1;
                    }

                    
                }
            }

            // 重新計算時間表
            chromosome = Scheduled(schedule_after_mutation);
        }
        // 更新排程時間
        private void UpdateScheduleTimes(List<Chromsome> chromosome)
        {
            // 按工作組和開始時間排序
            var sortedJobs = chromosome.OrderBy(x => x.WorkGroup).ThenBy(x => x.StartTime).ToList();
            var machineEndTimes = new Dictionary<string, DateTime>();

            foreach (var job in sortedJobs)
            {
                if (!machineEndTimes.ContainsKey(job.WorkGroup))
                {
                    machineEndTimes[job.WorkGroup] = DateTime.MinValue;
                }

                // 設置新的開始時間
                job.StartTime = machineEndTimes[job.WorkGroup] > job.AssignDate ?
                               machineEndTimes[job.WorkGroup] :
                               job.AssignDate;

                // 計算結束時間
                DateTime endTime = job.StartTime.Add(job.Duration);
                machineEndTimes[job.WorkGroup] = endTime;

                // 計算延遲
                var dueDate = job.AssignDate; // 需要實現此方法來獲取截止日期
                if (endTime > dueDate)
                {
                    job.Delay = (int)(endTime - dueDate).TotalMinutes;
                }
                else
                {
                    job.Delay = 0;
                }
            }
        }
        // 獲取工序的截止日期(需要根據你的數據結構實現)
        private DateTime GetDueDate(string orderID, double opID)
        {
            // 實現獲取訂單工序截止日期的邏輯
            // 這裡是示例實現，你需要根據實際情況修改
            return DateTime.Now.AddDays(7); // 示例：假設截止日期為一周後
        }

        // 計算自適應突變率
        private double CalculateAdaptiveMutationRate(int noImprovementCount, double baseMutationRate)
        {
            // 基本突變率
            double mutationRate = baseMutationRate; // 例如0.05

            // 長時間沒有改進時增加突變率
            if (noImprovementCount > 5)
            {
                mutationRate = Math.Min(0.3, baseMutationRate + (noImprovementCount - 5) * 0.02);
            }

            return mutationRate;
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
            var localChromosomeList = new Dictionary<int, List<Chromsome>>(ChromosomeList); // 創建局部變量
            var updatedChromosomes = new ConcurrentDictionary<int, List<Chromsome>>(); // 使用 ConcurrentDictionary

            Parallel.For(0, crossoverResultList.Count, parallelOptions, i =>
            {
                int total = localChromosomeList[i].Count; // 正確工單制程數
                var results = new List<Tuple<string, double>>();
                for (int j = 0; j < crossoverResultList[i].Count; j++)
                {
                    results.Add(Tuple.Create(crossoverResultList[i][j].OrderID, crossoverResultList[i][j].OPID));
                }
                List<Tuple<string, double>> distinct_2 = results.Distinct().ToList();
                var distinct_1 = new List<Tuple<string, double, string, TimeSpan, DateTime, int, string>>();
                // 把遺失的工單工序加回來
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
                // 重新給定機台排序
                List<LocalMachineSeq> MachineSeq = new List<LocalMachineSeq>();
                // 獲取各機台是否為委外機台
                var OutsourcingList = getOutsourcings();
                for (int machinenameseq = 0; machinenameseq < Devices.Count; machinenameseq++)
                {
                    int seq = 0;
                    // 排序以OPID排，避免同工單後制程放在前制程前面
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
                double new_fitness_value = Calculate_Combined_Fitness(tempOrder);
                //// 多一個比較sumdelay
                //int sum = tempOrder.Sum(x => x.Delay);
                if (fitness_idx_value.Exists(x => x.Fitness > new_fitness_value)) // 判斷突變之後是否有更好的解
                {
                    int index = fitness_idx_value.FindIndex(x => x.Fitness > new_fitness_value);
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
            else
            {
                // 發現更好解，重置計數器
                noImprovementCount = 0;
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

        /// <summary>
        /// 新增計算各工單等待時間
        /// </summary>
        /// <param name="Tep"></param>
        public void Delay_and_waiting(List<Chromsome> Tep)
        {
            TimeSpan temp;
            // 計算 Delay 和 Waiting
            foreach (var group in Tep.GroupBy(x => x.OrderID))
            {
                // 按 Range 排序分組內的數據
                var orderedItems = group.OrderBy(x => x.Range).ToList();

                for (int i = 0; i < orderedItems.Count; i++)
                {
                    var item = orderedItems[i];
                    try
                    {
                        // 計算 Delay
                        temp = item.AssignDate - item.EndTime;
                        var process_data = Tep.Find(x => x.OrderID == item.OrderID && x.Range == item.Range);
                        process_data.Delay = (temp.TotalDays > 0) ? 0 : Math.Abs(temp.Days);

                        // 計算 Waiting
                        if (i == 0)
                        {
                            process_data.Waiting = 0; // 第一筆數據無等待時間
                        }
                        else
                        {
                            var previousItem = orderedItems[i - 1];
                            process_data.Waiting = (item.StartTime - previousItem.EndTime).TotalDays;
                        }
                    }
                    catch (Exception ex)
                    {
                        // 記錄異常並繼續處理下一筆
                        Console.WriteLine($"Error processing item {item.OrderID}, Range {item.Range}: {ex.Message}");
                        continue;
                    }
                }
                
            }
            Tep.OrderBy(x => x.OrderID).ThenBy(x => x.Range);
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