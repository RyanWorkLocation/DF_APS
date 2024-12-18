using PMCDash.DTO;
using PMCDash.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Services
{
    public class FitnessService
    {
        ConnectStr _ConnectStr = new ConnectStr();

        //總MakeSpan
        public double MakeSpan(string tablename)
        {
            var originSchedule = new List<ScheduleDto>();
            var sqlStr = $@"select OrderID, OPID, Optime = (MachOpTime + HumanOpTime) * OrderQTY, StartTime, EndTime, WorkGroup
                            from {_ConnectStr.APSDB}.dbo.{tablename}";
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
                                originSchedule.Add(new ScheduleDto(sqlData["OrderID"].ToString().Trim(), Convert.ToInt32(sqlData["OPID"]), sqlData["WorkGroup"].ToString(),
                                    sqlData["StartTime"].ToString(), sqlData["EndTime"].ToString(), (int)Convert.ToDouble(sqlData["Optime"].ToString())));
                            }
                        }
                    }
                }
            }

            var startime = originSchedule.OrderBy(x => x.StartTime).Select(x => x.StartTime).ToList();
            var finishtime = originSchedule.OrderBy(x => x.EndTime).Select(x => x.EndTime).ToList();

            double result = new TimeSpan(finishtime[finishtime.Count - 1].Ticks - startime[0].Ticks).TotalMinutes;
            return result;
        }

        //總延遲時間
        public double TotalDelay(string Assign)
        {
            var temp = new List<BasicChartOPData>();
            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                                w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,
                                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0), a.Scheduled
                                FROM  {_ConnectStr.APSDB}.dbo.{Assign} as a 
                                LEFT JOIN  {_ConnectStr.APSDB}.dbo.[WIP] as w
                                ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                                where (a.Scheduled = 1 or a.Scheduled = 2 or a.Scheduled = 3) and a.StartTime is not null and a.AssignDate>DATEADD(DAY,-60,GETDate()) and a.StartTime>='2024-04-01 00:00:00'
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
                                temp.Add(new BasicChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Workgroup = SqlData["WorkGroup"].ToString(),
                                    AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString(), null),
                                    AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString(), null),
                                    Real_ST = string.IsNullOrEmpty(SqlData["Real_ST"].ToString()) ? string.Empty : SqlData["Real_ST"].ToString(),
                                    Real_ET = string.IsNullOrEmpty(SqlData["Real_ET"].ToString()) ? string.Empty : SqlData["Real_ET"].ToString(),
                                    PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString("yyyy-MM-dd HH:mm"),
                                    ASDate = SqlData["ASDate"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                    Priority = SqlData["PRIORITY"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                });
                            }
                        }
                    }
                }
            }

            List<DateTime> real_dates = temp.OrderBy(x => x.AssignST).Select(x => x.AssignET).ToList();
            List<string> req_dates = temp.OrderBy(x => x.AssignST).Select(x => x.ASDate).ToList();


            double totaldiff = 0;
            for (int i = 0; i < temp.Count; i++)
            {
                //計算實際完工日期撿到欲交日期
                TimeSpan real_date = new TimeSpan(real_dates[i].Ticks); //排程完工日期
                TimeSpan req_date = new TimeSpan(DateTime.Parse(req_dates[i]).Ticks); //訂單要求日期
                var diff = new TimeSpan(real_date.Ticks - req_date.Ticks).TotalMinutes;
                if (diff > 0)
                {
                    totaldiff += diff;
                }
            }

            double result = Math.Round(totaldiff);
            return result;
        }

        //準交率計算
        public decimal ontimerate(string Assign)
        {
            var temp = new List<BasicChartOPData>();
            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                                w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,
                                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0), a.Scheduled
                                FROM  {_ConnectStr.APSDB}.dbo.{Assign} as a 
                                LEFT JOIN  {_ConnectStr.APSDB}.dbo.[WIP] as w
                                ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                                where (a.Scheduled = 1 or a.Scheduled = 2 or a.Scheduled = 3) and a.StartTime is not null and a.AssignDate>DATEADD(DAY,-60,GETDate()) and a.StartTime>='2024-04-01 00:00:00'
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
                                temp.Add(new BasicChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Workgroup = SqlData["WorkGroup"].ToString(),
                                    AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString(), null),
                                    AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString(), null),
                                    Real_ST = string.IsNullOrEmpty(SqlData["Real_ST"].ToString()) ? string.Empty : SqlData["Real_ST"].ToString(),
                                    Real_ET = string.IsNullOrEmpty(SqlData["Real_ET"].ToString()) ? string.Empty : SqlData["Real_ET"].ToString(),
                                    PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString("yyyy-MM-dd HH:mm"),
                                    ASDate = SqlData["ASDate"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                    Priority = SqlData["PRIORITY"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                });
                            }
                        }
                    }
                }
            }

            decimal result = 0;
            if (temp.Count != 0)
            {
                var orderidlist = temp.Select(x => x.OrderID).Distinct().ToList();
                //List<DateTime> real_dates = temp.OrderBy(x => x.AssignST).Select(x => x.AssignET).ToList();
                //List<string> req_dates = temp.OrderBy(x => x.AssignST).Select(x => x.ASDate).ToList();

                int count = 0;
                double diff = 0;
                for (int i = 0; i < orderidlist.Count; i++)
                {
                    var real_date = temp.Where(x => x.OrderID == orderidlist[i]).OrderByDescending(x => x.AssignET).Select(x => x.AssignET).ToList()[0];
                    var req_date = temp.Where(x => x.OrderID == orderidlist[i]).OrderByDescending(x => x.AssignET).Select(x => x.ASDate).ToList()[0];
                    //計算實際完工日期撿到欲交日期
                    //TimeSpan real_date = new TimeSpan(real_dates[i].Ticks); //排程完工日期
                    //TimeSpan req_date = new TimeSpan(DateTime.Parse(req_dates[i]).Ticks); //訂單要求日期
                    diff = (real_date - DateTime.Parse(req_date)).TotalMinutes;
                    if (diff > 0)
                    {
                        count++;
                    }
                }

                result = (1m - Math.Round((decimal)count / orderidlist.Count, 2)) * 100;
            }
            return result;
        }

        //平均稼動率
        public decimal Utilize_rate(string Assign)
        {
            var temp = new List<BasicChartOPData>();
            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                                w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,
                                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0), a.Scheduled
                                FROM  {_ConnectStr.APSDB}.dbo.{Assign} as a 
                                LEFT JOIN  {_ConnectStr.APSDB}.dbo.[WIP] as w
                                ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                                where (a.Scheduled = 1 or a.Scheduled = 2 or a.Scheduled = 3) and a.StartTime is not null and a.AssignDate>DATEADD(DAY,-60,GETDate()) and a.StartTime>='2024-04-01 00:00:00'
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
                                temp.Add(new BasicChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Workgroup = SqlData["WorkGroup"].ToString(),
                                    AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString(), null),
                                    AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString(), null),
                                    Real_ST = string.IsNullOrEmpty(SqlData["Real_ST"].ToString()) ? string.Empty : SqlData["Real_ST"].ToString(),
                                    Real_ET = string.IsNullOrEmpty(SqlData["Real_ET"].ToString()) ? string.Empty : SqlData["Real_ET"].ToString(),
                                    PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString("yyyy-MM-dd HH:mm"),
                                    ASDate = SqlData["ASDate"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                    Priority = SqlData["PRIORITY"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                });
                            }
                        }
                    }
                }
            }


            List<each_machine_time> each_machine_times = new List<each_machine_time>();
            var machines = temp.Distinct(x => x.Workgroup).Select(x => x.Workgroup).ToList();
            foreach (var i in machines)
            {
                var orders = temp.Where(x => x.Workgroup == i).ToList();
                each_machine_times.Add(new each_machine_time()
                {
                    machinename = i,
                    makesapn = mackespand(orders),//makespan
                    scheduletimes = scheduletime(orders)//各段時間
                });
            }

            decimal mackespand(List<BasicChartOPData> data)
            {
                var st = data.OrderBy(x => x.AssignST).ToList()[0].AssignST;//第一筆開工時間
                var et = data.OrderBy(x => x.AssignET).ToList()[data.Count - 1].AssignET;//最後一筆完工時間
                return decimal.Parse(new TimeSpan(et.Ticks - st.Ticks).TotalMinutes.ToString());
            }

            List<scheduletime> scheduletime(List<BasicChartOPData> data)
            {
                data = data.OrderBy(x => x.AssignST).ToList();
                var firstopst = data[0].AssignST;
                var lastopet = data[0].AssignET;
                List<scheduletime> result = new List<scheduletime>();
                for (var i = 1; i < data.Count; i++)
                {
                    if (data[i].AssignST > lastopet)
                    {
                        //result.Add(new Services.scheduletime { st = i.AssignST, et = i.AssignET });
                        result.Add(new Services.scheduletime { st = firstopst, et = lastopet });
                        firstopst = data[i].AssignST;
                        lastopet = data[i].AssignET;
                    }

                    lastopet = data[i].AssignET;
                }
                result.Add(new Services.scheduletime { st = firstopst, et = lastopet });
                return result;
            }

            decimal avg_rate=0;
            decimal counter_rate = 0;
            double usage_time = 0;
            if (each_machine_times.Count != 0)
            {
                foreach (var item in each_machine_times)
                {
                    if (item != null)
                    {
                        usage_time = 0;
                        foreach (var i in item.scheduletimes)
                        {
                            usage_time += new TimeSpan(i.et.Ticks - i.st.Ticks).TotalMinutes;
                        }
                        decimal machine_rate = Math.Round((decimal)usage_time / item.makesapn, 2);
                        counter_rate += machine_rate;
                    }
                }
                avg_rate = Math.Round((decimal)counter_rate / each_machine_times.Count, 2);
            }
            else
            {
                avg_rate = 0;
            }
            return avg_rate * 100;
        }

        //搬移次數
        public decimal Moving_times(string Assign)
        {
            var temp = new List<BasicChartOPData>();
            var SqlStr = $@"SELECT a.OrderID, a.OPID, a.WorkGroup, a.StartTime as Assign_ST, a.EndTime as Assign_ET, w.StartTime as Real_ST,
                                w.EndTime as Real_ET, a.AssignDate_PM as PMDate, a.AssignDate as ASDate, w.WIPEvent, a.PRIORITY, a.Important,
                                Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0), a.Scheduled
                                FROM  {_ConnectStr.APSDB}.dbo.{Assign} as a 
                                LEFT JOIN  {_ConnectStr.APSDB}.dbo.[WIP] as w
                                ON w.OrderID=a.OrderID AND w.OPID=a.OPID
                                where (a.Scheduled = 1 or a.Scheduled = 2 or a.Scheduled = 3) and a.StartTime is not null and a.AssignDate>DATEADD(DAY,-60,GETDate()) and a.StartTime>='2024-04-01 00:00:00'
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
                                temp.Add(new BasicChartOPData
                                {
                                    OrderID = SqlData["OrderID"].ToString().Trim(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Workgroup = SqlData["WorkGroup"].ToString(),
                                    AssignST = Convert.ToDateTime(SqlData["Assign_ST"].ToString(), null),
                                    AssignET = Convert.ToDateTime(SqlData["Assign_ET"].ToString(), null),
                                    Real_ST = string.IsNullOrEmpty(SqlData["Real_ST"].ToString()) ? string.Empty : SqlData["Real_ST"].ToString(),
                                    Real_ET = string.IsNullOrEmpty(SqlData["Real_ET"].ToString()) ? string.Empty : SqlData["Real_ET"].ToString(),
                                    PMDate = Convert.ToDateTime(SqlData["PMDate"].ToString()).ToString("yyyy-MM-dd HH:mm"),
                                    ASDate = SqlData["ASDate"].ToString(),
                                    WIPEvent = SqlData["WIPEvent"].ToString(),
                                    Priority = SqlData["PRIORITY"].ToString(),
                                    Important = SqlData["Important"].ToString(),
                                });
                            }
                        }
                    }
                }
            }

            int moving_times = 0;
            string machinetemp = "";
            var jobs = temp.OrderBy(x => x.OrderID).Distinct(x => x.OrderID).Select(x => x.OrderID).ToList();
            foreach (var job in jobs)
            {
                machinetemp = "";
                foreach (var item in temp.Where(x => x.OrderID == job).OrderBy(x => x.AssignST).ToList())
                {
                    if (machinetemp != item.Workgroup)
                    {
                        moving_times++;
                        machinetemp = item.Workgroup;
                    }
                }

            }
            return moving_times;
        }

        /// <summary>
        /// 負載平衡
        /// </summary>
        /// <param name="each_machine_time">三維資料[a,b,c] a=機台數, b=單一機台的製程數, c=開始&結束時間</param>
        /// <param name="machine">機台數</param>
        /// <returns></returns>
        public static double LoadBalance(List<List<List<int>>> each_machine_time, int machine)
        {
            List<double> machine_load = new List<double>();
            for (int i = 0; i < machine; i++)
            {
                int useage_time = 0;
                if (each_machine_time[i] != null)
                {
                    foreach (var item in each_machine_time[i])
                    {
                        useage_time += item[1] - item[0];
                    }
                }
                machine_load.Add(useage_time);
            }

            return calculate_std(machine_load);


            double calculate_std(List<double> nums)
            {
                double avg = nums.Average();
                //double std = 0;
                //foreach (double item in nums)
                //{
                //    std += Math.Pow((avg - item), 2);
                //}
                //std /= nums.Count;

                //return Math.Sqrt(std);

                //更簡潔
                return Math.Sqrt(nums.Average(n => Math.Pow(avg - n, 2)));

            }

        }

        /// <summary>
        /// 搬運成本
        /// </summary>
        /// <param name="machineAssign">[n, m] n為工單、m為各工單之製程數</param>
        /// <returns></returns>
        private static double MoveFitness(List<List<int>> machineAssign)
        {
            double fitness = 0;
            //int[,] table = readExcelTable("Mk01搬運成本表.xlsx");
            //for (int i = 0; i < machineAssign.Count; i++)
            //{
            //    int start = machineAssign[i][0];
            //    for (int j = 1; j < machineAssign[i].Count; j++)
            //    {
            //        int end = machineAssign[i][j];
            //        fitness += table[start, end];
            //    }
            //}

            return fitness;
        }
    }

    class each_machine_time
    {
        public string machinename { get; set; }
        public List<scheduletime> scheduletimes { get; set; }
        public decimal makesapn { get; set; }
    }

    class scheduletime
    {
        public DateTime st { get; set; }
        public DateTime et { get; set; }
    }
}
