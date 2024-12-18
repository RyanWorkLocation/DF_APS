using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PMCDash.Models;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class OrderController : BaseApiController
    //public class OrderController : ControllerBase
    {

        //DPI
        //soco
        //SkyMarsDB

        ConnectStr _ConnectStr = new ConnectStr();
        //private readonly string _ConnectStr.Local = @"Data Source = 127.0.0.1; Initial Catalog = DPI; User ID = MES2014; Password = PMCMES;";
        //private readonly string _ConnectStr.Local = @"Data Source = 192.168.0.156; Initial Catalog = DPI;User ID = MES2014; Password = PMCMES;";

        private readonly string _timeFormat = "yyyy-MM-dd";
        public OrderController()
        {

        }

        /// <summary>
        /// 定義各個狀態意義
        /// </summary>
        /// <returns></returns>
        [HttpGet("Define")]
        public ActionResponse<List<StatusDefine>> GetDefine()
        {
            var mockOrdersStatus = new string[] { "SSS", "SS", "S", "N" };
            var mockOrderDisplay = new string[] { "最緊急", "緊急", "急", "一般" };
            var mockProcessStatus = new string[] { "ONTIME", "DELAY" };
            var mockProcessDisplay = new string[] { "準時", "延遲" };
            var mockOrderListTypeCode = new string[] { "2", "3" };
            var mockOrderListTypeDisplay = new string[] { "訂單", "工單" };
            var OrderListType = new List<OrderListType>();
            var result = new List<StatusDefine>();
            var orderStatusList = new List<OrderStatus>();
            var processStatusList = new List<ProcessStatus>();
            for (int i = 0; i < mockOrderListTypeCode.Length; i++)
            {
                OrderListType.Add(new OrderListType(mockOrderListTypeCode[i], mockOrderListTypeDisplay[i]));
            }
            for (int i = 0; i < mockOrdersStatus.Length; i++)
            {
                orderStatusList.Add(new OrderStatus(mockOrdersStatus[i], mockOrderDisplay[i]));
            }

            for (int i = 0; i < mockProcessStatus.Length; i++)
            {
                processStatusList.Add(new ProcessStatus(mockProcessStatus[i], mockProcessDisplay[i]));
            }
            result.Add(new StatusDefine(processStatusList, orderStatusList, OrderListType));

            return new ActionResponse<List<StatusDefine>>
            {
                Data = result
            };
        }


        /// <summary>
        /// 訂單首頁畫面資料
        /// </summary>
        /// <returns></returns>
        [HttpGet("Get2")]
        public ActionResponse<List<OrderDetail>> Get2()
        {
            List<string> oStatus = new List<string> { "N", "S", "SS", "SSS" };

            var result = new List<OrderDetail>();

            //string SqlStr = $@"SELECT [Id],[OrderID],[AssignDate],[CustomerInfo],[Type],[Note]
            //                  FROM [OrderOverview] where Id>95";

            string SqlStr = $@"SELECT 
                                a.[Id],a.[OrderID],a.[AssignDate],a.[CustomerInfo],a.[Type],a.[Note],
                                b.SeriesID,b.OrderID as workOrderID,b.OPID,b.Range,
                                c.WIPEvent
                                FROM {_ConnectStr.APSDB}.[dbo].[OrderOverview] as a 
                                left join {_ConnectStr.APSDB}.[dbo].Assignment as b on a.OrderID=b.ERPOrderID
                                left join {_ConnectStr.APSDB}.[dbo].WIP as c on b.SeriesID=c.SeriesID
                                order by OrderID";

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
                            int i = 0;
                            while (SqlData.Read())
                            {
                                string[] customer = SqlData["CustomerInfo"].ToString().Split("/");
                                if (customer.Length < 3)
                                    customer = (SqlData["CustomerInfo"].ToString() + "/-").Split("/");
                                string customer_city = new Address(string.IsNullOrEmpty(customer[2]) ? "-" : customer[2]).City;
                                result.Add(new OrderDetail
                                {
                                    OrderNum = SqlData["OrderID"].ToString().Trim(),
                                    DueDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString().Trim()) ? " - " : Convert.ToDateTime(SqlData["AssignDate"].ToString().Trim()).ToString(_timeFormat),
                                    DueDatePM = string.IsNullOrEmpty(SqlData["AssignDate"].ToString().Trim()) ? " - " : Convert.ToDateTime(SqlData["AssignDate"].ToString().Trim()).AddDays(-10).ToString(_timeFormat),
                                    CustomerInfo = string.IsNullOrEmpty(customer[1]) ? "-" : customer[1] + "/" + (string.IsNullOrEmpty(customer_city) ? "-" : customer_city),
                                    Status = oStatus[int.Parse(SqlData["type"].ToString())],
                                    OrderNote = string.IsNullOrEmpty(SqlData["Note"].ToString().Trim()) ? " - " : SqlData["Note"].ToString().Trim()
                                });

                                if (result[i].DueDate != " - ")
                                {
                                    if (Delayday(DateTime.Now.ToString(), result[i].DueDate) >= 0)
                                    {
                                        result[i].ProcessProgress = "ONTIME";
                                    }
                                    else if (Delayday(DateTime.Now.ToString(), result[i].DueDate) < 0)
                                    {
                                        result[i].ProcessProgress = "DELAY";
                                    }
                                }
                                else
                                {
                                    result[i].ProcessProgress = " - ";
                                }

                                var WOdetail = GetWorkOrdersOfOrder(result[i].OrderNum);

                                int progress = orderprogress(result[i].OrderNum);

                                if (WOdetail.Data.Count > 0)//有工單資料
                                {
                                    result[i].ProgressValue = Convert.ToInt32((Convert.ToDouble(WOdetail.Data.FindAll(x => x.ProgressValue >= 100).Count()) / Convert.ToDouble(WOdetail.Data.Count())) * 100);
                                }
                                else
                                {
                                    result[i].ProgressValue = 0;
                                }

                                result[i].ProgressValue = progress;

                                result[i].NeedDay = "尚須" + new Random().Next(3, 90).ToString() + "天";

                                i += 1;

                            }
                        }
                    }
                }
            }

            //將交期已過且生產進度大於等於100%的部分過濾掉部顯示
            result = result.OrderBy(x => x.DueDate).ToList();

            return new ActionResponse<List<OrderDetail>>
            {
                Data = result
            };
        }

        public ActionResponse<List<OrderDetail>> Get()
        {
            var result = new List<OrderDetail>();
            var temp = GetOrderOverView();//取得三個月以內的訂單資訊

            //取得不重複的訂單號
            var orders = temp.OrderBy(x => x.OrderID).Select(x => x.OrderID).Distinct().ToList();

            //保留工單都已經做完的order
            var orders_donelist = new List<string>(); //全部完工的訂單
            var orders_notdonelist = new List<string>(); //尚未完工的訂單
            List<OrderOverView> orderdata;
            foreach (var item in orders)
            {
                orderdata = temp.Where(x => x.OrderID == item).ToList();
                if (orderdata.FindAll(x => x.WIPEvent == 3).Count == orderdata.Count)
                    orders_donelist.Add(item);
                else
                    orders_notdonelist.Add(item);
            }

            //若為未全部完成之訂單
            OrderOverView data;
            foreach (var item in orders_notdonelist)
            {
                data = temp.Where(x => x.OrderID == item).First();
                result.Add(new OrderDetail
                {
                    OrderNum = data.OrderID,
                    DueDate = data.AssignDate.ToString(_timeFormat),
                    //DueDatePM = data.AssignDate.ToString(_timeFormat),
                    DueDatePM = string.IsNullOrEmpty(data.EstimatedDate.ToString())?"-": data.EstimatedDate.ToString(),
                    OrderNote = data.Note,
                    CustomerInfo = getCustomerInfo(data.CustomerInfo),
                    ProcessProgress = getStatus(data.AssignDate.ToString(), string.IsNullOrEmpty(data.EstimatedDate.ToString()) ? "-" : data.EstimatedDate.ToString()),
                    Status = "N",
                    ProgressValue = getProgress(item),
                    UndonePart = getDelayProduct(item).Trim(),
                    //UndonePart = getProgress(item) < 100 ? GetWorkOrdersOfOrder(item).Data.Find(x => x.ProgressValue < 100).PartName : "-",
                    NeedDay = getNeedDay(item)
                });
            }

           
            string getStatus(string DueDate, string DueDate_PM)
            {
                //if (Delayday(DateTime.Now.ToString(), DueDate) >= 0)
                //    return "ONTIME";
                if(DueDate_PM != "-")
                {
                    if (Delayday(DueDate_PM, DueDate) >= 0)
                    {
                        return "ONTIME";
                    }
                    else
                    {
                        return "DELAY";
                    }
                        
                }
                else
                {
                    if (Delayday(DateTime.Now.ToString(), DueDate) >= 0)
                    {
                        return "ONTIME";
                    }
                    else
                    {
                        return "DELAY";
                    }
                }
            }

            int getProgress(string data)
            {
                var list = temp.Where(x => x.OrderID == data).ToList();
                double a = (double)list.FindAll(x => x.WIPEvent == 3).Count;
                double b = (double)list.Count;
                return (int)(a / b * 100);
            }

            //取得卡單製程品項
            string getDelayProduct(string ordernum)
            {
                var list = temp.Where(x => x.OrderID == ordernum).OrderBy(x=>x.workOrderID).ThenBy(x=>x.Range).ToList();
                var undonepart = list.Find(x => x.WIPEvent != 3);
                return undonepart.PartOPName;
            }

            
            string getCustomerInfo(string data)
            {
                var text = data.Split('/');
                if (text[2] != "...")
                    return text[1] + "/" + new Address(text[2]).City;
                else
                    return text[1] + "/" + text[2];
            }

            //宛玲建議:將剩下未完成的作量結合標準工時，計算還須要花多久時間，參考照片
            string getNeedDay(string data)
            {
                var list = temp.Where(x => x.OrderID == data && x.WIPEvent != 3).OrderBy(x => x.Range).ToList();
                double totalTime = 0;
                double totalWaitTime = 0;
                for (int i = 1; i < list.Count; i++)
                {
                    if (list[i].StartTime!=null && list[i].EndTime != null)
                    {
                        totalTime += list[i].EndTime.Value.Subtract(list[i].StartTime.Value).TotalMinutes;//加工時間
                        if(list[i - 1].EndTime != null)
                        {
                            totalWaitTime += list[i].StartTime.Value.Subtract(list[i - 1].EndTime.Value).TotalMinutes;//等待時間
                        }
                    }
                    else
                    {
                        totalTime += Convert.ToDouble(list[i].ProcessTime);
                    }
                }
                if (list[0].StartTime != null && list[0].EndTime != null)
                {
                    totalTime += list[0].EndTime.Value.Subtract(list[0].StartTime.Value).TotalMinutes;//補上第一道製程工時
                }
                else
                {
                    totalTime += Convert.ToDouble(list[0].ProcessTime);
                }
                    

                TimeSpan ts = new TimeSpan(0, (int)(Math.Abs(totalTime) + Math.Abs(totalWaitTime)), 0);

                double day = ts.TotalDays;
                if (day >= 0 && day <= 1)
                    return "小於1天";
                else if (day < 0)
                    return "無法計算";
                else
                {
                    if (day < 30)
                        return "約" + GetTimeDiff(DateTime.Now.ToString(_timeFormat), DateTime.Now.AddMinutes(ts.TotalMinutes).ToString(_timeFormat), "D") + "天";
                    else
                        return "約" + GetTimeDiff(DateTime.Now.ToString(_timeFormat), DateTime.Now.AddMinutes(ts.TotalMinutes).ToString(_timeFormat), "M") + "個月";
                }


            }

            ////將已完成訂單加入回傳結果
            //foreach (var item in orders_donelist)
            //{
            //    //取得最後一道完成的製程時間
            //    var LastOPComplieTime = temp.Where(x => x.OrderID == item).OrderByDescending(x => x.EndTime).ToList();
            //    //是否已經完成超過30天以上，若超過30天以上就不顯示啦
            //    if (LastOPComplieTime.First().EndTime > DateTime.Now.AddDays(-30))
            //    {
            //        data = temp.Where(x => x.OrderID == item).First();
            //        result.Add(new OrderDetail
            //        {
            //            OrderNum = data.OrderID,
            //            DueDate = data.AssignDate.ToString(_timeFormat),
            //            DueDatePM = data.AssignDate.ToString(_timeFormat),
            //            OrderNote = data.Note,
            //            CustomerInfo = getCustomerInfo(data.CustomerInfo),
            //            ProcessProgress = getStatus(data.AssignDate.ToString(_timeFormat)),
            //            Status = "N",
            //            ProgressValue = 100,//getProgress(item),
            //            NeedDay = "0天"
            //        });
            //    }
            //}

            return new ActionResponse<List<OrderDetail>>
            {
                Data = result.OrderBy(x => x.OrderNum).ToList()
            };
        }

        /// <summary>
        /// 計算兩個日期相差天數或月數
        /// </summary>
        /// <param name="strFrom">開始時間</param>
        /// <param name="strTo">結束時間</param>
        /// <param name="strType">日:D/月:M</param>
        /// <returns></returns>
        private int GetTimeDiff(string strFrom, string strTo, string strType)
        {
            DateTime dtStart = DateTime.Parse(strFrom);
            DateTime dtEnd = DateTime.Parse(strTo);

            if (strType == "D")
            {
                //使用TimeSpan提供的Days屬性
                TimeSpan ts = (dtEnd - dtStart);
                int iDays = ts.Days + 1;
                return iDays;
            }
            else if (strType == "M")
            {
                int iMonths = dtEnd.Year * 12 + dtEnd.Month - (dtStart.Year * 12 + dtStart.Month) + 1;
                return iMonths;
            }
            else return 0;
        }


        private List<OrderOverView> GetOrderOverView()
        {
            DateTime validValue;

            var result = new List<OrderOverView>();
            //string SqlStr = $@"
            //                    SELECT 
            //                    a.[Id],a.[OrderID],a.[AssignDate],a.[CustomerInfo],a.[Type],a.[Note],
            //                    b.SeriesID,b.OrderID as workOrderID,b.OPID,b.Range,b.StartTime,b.EndTime,
            //                    b.MachOpTime,
            //                    c.WIPEvent
            //                    FROM [OrderOverview] as a 
            //                    inner join Assignment as b on a.OrderID=b.ERPOrderID
            //                    inner join WIP as c on b.SeriesID=c.SeriesID
            //                    where a.showed=1 and a.Id>(SELECT TOP (1) [Id] FROM [OrderOverview] where OrderID like '%B{DateTime.Now.AddMonths(-3).ToString("yyMM")}%' order by id desc)
            //                    order by OrderID";
            string SqlStr = $@"SELECT 
                            a.[Id], a.[OrderID], b.OPLTXA1 ,a.[AssignDate],
                            (SELECT MAX(EndTime) FROM {_ConnectStr.APSDB}.[dbo].Assignment WHERE ERPOrderID = a.OrderID) AS LatestEndTime, a.[CustomerInfo], a.[Type], a.[Note],
                            b.SeriesID, b.OrderID as workOrderID, b.OPID, b.Range, b.StartTime, b.EndTime,
                            b.MachOpTime,
                            c.WIPEvent
                            FROM  
                            {_ConnectStr.APSDB}.[dbo].[OrderOverview] as a 
                            INNER JOIN {_ConnectStr.APSDB}.[dbo].Assignment as b ON a.OrderID = b.ERPOrderID
                            INNER JOIN {_ConnectStr.APSDB}.[dbo].WIP as c ON b.SeriesID = c.SeriesID
                            INNER JOIN {_ConnectStr.MRPDB}.[dbo].Part as  p ON b.MAKTX = p.Number 
                            WHERE 
                            a.showed = 1 
                            AND a.Id > (SELECT TOP (1) [Id] FROM {_ConnectStr.APSDB}.[dbo].[OrderOverview] WHERE OrderID LIKE '%B{DateTime.Now.AddMonths(-3).ToString("yyMM")}%' ORDER BY Id DESC)
                            ORDER BY 
                            a.OrderID";
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
                                result.Add(new OrderOverView
                                {
                                    Id = int.Parse(SqlData["Id"].ToString()),
                                    OrderID = SqlData["OrderID"].ToString(),
                                    PartOPName = Regex.Replace(SqlData["OPLTXA1"].ToString(), "[a-zA-Z0-9!@#$%^&*()\\-+]", ""),
                                    AssignDate = DateTime.Parse(SqlData["AssignDate"].ToString()),
                                    EstimatedDate = string.IsNullOrEmpty(SqlData["LatestEndTime"].ToString())?"-":DateTime.Parse(SqlData["LatestEndTime"].ToString()).ToString(_timeFormat),
                                    CustomerInfo = SqlData["CustomerInfo"].ToString(),
                                    Note = SqlData["Note"].ToString(),
                                    SeriesID = int.Parse(SqlData["SeriesID"].ToString()),
                                    workOrderID = SqlData["workOrderID"].ToString(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Range = int.Parse(SqlData["Range"].ToString()),
                                    StartTime = DateTime.TryParse(SqlData["StartTime"].ToString(), out validValue) ? validValue : (DateTime?)null,
                                    EndTime = DateTime.TryParse(SqlData["EndTime"].ToString(), out validValue) ? validValue : (DateTime?)null,
                                    WIPEvent = int.Parse(SqlData["WIPEvent"].ToString()),
                                    ProcessTime = SqlData["MachOpTime"].ToString()
                                });
                            }
                        }
                    }
                }
            }
            return result;
        }

        /// <summary>
        /// 結單API
        /// </summary>
        /// <param name="ordernum"></param>
        /// <returns></returns>
        [HttpGet("EndOrder")]
        public string EndOrder(string ordernum)
        {
            var result = string.Empty;
            string SqlStr = $@"
                                UPDATE {_ConnectStr.APSDB}.[dbo].[OrderOverview]
                                SET showed=0 where OrderID=@OrderID";
            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = ordernum;
                    int line = comm.ExecuteNonQuery();
                    if(line>0)
                    {
                        result = "Update successful";
                    }
                    else
                    {
                        result = "Update failed";
                    }
                }
            }

            return result;
        }

        private int orderprogress(string orderNum)
        {
            string SqlStr = $@"select * from {_ConnectStr.APSDB}.[dbo].Assignment as a
                            left join {_ConnectStr.APSDB}.[dbo].WIP as b on a.OrderID = b.OrderID and a.OPID = b.OPID
                            where a.ERPOrderID = @ERPOrderID";
            //接一下訂單內功單生產狀況
            List<string> wipevents = new List<string>();
            //要回傳的%數
            int result = 0;

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@ERPOrderID", SqlDbType.NVarChar).Value = orderNum;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            int i = 0;
                            while (SqlData.Read())
                            {
                                wipevents.Add(SqlData["WIPEvent"].ToString().Trim());
                            }
                        }
                    }
                }
            }

            if (wipevents.Count != 0)
            {
                int done = wipevents.Where(x => x == "3").Count();
                result = (int)((double)done / (double)wipevents.Count() * 100);
            }

            return result;
        }

        /// <summary>
        /// 用訂單編號找尋所有工單內容
        /// </summary>
        /// <param name="orderNum">訂單編號</param>
        /// <returns>
        /// 1219111327
        /// </returns>
        [HttpGet("Information/{orderNum}")]
        public ActionResponse<List<WorkOrder>> GetWorkOrdersOfOrder(string orderNum)
        {
            var result = new List<WorkOrder>();

            var SqlStr = $@"select 
                            [OrderID],[WorkOrderID],[Name],[DrawingNname],[ProcessMethod],[MaterialID],
                            [CusDevice],[Remark],[MAKTX],[OrderQTY],[AssignDate_PM],[AssignDate],
                            (select top(1) EndTime from {_ConnectStr.APSDB}.[dbo].Assignment where OrderID=ss.WorkOrderID order by Range desc) as LatestEndTime from
                            (select OrderID,WorkOrderID,MAKTX,OrderQTY,AssignDate_PM,AssignDate from {_ConnectStr.APSDB}.[dbo].WorkOrderOverview 
                            group by OrderID,WorkOrderID,MAKTX,OrderQTY,AssignDate_PM,AssignDate having count(*)>=1) as ss
                            left join {_ConnectStr.MRPDB}.dbo.Part as b on ss.MAKTX=b.Number
                            where ss.OrderID=@OrderID";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = orderNum;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            int i = 0;
                            while (SqlData.Read())
                            {

                                result.Add(new WorkOrder
                                {
                                    OrderNo = SqlData["OrderID"].ToString().Trim(),
                                    WONo = SqlData["WorkOrderID"].ToString().Trim(),
                                    PartName = string.IsNullOrEmpty(SqlData["Name"].ToString().Trim()) ? " - " : SqlData["Name"].ToString().Trim(),//SqlData["Name"].ToString().Trim(),
                                    DrawingNname = string.IsNullOrEmpty(SqlData["DrawingNname"].ToString().Trim()) ? " - " : SqlData["DrawingNname"].ToString().Trim(),
                                    ProcessMethod = string.IsNullOrEmpty(SqlData["ProcessMethod"].ToString().Trim()) ? " - " : SqlData["ProcessMethod"].ToString().Trim(),//SqlData["ProcessMethod"].ToString().Trim(),
                                    ProductMaterial = string.IsNullOrEmpty(SqlData["MaterialID"].ToString().Trim()) ? " - " : SqlData["MaterialID"].ToString().Trim(),
                                    CusDevice = string.IsNullOrEmpty(SqlData["CusDevice"].ToString().Trim()) ? " - " : SqlData["CusDevice"].ToString().Trim(),
                                    //Noet = string.IsNullOrEmpty(SqlData["Remark"].ToString().Trim()) ? " - " : SqlData["Remark"].ToString().Trim(),
                                    Noet = string.IsNullOrEmpty(SqlData["LatestEndTime"].ToString()) ? "-" : DateTime.Parse(SqlData["LatestEndTime"].ToString()).ToString(_timeFormat),
                                    MaterialNo = SqlData["MAKTX"].ToString().Trim(),
                                    RequireCount = Convert.ToInt32(SqlData["OrderQTY"].ToString().Trim()),
                                    DueDatePM = string.IsNullOrEmpty(SqlData["AssignDate_PM"].ToString().Trim()) ? " - " : Convert.ToDateTime(SqlData["AssignDate_PM"].ToString().Trim()).ToString(_timeFormat),
                                    DueDate = string.IsNullOrEmpty(SqlData["AssignDate"].ToString().Trim()) ? " - " : Convert.ToDateTime(SqlData["AssignDate"].ToString().Trim()).ToString(_timeFormat)
                                });

                                result[i].ProcessInfos = CreateWorkOrderInfo(result[i].OrderNo, result[i].WONo);
                                result[i].ProgressValue = ProcessAVG(result[i].ProcessInfos);

                                if (result[i].ProcessInfos.Count() > 0)//有製程資料
                                {
                                    double temp = Math.Round((Convert.ToDouble(result[i].ProcessInfos.FindAll(x => x.ProgressValue >= 100).Count()) / Convert.ToDouble(result[i].ProcessInfos.Count())) * 100, 0, MidpointRounding.AwayFromZero);
                                    if (result[i].ProgressValue > 0) ;
                                    result[i].ProgressValue = Convert.ToInt32(temp);//計算已完成製程的比例
                                }
                                else
                                {
                                    result[i].ProgressValue = 0;
                                }

                                result[i].NeedDay = getNeedDay_workorder(SqlData["WorkOrderID"].ToString().Trim());

                                i += 1;
                            }
                        }
                    }
                }
            }

            return new ActionResponse<List<WorkOrder>>
            {
                Data = result
            };
        }

        

        private string getNeedDay_workorder(string OrderID)
        {
            string result = "";
            DateTime validValue;
            List<OrderOverView> data = new List<OrderOverView>();

            string SqlStr = $@"SELECT 
                                a.[Id],a.[OrderID],a.[AssignDate],a.[CustomerInfo],a.[Type],a.[Note],
                                b.SeriesID,b.OrderID as workOrderID,b.OPID,b.Range,b.StartTime,b.EndTime,
                                b.MachOpTime,
                                c.WIPEvent
                                FROM {_ConnectStr.APSDB}.[dbo].[OrderOverview] as a 
                                inner join {_ConnectStr.APSDB}.[dbo].Assignment as b on a.OrderID=b.ERPOrderID
                                inner join {_ConnectStr.APSDB}.[dbo].WIP as c on b.SeriesID=c.SeriesID
                                where b.OrderID='{OrderID}'
                                order by OrderID,b.Range";
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
                                data.Add(new OrderOverView
                                {
                                    Id = int.Parse(SqlData["Id"].ToString()),
                                    OrderID = SqlData["OrderID"].ToString(),
                                    AssignDate = DateTime.Parse(SqlData["AssignDate"].ToString()),
                                    CustomerInfo = SqlData["CustomerInfo"].ToString(),
                                    Note = SqlData["Note"].ToString(),
                                    SeriesID = int.Parse(SqlData["SeriesID"].ToString()),
                                    workOrderID = SqlData["workOrderID"].ToString(),
                                    OPID = SqlData["OPID"].ToString(),
                                    Range = int.Parse(SqlData["Range"].ToString()),
                                    StartTime = DateTime.TryParse(SqlData["StartTime"].ToString(), out validValue) ? validValue : (DateTime?)null,
                                    EndTime = DateTime.TryParse(SqlData["EndTime"].ToString(), out validValue) ? validValue : (DateTime?)null,
                                    WIPEvent = int.Parse(SqlData["WIPEvent"].ToString()),
                                    ProcessTime = SqlData["MachOpTime"].ToString()
                                });
                            }
                        }
                    }
                }

            }

            var list = data.Where(x => x.workOrderID == OrderID && x.WIPEvent != 3).OrderBy(x => x.Range).ToList();
            double totalTime = 0;
            double totalWaitTime = 0;
            if (list.Count != 0)
            {
                if (list.Count >= 2)
                {
                    for (int i = 1; i < list.Count; i++)
                    {
                        if (list[i].StartTime != null && list[i].EndTime != null && list[i - 1].StartTime != null && list[i - 1].EndTime != null)
                        {
                            totalTime += list[i].EndTime.Value.Subtract(list[i].StartTime.Value).TotalMinutes;
                            if (list[i].EndTime != null)
                            {
                                totalWaitTime += list[i].StartTime.Value.Subtract(list[i - 1].EndTime.Value).TotalMinutes;
                            }
                        }
                        else
                        {
                            totalTime += Convert.ToDouble(list[i].ProcessTime);
                        }

                    }
                }
                else
                {
                    if (list[0].StartTime != null && list[0].EndTime != null)
                    {
                        totalTime += list[0].EndTime.Value.Subtract(list[0].StartTime.Value).TotalMinutes;
                    }
                    else
                    {
                        totalTime += Convert.ToDouble(list[0].ProcessTime);
                    }
                }
                TimeSpan ts = new TimeSpan(0, (int)(totalTime + totalWaitTime), 0);

                double day = ts.TotalDays;
                if (day >= 0 && day <= 1)
                    result = "小於1天";
                else if (day < 0)
                    result = "無法計算";
                else
                {
                    if (day < 30)
                        result = "約" + GetTimeDiff(DateTime.Now.ToString(_timeFormat), DateTime.Now.AddMinutes(ts.TotalMinutes).ToString(_timeFormat), "D") + "天";
                    else
                        result = "約" + GetTimeDiff(DateTime.Now.ToString(_timeFormat), DateTime.Now.AddMinutes(ts.TotalMinutes).ToString(_timeFormat), "M") + "個月";
                }
            }
            else
            {
                return result = "0 天";
            }


            return result;
        }

        private int ProcessAVG(List<Process> data)
        {
            double avgResult = 0;
            if (data.Count != 0)
                avgResult = data.Average(x => x.ProgressValue);
            return (int)Math.Ceiling(avgResult);
        }

        /// <summary>
        /// 生產狀態-製程資料
        /// </summary>
        /// <param name="orderNum">訂單編號</param>
        /// <param name="workorderNum">工單編號</param>
        /// <returns></returns>
        private List<Process> CreateWorkOrderInfo(string orderNum, string workorderNum)
        {
            var result = new List<Process>();
            string SqlStr = "";
            //var SqlStr = $@"select ass.OPID,ass.Range,ass.OrderQTY,ass.OPLTXA1,Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0),　ass.Note as workorderNote
            //                ,o.OrderID,o.WorkOrderID,o.Progress,o.CPK,o.Note as ProcessNote
            //                ,w.WIPEvent,w.StartTime,w.EndTime
            //                ,p.*
            //                from Assignment as ass 
            //                left join WIP as w on ass.OrderID = w.OrderID and ass.OPID=w.OPID
            //                inner join OperationOverview as o on ass.OrderID=o.WorkOrderID and ass.OPID=o.OPID
            //                inner join {_ConnectStr.MRPDB}.dbo.Process as p on o.OPID=p.ID
            //                where o.WorkOrderID=@OrderID";
            SqlStr = $@"select 
                        Progress = ROUND(cast(cast(w.QtyTol as float) / cast(w.OrderQTY as float) * 100 as decimal), 0),
                        ass.OrderID,ass.OPID,ass.Range,ass.OPLTXA1,ass.OrderQTY,ass.CPK,
                        w.WIPEvent,w.StartTime as wipStartTime,
                        w.EndTime as wipEndTime,
                        p.FirstItemDay,p.FirstItem
                        from {_ConnectStr.APSDB}.[dbo].Assignment as ass 
                        left join {_ConnectStr.APSDB}.[dbo].WIP as w on ass.OrderID = w.OrderID and ass.OPID=w.OPID
                        left join {_ConnectStr.MRPDB}.dbo.Process as p on w.OPID=p.ID
                        where ass.OrderID=@OrderID
                        order by ass.Range";

            var cpk = new string[] { "合格", "不合格" };

            Random rnd = new Random(Guid.NewGuid().GetHashCode());

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    comm.Parameters.Add("@OrderID", SqlDbType.NVarChar).Value = workorderNum;
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            int i = 0;
                            while (SqlData.Read())
                            {
                                string wipST = string.IsNullOrEmpty(SqlData["wipStartTime"].ToString()) ? "-" : DateTime.Parse(SqlData["wipStartTime"].ToString()).ToString("MM/dd HH:mm");
                                string wipET = string.IsNullOrEmpty(SqlData["wipEndTime"].ToString()) ? "-" : DateTime.Parse(SqlData["wipEndTime"].ToString()).ToString("MM/dd HH:mm");
                                string remark = "實際開工:" + wipST + " 實際完工:" + wipET;
                                result.Add(new Process
                                {
                                    No = SqlData["OPID"].ToString().Trim(),
                                    Range = int.Parse(SqlData["Range"].ToString()),
                                    Name = SqlData["OPLTXA1"].ToString().Trim(),
                                    ProgressValue = ProgressTransfer(SqlData["Progress"].ToString().Trim()),
                                    CPKEvaluation = cpk[0],//待處理
                                    Remark = remark,//string.IsNullOrEmpty(SqlData["Note"].ToString().Trim()) ? " - " : SqlData["Note"].ToString().Trim(),
                                    ExecutionTime = getExecutionTime(SqlData["wipStartTime"].ToString().Trim(), SqlData["wipEndTime"].ToString().Trim(), SqlData["WIPEvent"].ToString().Trim()),
                                    IsDelay = getExecutionTimeAction(SqlData)
                                });

                                if (!string.IsNullOrEmpty(SqlData["CPK"].ToString().Trim()))
                                {
                                    if (Convert.ToDouble(SqlData["CPK"].ToString().Trim()) < 1.33)
                                    {
                                        result[i].CPKEvaluation = "立即改善";
                                    }
                                    else if (Convert.ToDouble(SqlData["CPK"].ToString().Trim()) >= 1.33 && Convert.ToDouble(SqlData["CPK"].ToString().Trim()) <= 1.67)
                                    {
                                        result[i].CPKEvaluation = "不合格";
                                    }
                                    else
                                    {
                                        result[i].CPKEvaluation = "合格";
                                    }
                                }
                                else
                                {
                                    result[i].CPKEvaluation = " - ";
                                }

                                i += 1;
                            }


                        }
                    }
                }
            }
            result = result.OrderBy(x => x.Range).ToList();
            return result;

        }

        //計算是否超過製程標準時間
        private bool getExecutionTimeAction(SqlDataReader dataReader)
        {
            bool result = false;
            int standadday = int.Parse(string.IsNullOrEmpty(dataReader["FirstItemDay"].ToString()) ? "0" : dataReader["FirstItemDay"].ToString());
            string[] standadtime = (string.IsNullOrEmpty(dataReader["FirstItem"].ToString()) ? "00:00:00" : dataReader["FirstItem"].ToString()).Split(":");

            TimeSpan timedaytosecond = new TimeSpan(standadday, 0, 0, 0);
            TimeSpan timehourstosecond = new TimeSpan(0, int.Parse(standadtime[0]), int.Parse(standadtime[1]), int.Parse(standadtime[2]));
            TimeSpan toltalstandadsecond = timedaytosecond + timehourstosecond * int.Parse(dataReader["OrderQTY"].ToString());

            string st = dataReader["wipStartTime"].ToString();
            string et = dataReader["wipEndTime"].ToString();

            if (st != "" && et == "")
            {
                if (DateTime.Now.Subtract(DateTime.Parse(st)) <= toltalstandadsecond)
                    result = false;
                else
                    result = true;
            }
            else if (st != "" && et != "")
            {
                if (DateTime.Parse(et).Subtract(DateTime.Parse(st)) <= toltalstandadsecond)
                    result = false;
                else
                    result = true;
            }
            return result;
        }

        //計算執行中的製程執行多久時間了
        private string getExecutionTime(string starttime, string endtime, string wipwvent)
        {
            string result = "00:00:00";

            if (starttime != "" && endtime == "")
            {
                TimeSpan ts = DateTime.Now.Subtract(DateTime.Parse(starttime));
                if (ts.Days >= 1)
                    result = ts.Days + "(d):" + ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00");
                else
                    result = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00") + "(m)";
            }
            else if (starttime != "" && endtime != "")
            {
                TimeSpan ts = DateTime.Parse(endtime).Subtract(DateTime.Parse(starttime));
                if (ts.Days >= 1)
                    result = ts.Days + "(d):" + ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00");
                else
                    result = ts.Hours.ToString("00") + ":" + ts.Minutes.ToString("00") + "(m)";
            }
            else if (starttime == "" && endtime == "")
            {
                result = "未開工";
            }
            return result;
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

        private int ProgressTransfer(string data)
        {
            if (!string.IsNullOrEmpty(data))
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
        /// 計算延遲天數(time2-time1)
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
        /// 取得表頭順序 (O)*
        /// </summary>
        /// <returns>2:訂單列表、3:工單列表</returns>
        [HttpGet("Sequence/{tabletype}")]
        public ActionResponse<List<Dashboard>> DashboardSeq(string tabletype)
        {
            var result = new List<Dashboard>();
            var SqlStr = $@"SELECT *
                           FROM Dashboard
                           where page={tabletype}
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
                                    Name = _checkNoword(SqlData["Name"].ToString()),
                                    Sequence = Convert.ToInt16(SqlData["Sequence"]),
                                    Key = _checkNoword(SqlData["EngName"].ToString()),
                                    DateType = _checkNoword(SqlData["DataType"].ToString().Trim()),
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
        /// 更新表頭順序 (O)*
        /// </summary>
        /// <param name="DashboardSeq">請輸入欲更改欄位之中文名稱與順序</param>
        /// <returns></returns>
        [HttpPost("SequenceUpdate")]
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
                            SqlStr += $"UPDATE {_ConnectStr.APSDB}.[dbo].Dashboard SET Sequence = {DashboardSeq[i].Sequence} WHERE Name = '{DashboardSeq[i].Name}' and Page={DashboardSeq[i].Page}" + Environment.NewLine;
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
    }

    class OrderOverView
    {
        public int Id { get; set; }
        public string OrderID { get; set; }
        public string PartOPName { get; set; }
        public DateTime AssignDate { get; set; }
        public string CustomerInfo { get; set; }
        public string Note { get; set; }
        public int SeriesID { get; set; }
        public string workOrderID { get; set; }
        public string OPID { get; set; }
        public int Range { get; set; }
        public DateTime? StartTime { get; set; }
        public DateTime? EndTime { get; set; }
        public int WIPEvent { get; set; }
        public string ProcessTime { get; set; }
        /// <summary>
        /// 排程估計交期
        /// </summary>
        public string EstimatedDate { get; set; }
    }
}
