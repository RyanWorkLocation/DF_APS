using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PMCDash.Models;
using System;
using System.Data;
using System.Collections.Generic;
using PMCDash.DTO;
using System.Linq;
using System.Data.SqlClient;

namespace PMCDash.Controllers
{
    
    [Route("api/[controller]")]
    public class ProductController : BaseApiController
    {
        //DPI
        //soco
        //SkyMarsDB
        ConnectStr _ConnectStr = new ConnectStr();

        private readonly string _timeFormat = "yyyy/MM/dd";
        public ProductController()
        {
            
        }

        /// <summary>
        ///  統計製程管理(物料製程總表)
        /// </summary>
        /// <returns></returns>
        public ActionResponse<List<Material>> Get()
        {
            var result = new List<Material>();
            var SqlStr = $@"select distinct ass.OPID,　ass.MAKTX,ass.CPK, ass.OPLTXA1,max(w.CreateTime) as LastProduceTime
                            from Assignment as ass left join WIPLog as w
                            on ass.OrderID = w.OrderID and ass.OPID=w.OPID
                            group by  ass.CPK,ass.OPID,ass.MAKTX, ass.OPLTXA1
                            order by LastProduceTime, MAKTX, OPID";
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
                            int i = 0;
                            while (SqlData.Read())
                            {
                                result.Add(new Material
                                {
                                    No = SqlData["MAKTX"].ToString().Trim(),
                                    ProcessNo = SqlData["OPID"].ToString().Trim(),
                                    //CPKEvaluation = cpk[rnd.Next(cpk.Length)],
                                    ProcessName = SqlData["OPLTXA1"].ToString().Trim(),
                                    LastProduction = string.IsNullOrEmpty(SqlData["LastProduceTime"].ToString()) ? "N/A" : DateTime.Parse(SqlData["LastProduceTime"].ToString()).ToString(_timeFormat)
                                }) ;

                                if(!string.IsNullOrEmpty(SqlData["CPK"].ToString().Trim()))
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
                                    result[i].CPKEvaluation = "N/A";
                                }
                                i += 1;
                                
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

        /// <summary>
        /// 透過物料編號、製程編號、資料筆數，取回品檢項目CPK分析資料
        /// </summary>
        /// <param name="request">物料編號、製程編號、資料筆數(目前可使用30、60、90)</param>
        /// <returns></returns>
        [HttpPost("CPK")]
        public ActionResponse<List<Quality>> GetCPKDetails([FromBody] MaterialDto request)
        {
            var spc = new string[] { "內牙內徑", "內牙外徑", "外牙外徑", "外牙內徑" };
            var result = new List<Quality>();
            var cl = 41.5M;
            var USL = new decimal[] { 41.65M, 41.575M };
            var LSL = new decimal[] { 41.35M, 41.275M };
            var Ca =  new decimal[] { 0.4600M, 0.4129M, 0.4509M};
            var Cp = new decimal[] { 2.9940M,  1.7004M, 2.3941M};
            var CPK = new decimal[] { 1.616M,  0.9983M,  1.3146M};
            var dataRand = new decimal[] { 41.44M, 41.34M, 41.42M, 41.52M, 41.53M, 41.44M, 41.34M, 41.42M,
                41.52M, 41.53M, 41.44M, 41.34M, 41.42M, 41.52M, 41.53M, 41.44M, 41.34M, 41.42M, 41.52M, 41.53M,
                41.44M, 41.34M, 41.42M, 41.52M, 41.53M, 41.44M, 41.34M, 41.42M, 41.52M, 41.53M };
            var avg = 41.45M;
            var mse = 0.0698M;
            var rm = new List<decimal>();

            var times = Convert.ToInt32(request.ProcessNo);
            var count = 1;
            if(times > 20)
            {
                count = 3;
            }
            List<decimal> test = new List<decimal>(dataRand);
            if(request.Count != test.Count)
            {
                for (int i = 0; i < request.Count / 30 - 1; i++)
                {
                    test.AddRange(dataRand);
                }
            }
        
            for(int i = 0; i < test.Count - 1; i++)
            {
                rm.Add(Math.Abs(test[i + 1] - test[i]));
            }

            var temp = new List<WorkOrderInfo>();
            var testCount = 15;
            var testWONo = new string[] { "AB202201240531", "AF202201240533", "AC202201240538" };
            int WONo = 0;
            var woInterval = new List<int>
            {
                1
            };
            for (int j = 1; j <= test.Count; j++)
            {           
                temp.Add(new WorkOrderInfo(testWONo[WONo], "PMC", testCount, "2022-01-24"));
                if (j % 10 == 0)
                {
                    testCount += 5;
                    if(WONo < 2)
                    {
                        WONo += 1;
                    }
                    woInterval.Add(j);
                }
            }
            for (int i = 0; i < count; i++)
            {
                result.Add(new Quality(spc[i], "cm", CPK[i], Cp[i], Ca[i], "9000K", 2, cl, LSL.ToArray(), USL.ToArray(), test.ToArray(), rm.ToArray(), avg, mse, temp, woInterval));
                for(int j = 0; j < test.Count; j++)
                {
                    test[j] += 10M;
                }
                cl += 10;
                for(int j = 0; j < 2; j++)
                {
                    USL[j] += 10;
                    LSL[j] += 10;
                }
                avg += 10;
            }
            return new ActionResponse<List<Quality>>
            {
                Data = result
            };
        }

    }
}
