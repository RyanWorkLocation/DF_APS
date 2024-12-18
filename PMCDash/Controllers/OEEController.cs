using Microsoft.AspNetCore.Mvc;
using PMCDash.Models;
using System.Linq;
using System.Collections;
using System;
using System.Collections.Generic;
using Microsoft.AspNetCore.Authorization;
using System.Globalization;

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    public class OEEController : BaseApiController
    {
        public OEEController()
        {

        }

        /// <summary>
        /// 取得當天工廠OE
        /// </summary>
        /// <param name="rquset">廠區名稱(EX:FA-01、all(此為整公司)) 產線名稱(可忽略(Empty)若指定特定產線再填入)</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResponse<OEEOverView> Post([FromBody] RequestProductionLine rquset)
        {
            double oee = 0;
            double availbility = (7.6 / 8);
            double performance = (7.4 / 8);
            double yield = 0.973;

            oee = availbility * performance * yield;
            oee = Math.Round(oee, 3, MidpointRounding.AwayFromZero);

            return new ActionResponse<OEEOverView>
            {
                Data = new OEEOverView
                (
                    oEE: new OEERate(oee * 100, 60d),
                    availbility: new AvailbilityRate(availbility * 100, 90d),
                    performance: new PerformanceRate(performance * 100, 64d),
                    yield: new YieldRate(yield * 100, 95d),
                    delivery: new DeliveryRate(94.6d, 95d)
                )
            };

            //return new ActionResponse<OEEOverView>
            //{
            //    Data = new OEEOverView
            //    (
            //        oEE: new OEERate(80d, 60d),
            //        availbility: new AvailbilityRate(60d, 90d),
            //        performance: new PerformanceRate(60d, 64d),
            //        yield: new YieldRate(100d, 95d),
            //        delivery: new DeliveryRate(94.6d, 95d)
            //    )
            //};
        }

        /// <summary>
        /// 取固定天數的整廠OEE
        /// </summary>       
        /// <param name="days">整數值(EX:7、15....)</param> 
        /// <returns></returns>
        [HttpGet("days/{days}")]
        public ActionResponse<List<OEEOverViewHistory>> Get(int days)
        {
            var result = new List<OEEOverViewHistory>();
            var random = new Random();
            for (int i = 0; i < days; i++)
            {
                result.Add(new OEEOverViewHistory
                (
                    date: DateTime.Now.AddDays(-i).ToString("yyyy/MM/dd"),
                    oeeOverView: new OEEOverView
                    (
                        oEE: new OEERate(Math.Round(random.NextDouble() * 100, 2), 60d),
                        availbility: new AvailbilityRate(Math.Round(random.NextDouble() * 100, 2), 90d),
                        performance: new PerformanceRate(Math.Round(random.NextDouble() * 100, 2), 64d),
                        yield: new YieldRate(Math.Round(random.NextDouble() * 100, 2), 95d),
                        delivery: new DeliveryRate(Math.Round(random.NextDouble() * 100, 2), 95d)
                    )
                ));
            }
            return new ActionResponse<List<OEEOverViewHistory>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 取回良品率細節列表
        /// </summary>
        /// <returns></returns>
        [HttpGet("YieldDetails")]
        public ActionResponse<List<YiledDetails>> GetYieldRateDetails()
        {
            string[] product = new string[] { "AS-ASF0060WQR-FPOX", "AF-ASF0060PPW-FAOX", "AS-ASF0075WRQ-FPOX", "AS-ASF0070WQR-FPOX", "AS-ASF0080WQR-FPOX",
                "AK-ASF0060QQR-FPOX", "AS-ASF0100WQR-FPOX", "AT-ASF0060WQR-FPOX" ,"AP-ASF0060WQR-FPOX", "AK-ASF0060WQR-FFPS" };
            var result = new List<YiledDetails>();
            for (int i = 0; i < 10; i++)
            {
                result.Add(new YiledDetails
                (
                    product[i],
                    (i * 10) + 5.6
                ));
            }
            return new ActionResponse<List<YiledDetails>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 取回機台一周稼動率統計
        /// </summary>
        /// <param name="device">請輸入 Factroy ProductionLine Device Name</param>
        /// <returns></returns>
        [HttpPost("week")]
        public ActionResponse<WeekUtilization> GetOeeofWeek([FromBody] RequestFactory device)
        {
            var result = new List<Utilization>();
            for (int i = 1; i <= 7; i++)
            {
                result.Add(new Utilization($@"{DateTime.Now.AddDays(-i):MM/dd}({CultureInfo.CurrentCulture.DateTimeFormat.GetDayName(DateTime.Now.AddDays(-i).DayOfWeek)[2]})"
                    , 40.0 + i, 20 - i, 20 + 2 * i, 20 - 2 * i));
            }
            return new ActionResponse<WeekUtilization>
            {
                Data = new WeekUtilization(device.DeviceName, result)
            };
        }
    }
}
