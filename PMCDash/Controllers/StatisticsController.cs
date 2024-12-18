using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PMCDash.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Text.Json;
using PMCDash.Services;
namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    public class StatisticsController : BaseApiController
    {
        private readonly AlarmService _alarmService;
        public StatisticsController(AlarmService alarmService)
        {
            _alarmService = alarmService;
        }

        /// <summary>
        /// 取得各廠狀態統計資料
        /// </summary>
        /// <returns></returns>
        public ActionResponse<List<FactoryStatistics>> GetStaticsByFactory()
        {
            var result = new List<FactoryStatistics>();
            for (int i = 0; i < 5; i++)
            {
                result.Add(new FactoryStatistics
                (
                    $@"FA-0{i + 1}",
                    new StatusStatistics
                    (
                        run: 52,
                        idle: 28,
                        alarm: 17,
                        off: 3
                    )
                ));
            }
            return new ActionResponse<List<FactoryStatistics>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 取得特定廠區中的產線統計資料
        /// </summary>
        /// <param name="factory">廠區名稱</param>
        /// <returns></returns>
        [HttpGet("productionline/{factory}")]
        public ActionResponse<FactoryStatisticsImformation> GetStaticsByProductionLine(string factory)
        {
            var result = new List<ProductionLineStatistics>();
            for (int i = 0; i < 5; i++)
            {
                result.Add(new ProductionLineStatistics
                (
                    $@"PRL-0{i + 1}",
                    new StatusStatistics
                    (
                        run: 52,
                        idle: 28,
                        alarm: 17,
                        off: 3
                    )
                ));
            }
            return new ActionResponse<FactoryStatisticsImformation>
            {
                Data = new FactoryStatisticsImformation(factory, result)
            };
        }

        /// <summary>
        /// 取的特定產線中的機台狀態統計資料
        /// </summary>       
        /// <returns></returns>
        [HttpPost("status")]
        public ActionResponse<ProductionLineMachineImformation> GetMachineStatus([FromBody] RequestFactory prl)
        {
            var result = new List<MachineStatus>();
            var status = new string[] { "RUN", "IDLE", "ALARM", "OFF" };
            for (int i = 0; i < 20; i++)
            {
                result.Add(new MachineStatus
                (
                    $@"CNC-{i + 1,2:00}",
                    status[(i + 1) % 4]
                ));
            }
            return new ActionResponse<ProductionLineMachineImformation>
            {
                Data = new ProductionLineMachineImformation(result)
            };
        }

        /// <summary>
        /// 取得TOP 10 異常訊息累計資料
        /// </summary>
        /// <param name="request">廠區名稱 EX: FA-05、all(整廠) 產線名稱 EX: 空白、PR-01</param>
        /// <returns></returns>
        [HttpPost("alarm")]
        public ActionResponse<List<AlarmStatistics>> GetAlarm([FromBody] RequestFactory request)
        {
            return new ActionResponse<List<AlarmStatistics>>
            {
                Data = _alarmService.GetAlarm(request)
            };
        }

        /// <summary>
        /// 取得停機次數統計
        /// </summary>
        /// <param name="request">廠區名稱 EX: FA-05、all(整廠) 產線名稱 EX: 空白、PR-01</param>
        /// <returns></returns>
        [HttpPost("stop")]
        public ActionResponse<List<StopStatistics>> GetStop([FromBody] RequestFactory request)
        {
            var result = new List<StopStatistics>();
            var random = new Random(Guid.NewGuid().GetHashCode());
            for (int i = 0; i < 10; i++)
            {
                result.Add(new StopStatistics($@"CNC-{i,2:00}", random.Next(0, 150)));
            }
            return new ActionResponse<List<StopStatistics>>
            {
                Data = result
            };
        }

    }
}
