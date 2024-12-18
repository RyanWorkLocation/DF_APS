using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PMCDash.Models;
using PMCDash.Services;
namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DistributionController : BaseApiController
    {
        private readonly DeviceDistributionService _deviceDistributionService;
        public DistributionController(DeviceDistributionService deviceDistributionService)
        {
            _deviceDistributionService = deviceDistributionService;
        }

        /// <summary>
        /// 取得各廠機台狀態分布比例
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResponse<List<FactorysStatusDistribution>> Get()
        {
            var result = new List<FactorysStatusDistribution>();
            for (int i = 0; i < 5; i++)
            {
                result.Add(new FactorysStatusDistribution
                (
                    $@"FA-0{i + 1}",
                    new StatusDistribution
                    (
                        run: 52.3m,
                        idle: 27.8m,
                        alarm: 16.7m,
                        off: 3.2m
                    )
                ));
            }

            return new ActionResponse<List<FactorysStatusDistribution>>
            {
                Data = result
            };
        }

      /// <summary>
      /// 取得各產線狀態分布
      /// </summary>
      /// <param name="factory">廠區代號</param>
      /// <returns></returns>
        [HttpGet("productionline/{factory}")]
        public ActionResponse<List<ProductionLineStatusDistribution>> GetPrductionLines(string factory)
        {
            var result = new List<ProductionLineStatusDistribution>();
            for (int i = 0; i < 5; i++)
            {
                result.Add(new ProductionLineStatusDistribution
                (
                    $@"PRL-0{i + 1}",
                    new StatusDistribution
                    (
                        run: 52.3m,
                        idle: 27.8m,
                        alarm: 16.7m,
                        off: 3.2m
                    )
                ));
            }

            return new ActionResponse<List<ProductionLineStatusDistribution>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 取得全廠機台狀態分布比例
        /// </summary>
        /// <returns></returns>
        [HttpGet("all")]
        public ActionResponse<StatusDistribution> GetAllStatusDistribution()
        {
            return new ActionResponse<StatusDistribution>
            {
                Data = new StatusDistribution
                    (
                        run: 52.3m,
                        idle: 27.8m,
                        alarm: 16.7m,
                        off: 3.2m
                    )
            };
        }

    }
}
