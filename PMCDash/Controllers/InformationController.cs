using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PMCDash.Models;
namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    public class InformationController : BaseApiController
    {
        public InformationController()
        {

        }

        /// <summary>
        /// 取回所有工廠清單
        /// </summary>
        /// <returns></returns>
        //[HttpGet("Factorys")]
        //public ActionResponse<List<FactoryImformation>> GetFactory()
        //{
        //    var result = new List<FactoryImformation>();
        //    var factorynName = new string[] { "大里廠", "松竹廠", "松竹五廠", "鐮村廠", "松竹七廠" };
        //    for(int i = 0; i < 5; i++)
        //    {
        //        result.Add(new FactoryImformation($@"FA-0{i + 1}", factorynName[i]));
        //    }
        //    return new ActionResponse<List<FactoryImformation>>
        //    {
        //        Data = result
        //    };
        //}
        [HttpGet]
        public ActionResponse<FactoryDefine> Get()
        {
            var devices = new List<Device_e>();
            var productionLines = new List<ProductionLine>();
            for (int i = 0; i < 10; i++)
            {
                devices.Add(new Device_e($@"CNC-{i + 1,00}", $@"CNC-{i + 1,00}"));
                if (i == 4)
                {
                    productionLines.Add(new ProductionLine(@$"PRL-{i - 3,00}", @$"PRL-{i - 3,00}"));
                    productionLines[0].Devices = devices.ToList();
                    devices.Clear();
                }

                if (i == 9)
                {
                    productionLines.Add(new ProductionLine(@$"PRL-02", @$"PRL-02"));
                    productionLines[1].Devices = devices.ToList();
                    devices.Clear();
                }
            }
            var factorys = new List<Factory>();
            var facotorys = new Factory("FA-01", "大里廠");
            facotorys.ProductionLines = productionLines.ToList();
            factorys.Add(facotorys);
            var result = new FactoryDefine();
            result.Factorys = factorys;
            return new ActionResponse<FactoryDefine>
            {
                Data = result
            };
        }
        /// <summary>
        /// 取得特定產線名稱
        /// </summary>
        /// <param name="factory">廠區名稱 EX:FA-01</param>
        /// <returns></returns>
        [HttpGet("Productionlines/{factory}")]
        public ActionResponse<List<ProductionLineImformation>> GetProduction(string factory)
        {
            var result = new List<ProductionLineImformation>();
            var factorynName = new string[] { "WGAM", "WGCM", "WEA", "WTA", "WGPK" };
            for (int i = 0; i < 5; i++)
            {
                result.Add(new ProductionLineImformation($@"PRL-0{i}", factorynName[i]));
            }
            return new ActionResponse<List<ProductionLineImformation>>
            {
                Data = result
            };
        }

        /// <summary>
        /// 取得產線中所有的機台名稱
        /// </summary>
        /// <param name="prl">輸入廠區名稱與產線名稱</param>
        /// <returns></returns>
        [HttpPost("Machines")]
        public ActionResponse<List<MachineInformation>> GetMachine([FromBody] RequestProductionLine prl)
        {
            var result = new List<MachineInformation>();

            var status = new string[] { "RUN", "IDLE", "ALARM", "OFF" };

            for (int i = 0; i < 10; i++)
            {
                result.Add(new MachineInformation($@"CNC-{i + 1, 2:00}", status[i % 4], $@"CNC-{i + 1,2:00}"));
            }

            return new ActionResponse<List<MachineInformation>>
            {
                Data = result
            };
        }
    }
}
