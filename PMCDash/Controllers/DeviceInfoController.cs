using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PMCDash.Models;
namespace PMCDash.Controllers
{

    [Route("api/[controller]")]
    public class DeviceInfoController : BaseApiController
    {
        public DeviceInfoController()
        {

        }

        [HttpPost]
        public ActionResponse<OperationInfo> Post([FromBody] RequestFactory device)
        {
            return new ActionResponse<OperationInfo>
            {
                Data = new OperationInfo(
                    utilizationRate: 85.9d, status: "RUN", 
                    productionProgress: 67.8d,
                    @"-",
                    workorderInfo: new WorkOrderDetailInformation(orderNo: $@"10411110002", oPNo: 66, opName: "車牙",
                    productNo: $@"11110001", requireCount: 3000, currentCount:1500, dueDate: DateTime.Parse("2021-12-03")))
            };
        }
    }
}
