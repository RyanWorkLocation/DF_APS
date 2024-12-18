using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PMCDash.Models;
namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LocationController : BaseApiController
    {
        public LocationController()
        {

        }

        /// <summary>
        /// 取回圖檔與座標描述
        /// </summary>
        /// <param name="request">輸入需要哪個廠區或產線瀏覽圖(產線可空白廠區不行)</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResponse<ImageDefine> Post(RequestProductionLine request)
        {
            var choice = string.IsNullOrEmpty(request.FactoryName) ? "404" :  
                string.IsNullOrEmpty(request.ProductionName) ? "factory" :"productionline";
            switch (choice)
            {
                case "factory":
                    return new ActionResponse<ImageDefine>
                    {
                        Data = new ImageDefine(imageUrl: "/Images/Factory.png",
                        imageInfos: new List<ImageInfo>
                        {
                            new ImageInfo("大里廠產線一號", 65.0d, 50.0d),
                        })
                    };              
                case "productionline":
                    return new ActionResponse<ImageDefine>
                    {
                        Data = new ImageDefine(imageUrl: "/Images/ProductionLine.png",
                        imageInfos: new List<ImageInfo>
                        {
                            new ImageInfo("AA-10", 30.61d, 21.48d),
                            new ImageInfo("AA-11", 35.61d, 32.48d),
                            new ImageInfo("AA-12", 41.56d, 21.48d),
                            new ImageInfo("AA-13", 45.00d, 32.48d),
                            new ImageInfo("AA-14", 50.50d, 21.48d),
                            new ImageInfo("AA-15", 53.91d, 32.48d),

                            new ImageInfo("AA-16", 60.26d, 18.48d),
                            new ImageInfo("AA-17", 60.26d, 30.48d),
                            new ImageInfo("AA-18", 67.61d, 18.48d),
                            new ImageInfo("AA-19", 67.61d, 30.48d),
                            new ImageInfo("AA-20", 74.96d, 18.48d),
                            new ImageInfo("AA-21", 74.96d, 30.48d)
                        })
                    };
                default:
                    return new ActionResponse<ImageDefine>
                    {
                        Data = new ImageDefine(imageUrl: "Not Find",
                        imageInfos: new List<ImageInfo>
                        {
                        })
                    };
            }
            
           
        }

    }
}
