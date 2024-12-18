using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class measurervalueQuery
    {
        /// <summary>
        /// 量測設備編號
        /// </summary>
        public string id { get; set; }
        /// <summary>
        /// 開始量測時間點 ex:2022-04-26 15:04:14
        /// </summary>
        public string time { get; set; }
    }
}
