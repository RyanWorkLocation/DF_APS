using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
    public class BOM
    {
        /// <summary>
        /// BOM編號
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// BOM名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// BOM規格描述
        /// </summary>
        public string Spec { get; set; }

        /// <summary>
        /// 單位
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// 數量
        /// </summary>
        public string Quantity { get; set; }

        /// <summary>
        /// 版本
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// 審核狀態
        /// </summary>
        public string Status { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }


        /// <summary>
        /// 物料集合
        /// </summary>
        public List<PART> BOMItem { get; set; }

        /// <summary>
        /// 建檔時間
        /// </summary>
        public string CreateTime { get; set; }

        /// <summary>
        /// 建檔人員
        /// </summary>
        public string CreatePerson { get; set; }

        /// <summary>
        /// 最後更新時間
        /// </summary>
        public string LastEditTime { get; set; }

        /// <summary>
        /// 最後更新人員
        /// </summary>
        public string LastEditPerson { get; set; }

    }
}
