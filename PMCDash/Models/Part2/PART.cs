using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
    public class PART
    {
        /// <summary>
        /// 物料資料
        /// </summary>
        /// <param name="_Id">物料編號</param>
        /// <param name="_Name">物料名稱</param>
        /// <param name="_Description">物料描述</param>
        /// <param name="_UseForMetrail">物料選用素材</param>
        /// <param name="_custom_Column01">自定義欄位1</param>
        /// <param name="_custom_Column02">自定義欄位2</param>
        /// <param name="_custom_Column03">自定義欄位3</param>
        /// <param name="_custom_Column04">自定義欄位4</param>
        /// <param name="_custom_Column05">自定義欄位5</param>
        /// <param name="_custom_Column06">自定義欄位6</param>
        /// <param name="_custom_Column07">自定義欄位7</param>
        /// <param name="_custom_Column08">自定義欄位8</param>
        /// <param name="_custom_Column09">自定義欄位9</param>
        /// <param name="_custom_Column10">自定義欄位10</param>
        /// <param name="_RoutingNo">途程資料</param>
        /// <param name="_CADUrl">設計圖檔位置</param>
        /// <param name="_UnitPrice">物料單價</param>
        public PART(string _Id, string _Name, string _Description, Metrail _UseForMetrail, CustomColum _custom_Column01, CustomColum _custom_Column02, CustomColum _custom_Column03, CustomColum _custom_Column04, CustomColum _custom_Column05, CustomColum _custom_Column06, CustomColum _custom_Column07, CustomColum _custom_Column08, CustomColum _custom_Column09, CustomColum _custom_Column10, Routing _RoutingNo, string _CADUrl, string _UnitPrice)
        {
            Id = _Id;
            Name = _Name;
            Description = _Description;
            UseForMetrail = _UseForMetrail;
            custom_Column01 = _custom_Column01;
            custom_Column02 = _custom_Column02;
            custom_Column03 = _custom_Column03;
            custom_Column04 = _custom_Column04;
            custom_Column05 = _custom_Column05;
            custom_Column06 = _custom_Column06;
            custom_Column07 = _custom_Column07;
            custom_Column08 = _custom_Column08;
            custom_Column09 = _custom_Column09;
            custom_Column10 = _custom_Column10;
            RoutingNo = _RoutingNo;
            CADUrl = _CADUrl;
            UnitPrice = _UnitPrice;
        }

        /// <summary>
        /// 物料編號
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 物料名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 物料描述
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// 物料選用素材
        /// </summary>
        public Metrail UseForMetrail { get; set; }

        /// <summary>
        /// 自定義欄位1
        /// </summary>
        public CustomColum custom_Column01 { get; set; }

        /// <summary>
        /// 自定義欄位2
        /// </summary>
        public CustomColum custom_Column02 { get; set; }

        /// <summary>
        /// 自定義欄位3
        /// </summary>
        public CustomColum custom_Column03 { get; set; }

        /// <summary>
        /// 自定義欄位4
        /// </summary>
        public CustomColum custom_Column04 { get; set; }

        /// <summary>
        /// 自定義欄位5
        /// </summary>
        public CustomColum custom_Column05 { get; set; }

        /// <summary>
        /// 自定義欄位6
        /// </summary>
        public CustomColum custom_Column06 { get; set; }

        /// <summary>
        /// 自定義欄位7
        /// </summary>
        public CustomColum custom_Column07 { get; set; }

        /// <summary>
        /// 自定義欄位8
        /// </summary>
        public CustomColum custom_Column08 { get; set; }

        /// <summary>
        /// 自定義欄位9
        /// </summary>
        public CustomColum custom_Column09 { get; set; }

        /// <summary>
        /// 自定義欄位10
        /// </summary>
        public CustomColum custom_Column10 { get; set; }

        /// <summary>
        /// 途程資料
        /// </summary>
        public Routing RoutingNo { get; set; }

        /// <summary>
        /// 設計圖檔位置
        /// </summary>
        public string CADUrl { get; set; }

        /// <summary>
        /// 物料單價
        /// </summary>
        public string UnitPrice { get; set; }




        public string CreateTime { get; set; }
        public string LastEditTime { get; set; }
    }

    public class CustomColum
    {
        /// <summary>
        /// 自訂義欄位
        /// </summary>
        /// <param name="_Title">欄位名稱</param>
        /// <param name="_Data">欄位資料</param>
        public CustomColum(string _Title,string _Data)
        {
            Title = _Title;
            data = _Data;
        }

        /// <summary>
        /// 欄位名稱
        /// </summary>
        public string Title { get; set; }
        /// <summary>
        /// 欄位資料
        /// </summary>
        public string data { get; set; }
    }

    public class Metrail
    {
        /// <summary>
        /// 素材資料
        /// </summary>
        /// <param name="_Id">素材編號</param>
        /// <param name="_Name">素材名稱</param>
        /// <param name="_Spec">素材規格</param>
        /// <param name="_Remark">素材備註</param>
        public Metrail(string _Id, string _Name, string _Spec, string _Remark)
        {
            Id = _Id;
            Name = _Name;
            Spec = _Spec;
            Remark = _Remark;
        }

        /// <summary>
        /// 素材編號
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// 素材名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 素材規格
        /// </summary>
        public string Spec { get; set; }
        /// <summary>
        /// 素材備註
        /// </summary>
        public string Remark { get; set; }
    }
}
