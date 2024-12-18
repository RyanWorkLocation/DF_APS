using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
    public class Order
    {
        /// <summary>
        /// 訂單資料
        /// </summary>
        /// <param name="_Id">訂單編號</param>
        /// <param name="_Name">訂單名稱</param>
        /// <param name="_Type">訂單類別</param>
        /// <param name="_OrderDate">訂單訂購日期</param>
        /// <param name="_Status">訂單狀態</param>
        /// <param name="_CustumerInformation">客戶相關資訊</param>
        /// <param name="_DeliveryInformation">送貨相關資訊</param>
        /// <param name="_OrderItems">訂單內容明細</param>
        /// <param name="_SubTotal">小計</param>
        /// <param name="_Tax">稅額</param>
        /// <param name="_Discount">折扣</param>
        /// <param name="_Shipping">運費</param>
        /// <param name="_Total">總計</param>
        public Order(string _Id, string _Name, string _Type, string _OrderDate, string _Status, string _CustumerInformation, string _DeliveryInformation, List<OrderItem> _OrderItems, string _SubTotal, string _Tax, string _Discount, string _Shipping, string _Total)
        {
            Id = _Id;
            Name = _Name;
            Type = _Type;
            OrderDate = _OrderDate;
            Status = _Status;
            CustumerInformation = _CustumerInformation;
            DeliveryInformation = _DeliveryInformation;
            OrderItems = _OrderItems;
            SubTotal = _SubTotal;
            Tax = _Tax;
            Discount = _Discount;
            Shipping = _Shipping;
            Total = _Total;
        }

        /// <summary>
        /// 訂單編號
        /// </summary>
        public string Id { get; set; }
        /// <summary>
        /// 訂單名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 訂單類別
        /// </summary>
        public string Type { get; set; }
        /// <summary>
        /// 訂單訂購日期
        /// </summary>
        public string OrderDate { get; set; }
        /// <summary>
        /// 訂單狀態
        /// </summary>
        public string Status { get; set; }
        /// <summary>
        /// 客戶相關資訊
        /// </summary>
        public string CustumerInformation { get; set; }
        /// <summary>
        /// 送貨相關資訊
        /// </summary>
        public string DeliveryInformation { get; set; }
        /// <summary>
        /// 訂單內容明細
        /// </summary>
        public List<OrderItem> OrderItems { get; set; }
        /// <summary>
        /// 小計
        /// </summary>
        public string SubTotal { get; set; }
        /// <summary>
        /// 稅額
        /// </summary>
        public string Tax { get; set; }
        /// <summary>
        /// 折扣
        /// </summary>
        public string Discount { get; set; }
        /// <summary>
        /// 運費
        /// </summary>
        public string Shipping { get; set; }
        /// <summary>
        /// 總計
        /// </summary>
        public string Total { get; set; }
    }

    public class OrderItem
    {
        /// <summary>
        /// 訂單內容明細
        /// </summary>
        /// <param name="_Id">編號</param>
        /// <param name="_PartNo">物料/BOM編號</param>
        /// <param name="_Quantity">數量</param>
        /// <param name="_SubTotal">單價</param>
        /// <param name="_Total">金額</param>
        /// <param name="_Remark">備註</param>
        public OrderItem(string _Id, string _PartNo, string _Quantity, string _SubTotal, string _Total, string _Remark)
        {
            Id = _Id;
            PartNo = _PartNo;
            Quantity = _Quantity;
            SubTotal = _SubTotal;
            Total = _Total;
            Remark = _Remark;
        }

        /// <summary>
        /// 編號
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 物料/BOM編號
        /// </summary>
        public string PartNo { get; set; }

        /// <summary>
        /// 數量
        /// </summary>
        public string Quantity { get; set; }

        /// <summary>
        /// 單價
        /// </summary>
        public string SubTotal { get; set; }

        /// <summary>
        /// 金額
        /// </summary>
        public string Total { get; set; }

        /// <summary>
        /// 備註
        /// </summary>
        public string Remark { get; set; }


    }


}
