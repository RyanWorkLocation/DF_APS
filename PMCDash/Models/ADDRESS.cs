using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class Address
    {
        /// <summary>
        /// 地址組成：
        /// 1.郵遞區號: 3~5碼數字
        /// 2.縣市： xx 縣/市
        /// 3.鄉鎮市區：xx 鄉/鎮/市/區
        /// 4.其他：鄉鎮市區以後的部分
        /// 規則：開頭一定要是3或5個數字的郵遞區號，如果不是，解析不會出錯，但ZipCode為空
        /// 地址一定要有XX縣/市 + XX鄉/鎮/市/區 + 其他
        /// </summary>
        /// <param name="address"></param>
        public Address(string address)
        {
            this.OrginalAddress = address;
            this.ParseByRegex(address);
        }

        /// <summary>
        /// 縣市
        /// </summary>
        public string City { get; set; }

        /// <summary>
        /// 鄉鎮市區
        /// </summary>
        public string District { get; set; }

        /// <summary>
        /// 是否符合pattern規範
        /// </summary>
        public bool IsParseSuccessed { get; set; }

        /// <summary>
        /// 原始傳入的地址
        /// </summary>
        public string OrginalAddress { get; private set; }

        /// <summary>
        /// 鄉鎮市區之後的地址
        /// </summary>
        public string Others { get; set; }

        /// <summary>
        /// 郵遞區號
        /// </summary>
        public string ZipCode { get; set; }

        /// <summary>
        /// 組成完整的地址
        /// </summary>
        /// <returns>完整的地址</returns>
        public override string ToString()
        {
            var result = string.Format("{0}{1}{2}{3}", this.ZipCode, this.City, this.District, this.Others);
            return result;
        }

        private void ParseByRegex(string address)
        {
            var pattern = @"(?<zipcode>(^\d{5}|^\d{3})?)(?<city>\D+[縣市])(?<district>\D+?(市區|鎮區|鎮市|[鄉鎮市區]))(?<others>.+)";
            Match match = Regex.Match(address, pattern);

            if (match.Success)
            {
                this.IsParseSuccessed = true;

                this.ZipCode = match.Groups["zipcode"].ToString();
                this.City = match.Groups["city"].ToString();
                this.District = match.Groups["district"].ToString();
                this.Others = match.Groups["others"].ToString();
            }
        }
    }
}
