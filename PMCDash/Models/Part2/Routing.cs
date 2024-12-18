using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models.Part2
{
    public class Routing
    {
        /// <summary>
        /// 途程資料
        /// </summary>
        /// <param name="_Id">途程編號</param>
        /// <param name="_Name">途程名稱</param>
        /// <param name="_PartNo">物料編號</param>
        /// <param name="_PartName">物料名稱</param>
        /// <param name="_OPList">製程明細(有順序)</param>
        /// <param name="_CreatTime">創建時間</param>
        /// <param name="_CreatPerson">創建人員</param>
        /// <param name="_LastEditTime">最後修改時間</param>
        /// <param name="_LastEditPerson">最後修改人員</param>
        public Routing(string _Id, string _Name, string _PartNo, string _PartName, List<OP> _OPList, DateTime _CreatTime, string _CreatPerson, DateTime _LastEditTime, string _LastEditPerson)
        {
            Id = _Id;
            Name = _Name;
            PartNo = _PartNo;
            PartName = _PartName;
            OPList = _OPList;
            CreatTime = _CreatTime;
            CreatPerson = _CreatPerson;
            LastEditTime = _LastEditTime;
            LastEditPerson = _LastEditPerson;

        }

        /// <summary>
        /// 途程編號
        /// </summary>
        public string Id { get; set; }

        /// <summary>
        /// 途程名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 物料編號
        /// </summary>
        public string PartNo { get; set; }

        /// <summary>
        /// 物料名稱
        /// </summary>
        public string PartName { get; set; }

        /// <summary>
        /// 製程明細(有順序)
        /// </summary>
        public List<OP> OPList { get; set; }

        /// <summary>
        /// 創建時間
        /// </summary>
        public DateTime CreatTime { get; set; }

        /// <summary>
        /// 創建人員
        /// </summary>
        public string CreatPerson { get; set; }

        /// <summary>
        /// 最後修改時間
        /// </summary>
        public DateTime LastEditTime { get; set; }

        /// <summary>
        /// 最後修改人員
        /// </summary>
        public string LastEditPerson { get; set; }

    }
}
