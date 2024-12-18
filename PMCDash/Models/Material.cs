using System;
namespace PMCDash.Models
{
    public class Material
    {
        //public Material(string no, string processNo, string cPKEvaluation, string processName, string lastProduction)
        //{
        //    No = no;
        //    ProcessNo = processNo;
        //    CPKEvaluation = cPKEvaluation;
        //    ProcessName = processName;
        //    LastProduction = lastProduction;
        //}

        /// <summary>
        /// 物料編號
        /// </summary>
        public string No { get; set; }

        /// <summary>
        /// 製程編號
        /// </summary>
        public string ProcessNo { get; set; }

        /// <summary>
        /// CPK評價
        /// </summary>
        public string CPKEvaluation { get; set; }

        /// <summary>
        /// 製程名稱
        /// </summary>
        public string ProcessName { get; set; }
        
        /// <summary>
        /// 最後一個生產日
        /// </summary>
        public string LastProduction { get; set; }

    }
}
