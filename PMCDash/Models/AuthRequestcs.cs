using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class AuthRequestcs
    {
        public string UserName { get; set; }

        public string Password { get; set; }
    }

    public class FunctionAccess
    {
        /// <summary>
        /// 可使用功能
        /// </summary>
        public string FunctionName { get; set; }
        /// <summary>
        /// 檢視權限
        /// </summary>
        public string ViewRight { get; set; }
        /// <summary>
        /// 修改權限
        /// </summary>
        public string ModifyRight { get; set; }
        /// <summary>
        /// 刪除權限
        /// </summary>
        public string DeleteRight { get; set; }
        /// <summary>
        /// 審核權限
        /// </summary>
        public string AuditRight { get; set; }
        /// <summary>
        /// 啟用
        /// </summary>
        public string Used { get; set; }
    }
}
