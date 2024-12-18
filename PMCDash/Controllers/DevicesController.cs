using Microsoft.AspNetCore.Mvc;
using PMCDash.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Threading.Tasks;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    public class DevicesController : BaseApiController
    {
        ConnectStr _ConnectStr = new ConnectStr();

        [Route("GroupQuery")]
        [HttpPost]
        public ActionResponse<List<DeviceGroup>> DeviceGroupQuery()
        {
            List<DeviceGroup> group = new List<DeviceGroup>();

            var SqlStr = $@"SELECT *
                            FROM {_ConnectStr.APSDB}.[dbo].[GroupList]";

            using (var conn = new SqlConnection(_ConnectStr.Local))
            {
                using (var comm = new SqlCommand(SqlStr, conn))
                {
                    if (conn.State != ConnectionState.Open)
                        conn.Open();
                    using (SqlDataReader SqlData = comm.ExecuteReader())
                    {
                        if (SqlData.HasRows)
                        {
                            int i = 1;
                            while (SqlData.Read())
                            {
                                group.Add(new DeviceGroup
                                {
                                    Id = int.Parse(SqlData["GroupID"].ToString().Trim()),
                                    Name = string.IsNullOrEmpty(SqlData["GroupName"].ToString().Trim()) ? "N/A" : SqlData["GroupName"].ToString().Trim(),
                                    Number = string.IsNullOrEmpty(SqlData["GroupNo"].ToString().Trim()) ? "N/A" : SqlData["GroupNo"].ToString().Trim(),
                                    Range = int.Parse(SqlData["Range"].ToString().Trim()),
                                    Description = string.IsNullOrEmpty(SqlData["Note"].ToString().Trim()) ? "N/A" : SqlData["Note"].ToString().Trim()
                                });
                                i++;
                            }
                        }
                    }
                }
            }
            return new ActionResponse<List<DeviceGroup>>
            {
                Data = group.OrderBy(x => x.Range).ToList()
            };
        }
    }

    public class DeviceGroup
    {
        /// <summary>
        /// 群組ID
        /// </summary>
        public int Id { get; set; }
        /// <summary>
        /// 群組順序
        /// </summary>
        public int Range { get; set; }
        /// <summary>
        /// 群組編號
        /// </summary>
        public string Number { get; set; }
        /// <summary>
        /// 群組名稱
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 群組描述
        /// </summary>
        public string Description { get; set; }

    }

    public class DeviceGroupQuery : BaseQueryData
    {
        public string Name { get; set; }

        public string Number { get; set; }
    }

    public class BaseQueryData
    {
        /// <summary>
        /// 頁面
        /// </summary>
        public int Page { get; set; } = 0;

        /// <summary>
        /// 筆數
        /// </summary>
        public int PageSize { get; set; } = 20;

        /// <summary>
        /// 排序欄位
        /// </summary>
        public string SortField { get; set; }

        /// <summary>
        /// 排序方式 1 or -1
        /// </summary>
        public int SortOrder { get; set; }

        /// <summary>
        /// 搜尋全部
        /// </summary>
        public bool All { get; set; } = false;
    }
}
