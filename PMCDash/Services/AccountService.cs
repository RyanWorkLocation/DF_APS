using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Claims;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using PMCDash.Models;
using PMCDash.Helper;
using System.Data.SqlClient;
using System.Data;
using System.Security.Cryptography;
using System.Text;
using System.IO;

namespace PMCDash.Services
{
    public class AccountService
    {
        private readonly JwtProviderHelper _jwtProvider;

        public AccountService(JwtProviderHelper jwtProvider)
        {
            _jwtProvider = jwtProvider;
        }

        //DPI
        //soco
        //SkyMarsDB


        ConnectStr _ConnectStr = new ConnectStr();
        //資料庫連線
        //private readonly string _ConnectStr.Local = @"Data Source = 127.0.0.1; Initial Catalog = soco; User ID = MES2014; Password = PMCMES;";

        public string SignIn(AuthRequestcs authRequestcs)
        {
            ////城市刀具掃描登入用
            //if (authRequestcs.UserName == "" && authRequestcs.Password.Split('/').Length >= 2)
            //{
            //    var actAccount = authRequestcs.Password.Split('/');
            //    authRequestcs.UserName = actAccount[0];
            //    authRequestcs.Password = actAccount[1];
            //}


            string sqlStr = @$"SELECT top 1 user_id, user_account, user_name, usergroup_id
                              FROM {_ConnectStr.AccuntDB}.[dbo].[User]
                              where user_account = @account and user_password = @password";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@account", authRequestcs.UserName);
                    if (!string.IsNullOrEmpty(authRequestcs.Password))
                    {
                        comm.Parameters.AddWithValue("@password", PasswordUtility.SHA512Encryptor(authRequestcs.Password));
                    }
                    //comm.Parameters.AddWithValue("@password", PasswordUtility.SHA512Encryptor(PasswordUtility.aesDecryptBase64(authRequestcs.Password,"77974590")));

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            //sqldata.Read();
                            //user = new UserData
                            //{
                            //    JwtToken = _jwtProvider.CreateJwtToken(authRequestcs),
                            //    User_Id = Convert.ToInt64(sqldata["user_id"]),
                            //    Account = sqldata["user_account"].ToString(),
                            //    Name = sqldata["user_name"].ToString(),
                            //    Role = Convert.ToInt64(sqldata["usergroup_id"])
                            //};
                            //return user;
                            return _jwtProvider.CreateJwtToken(authRequestcs);
                        }
                    }
                }
            }
            //if (authRequestcs.UserName == "admin" && authRequestcs.Password == "1234")
            //    return _jwtProvider.CreateJwtToken(authRequestcs);
            return string.Empty;
        }

        public UserData GetUserData(string account)
        {
            UserData user = new UserData();

            string sqlStr = @$"SELECT aa.user_id,aa.user_account,aa.user_name,aa.usergroup_id,bb.usergroup_name,bb.DeviceGroup FROM (
                                SELECT user_id, user_account, user_name, usergroup_id
                                FROM {_ConnectStr.AccuntDB}.[dbo].[User])as aa
                                LEFT JOIN {_ConnectStr.AccuntDB}.[dbo].[Units] as bb
                                ON aa.usergroup_id=bb.usergroup_id
                                WHERE aa.user_account=@account";
            using (SqlConnection conn = new SqlConnection(_ConnectStr.Local))
            {
                if (conn.State != ConnectionState.Open)
                    conn.Open();
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@account", account);

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            sqldata.Read();
                            user.User_Id = Convert.ToInt64(sqldata["user_id"]);
                            user.EmpolyeeAccount = sqldata["user_account"].ToString().Trim();
                            user.EmpolyeeName = sqldata["user_name"].ToString().Trim();
                            user.GroupId = Convert.ToInt64(sqldata["usergroup_id"]);
                            user.GroupName = sqldata["usergroup_name"].ToString().Trim();
                            user.DeviceGroupId = sqldata["DeviceGroup"].ToString().Trim();
                        }
                    }
                }
                sqlStr = @$"SELECT * FROM {_ConnectStr.APSDB}.[dbo].[Device] as a
                            left join　{_ConnectStr.AccuntDB}.[dbo].[Units] as b on a.GroupName = b.DeviceGroup
                            left join {_ConnectStr.AccuntDB}.[dbo].[User] as c on b.usergroup_id=c.usergroup_id
                            where c.user_id = @userid";
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@userid", user.User_Id);

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            List<GroupDevice> GroupDevices = new List<GroupDevice>();
                            while (sqldata.Read())
                            {
                                GroupDevices.Add(new GroupDevice
                                {
                                    ID = Convert.ToInt64(sqldata["ID"]),
                                    MachineNmae = sqldata["MachineName"].ToString().Replace("\n", string.Empty),
                                    remark = sqldata["remark"].ToString(),
                                    GroupName = sqldata["GroupName"].ToString(),
                                    CommonName = sqldata["CommonName"].ToString()
                                });
                            }
                            user.GroupDevices = GroupDevices;
                        }
                    }
                }

                sqlStr = @$"select a.*, b.*,b.FuncName,  b.DetailSet, a.ViewRight, a.AuditRight, a.ModifyRight, a.DeleteRight , 
                            Used = case when (a.FuncName = b.FuncName) then 'True' else 'False' end from  {_ConnectStr.AccuntDB}.[dbo].Functions b  left join {_ConnectStr.AccuntDB}.[dbo].GroupRights as a on a.GroupSeq =
                            (select TOP(1) gm.GroupSeq FROM {_ConnectStr.AccuntDB}.[dbo].GroupMembers as gm Where gm.MemberSeqNo = @userid)
                            and b.Status = '啟用' and a.FuncName = b.FuncName
                            order by Belong";
                using (SqlCommand comm = new SqlCommand(sqlStr, conn))
                {
                    comm.Parameters.AddWithValue("@userid", user.User_Id);

                    using (SqlDataReader sqldata = comm.ExecuteReader())
                    {
                        if (sqldata.HasRows)
                        {
                            List<FunctionAccess> FunctionAccess = new List<FunctionAccess>();
                            while (sqldata.Read())
                            {
                                bool used = Convert.ToBoolean(sqldata["Used"]);
                                if (used)
                                {
                                    FunctionAccess.Add(new FunctionAccess
                                    {
                                        FunctionName = sqldata["FuncName"].ToString(),
                                        ViewRight = sqldata["ViewRight"].ToString(),
                                        ModifyRight = sqldata["ModifyRight"].ToString(),
                                        DeleteRight = sqldata["DeleteRight"].ToString(),
                                        AuditRight = sqldata["AuditRight"].ToString(),
                                        Used = sqldata["Used"].ToString()
                                    });
                                }
                                
                            }
                            user.FunctionAccess = FunctionAccess;
                            return user;
                        }
                    }
                }

                return user;
            }
        }

        public class UserData
        {
            /// <summary>
            /// 使用者id
            /// </summary>
            public Int64 User_Id { get; set; }
            /// <summary>
            /// 使用者帳號
            /// </summary>
            public string EmpolyeeAccount { get; set; }
            /// <summary>
            /// 使用者姓名
            /// </summary>
            public string EmpolyeeName { get; set; }
            /// <summary>
            /// 所屬群組代號
            /// </summary>
            public Int64 GroupId { get; set; }
            /// <summary>
            /// 所屬群組名稱
            /// </summary>
            public string GroupName { get; set; }
            /// <summary>
            /// 所屬群組編號
            /// </summary>
            public string DeviceGroupId { get; set; }
            /// <summary>
            /// 所屬群組之機台
            /// </summary>
            public List<GroupDevice> GroupDevices { get; set; }
            /// <summary>
            /// 使用者權限
            /// </summary>
            public List<FunctionAccess> FunctionAccess { get; set; }
        }



        public class GroupDevice
        {
            public Int64 ID { get; set; }
            public string MachineNmae { get; set; }
            public string remark { get; set; }
            public string GroupName { get; set; }
            public string CommonName { get; set; }
        }
    }
}
