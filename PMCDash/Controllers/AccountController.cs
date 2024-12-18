using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PMCDash.Filters;
using PMCDash.Services;
using PMCDash.Models;
using Microsoft.AspNetCore.Authorization;
using static PMCDash.Services.AccountService;
using System.Text;
using System.Security.Cryptography;
using System.IO;

namespace PMCDash.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class AccountController : BaseApiController
    {
        private readonly AccountService _accountService;
        public AccountController(AccountService accountService)
        {
            _accountService = accountService;
        }

        /// <summary>
        /// 登入使用者
        /// </summary>
        /// <param name="authRequest">帳號密碼請洽PMC</param>
        /// <returns></returns>
        [AllowAnonymous]
        [HttpPost("login")]
        public string Login(AuthRequestcs authRequest)
        {
            return _accountService.SignIn(authRequest);        
        }

        /// <summary>
        /// 驗證Token是否還在效期內
        /// </summary>
        /// <returns></returns>
        [Authorize]
        [HttpGet("Validate")]
        public string ValidateUser()
        {
            return User.Identity.Name;
        }

        /// <summary>
        /// 取得登入者帳號、名稱、群組代號、群組名稱、群組設備
        /// </summary>
        /// <returns></returns>
        [Authorize]
        [HttpGet("UserRoleInfo")]
        public ActionResponse<UserData> UserRoleInfo()
        {
            var result = _accountService.GetUserData(User.Identity.Name);

            return new ActionResponse<UserData>
            {
                Data = result
            };
            //return _accountService.GetUserData(User.Identity.Name);
        }

        

    }
}
