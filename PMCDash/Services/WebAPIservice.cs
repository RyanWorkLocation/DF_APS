using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace PMCDash.Services
{
    public class WebAPIservice
    {
        public static string gettoken(string baseurl)
        {
            string token = "Bearer 0gS-NbhtNHhifEclY0Jji9GPm0TIFtjWKRhXASAhobAPsnEE2f47K0w52TEam5YgHF9IDn3gerSu2JaESSTbH3z-PxRt2AU_WdjgL2avKTKzowHcnPZKZ1dFaMd-2o9leDr9qhLcoYSRpnbXqeAPkeV4SyY8U08zE6Q_bxb-kgwM57Gmg6Lo4oql2CZ-nBUg0tuWczCMxVkY3gFbkuH5p7lUFzWVTykIRcPZa4rbbar9Y6UYCFmtXweInRQQATUELDkdiRDUXHRSzbuT1OBIyft1s0GsuusJoAhTMS3P-MzEKe2W5eZLk_EF0wX3nuhuTB-o7eCXXGUjDQKVqm1dJLeiWfCvnoi2GtOSsfnosstp78Iskgm3FHroLEYBvjeTxs_xKBK4FnyuzHDo99l5tvKXx4J5qMV9Sa2UO-NeTIdeWB4_tcEkMULsYGbD5CjT6ievA2K5Q6KLYU8RjNgGQZMXlAjtybL5gLPK1v0m0iW5HYEeI61jcdWa4ARTUp9raPzYqcBdiRXrqcPU1zcYu93zVxl80srRq3tHQ3uFcRbqmvnQjn5QP5zibdtiZiPs4iLfRw";
            string result = RequestWebAPI(@"userName=admin%40example.com&password=123456&grant_type=password", baseurl + "/token", token);
            Accunt.Root data = JsonConvert.DeserializeObject<Accunt.Root>(result);

            return "Bearer " + data.access_token;
        }

        //取的WEB API【POST】
        public static string RequestWebAPI(string sendData, string url, string SetAccessToken)
        {
            try
            {
                //Thread.Sleep(100);
                string backMsg = "";
                using (var client = new WebClient())
                {
                    client.Headers.Add("Authorization", SetAccessToken);
                    client.Headers.Add("Content-Type", "application/json;charset=utf-8");
                    backMsg = client.UploadString(url, "POST", sendData);
                }
                return backMsg;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //取的WEB API【PPUT】
        public static void RequestWebAPI_PUT(string sendData, string url, string SetAccessToken)
        {
            try
            {
                // 建立 WebClient
                string backMsg = "";
                using (var client = new WebClient())
                {
                    // 指定 WebClient 編碼
                    client.Encoding = Encoding.UTF8;
                    // 指定 WebClient 的 Content-Type header
                    client.Headers.Add(HttpRequestHeader.ContentType, "application/json");
                    // 指定 WebClient 的 authorization header
                    client.Headers.Add("authorization", SetAccessToken);
                    // 執行 PUT 動作
                    var result = client.UploadString(url, "PUT", sendData);
                    // linqpad 將 post 結果輸出
                    //result.Dump();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        //取的WEB API【GET】
        public static string RequestWebAPI_get(string url, string SetAccessToken)
        {
            try
            {
                //Thread.Sleep(100);
                string backMsg = "";
                using (var client = new WebClient())
                {
                    client.Headers.Add("Authorization", SetAccessToken);
                    client.Headers.Add("Content-Type", "application/json;charset=utf-8");
                    backMsg = client.DownloadString(url);
                }
                return backMsg;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        //取的WEB API圖片用【GET】
        public static byte[] RequestWebAPI_getIMG(string url, string SetAccessToken)
        {
            try
            {
                //Thread.Sleep(100);
                byte[] backMsg = null;
                using (var client = new WebClient())
                {
                    client.Headers.Add("Authorization", SetAccessToken);
                    client.Headers.Add("Content-Type", "application/json;charset=utf-8");
                    backMsg = client.DownloadData(url);
                }
                return backMsg;

            }
            catch
            {
                return null;
            }
        }
    }
}
