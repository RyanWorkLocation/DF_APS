using Microsoft.IdentityModel.Protocols;
using Newtonsoft.Json;
using PMCDash.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace PMCDash.Services
{
    

    public class erpService
    {
        ConnectStr _ConnectStr = new ConnectStr();
        string _connectStrAPS = "";//"Data Source = 127.0.0.1; Initial Catalog = 01_APS; User ID = MES2014; Password = PMCMES;";
        public string ERPServerUrl = "http://192.168.50.102:8085";
        public string SetAccessToken = "";

        public erpService()
        {
            _connectStrAPS = _ConnectStr.Local;
            ERPServerUrl = _ConnectStr.ERPurl;
            SetAccessToken = gettoken(ERPServerUrl);
        }

        private string gettoken(string baseurl)
        {//  http://localhost:8085/token
            string token = "Bearer 0gS-NbhtNHhifEclY0Jji9GPm0TIFtjWKRhXASAhobAPsnEE2f47K0w52TEam5YgHF9IDn3gerSu2JaESSTbH3z-PxRt2AU_WdjgL2avKTKzowHcnPZKZ1dFaMd-2o9leDr9qhLcoYSRpnbXqeAPkeV4SyY8U08zE6Q_bxb-kgwM57Gmg6Lo4oql2CZ-nBUg0tuWczCMxVkY3gFbkuH5p7lUFzWVTykIRcPZa4rbbar9Y6UYCFmtXweInRQQATUELDkdiRDUXHRSzbuT1OBIyft1s0GsuusJoAhTMS3P-MzEKe2W5eZLk_EF0wX3nuhuTB-o7eCXXGUjDQKVqm1dJLeiWfCvnoi2GtOSsfnosstp78Iskgm3FHroLEYBvjeTxs_xKBK4FnyuzHDo99l5tvKXx4J5qMV9Sa2UO-NeTIdeWB4_tcEkMULsYGbD5CjT6ievA2K5Q6KLYU8RjNgGQZMXlAjtybL5gLPK1v0m0iW5HYEeI61jcdWa4ARTUp9raPzYqcBdiRXrqcPU1zcYu93zVxl80srRq3tHQ3uFcRbqmvnQjn5QP5zibdtiZiPs4iLfRw";
            string result = RequestWebAPI(@"userName=admin%40example.com&password=123456&grant_type=password", ERPServerUrl + "/token", token);
            Accunt.Root data = JsonConvert.DeserializeObject<Accunt.Root>(result);

            return "Bearer " + data.access_token;
        }

        public string RequestWebAPI(string sendData, string url, string SetAccessToken)
        {
            try
            {
                string backMsg = "";
                WebClient client = new WebClient();
                client.Headers.Add("Authorization", SetAccessToken);
                client.Headers.Add("Content-Type", "application/json;charset=utf-8");
                backMsg = client.UploadString(url, "POST", sendData);
                return backMsg;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }

        public string RequestWebAPI_PUT(string sendData, string url, string SetAccessToken)
        {
            try
            {
                string backMsg = "";
                WebClient client = new WebClient();
                client.Headers.Add("Authorization", SetAccessToken);
                client.Headers.Add("Content-Type", "application/json;charset=utf-8");
                backMsg = client.UploadString(url, "PUT", sendData);
                return backMsg;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }

    class Accunt
    {
        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
        public class Root
        {
            public string access_token { get; set; }
            public string token_type { get; set; }
            public int expires_in { get; set; }
            public string refresh_token { get; set; }
            public string userName { get; set; }

            [JsonProperty(".issued")]
            public string issued { get; set; }

            [JsonProperty(".expires")]
            public string expires { get; set; }
        }
    }

    class JobOrder_Query
    {
        // Root myDeserializedClass = JsonConvert.DeserializeObject<Root>(myJsonResponse);
        public class Product
        {
            public int Id { get; set; }
            public string Number { get; set; }
            public string Name { get; set; }
        }

        public class Process
        {
            public int Id { get; set; }
            public string Name { get; set; }
            public string Number { get; set; }
        }

        public class DrawMaterial
        {
            public object Id { get; set; }
            public object Number { get; set; }
        }

        public class Root
        {
            public int Id { get; set; }
            public string Number { get; set; }
            public int Type { get; set; }
            public int Status { get; set; }
            public float Quantity { get; set; }
            public float RequiredQuantity { get; set; }
            public DateTime LeadTime { get; set; }
            public DateTime EstimatedDate { get; set; }
            public Product Product { get; set; }
            public Process Process { get; set; }
            public DrawMaterial DrawMaterial { get; set; }
            public DateTime CreateDateTime { get; set; }
            public DateTime UpdateDateTime { get; set; }
            public string CreateEmployeeStr { get; set; }
            public string UpdateEmployeeStr { get; set; }
            public float OrderId { get; set; }
            public string OrderNumber { get; set; }
            public float OrderRequiredQuantity { get; set; }
            public DateTime OrderDeliveryDate { get; set; }
        }
    }
}
