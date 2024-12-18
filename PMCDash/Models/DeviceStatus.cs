using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
namespace PMCDash.Models
{
    public class DeviceStatus
    {
        public DeviceStatus(string deviceName, string status)
        {
            DeviceName = deviceName;
            Status = status;
        }

        public string DeviceName { get; set; }

        public string Status { get; set; }
    }

    public class DailyDeviceStatus
    {
        /// <summary>
        /// 顯示時間區間
        /// 【ST:開始時間、ET:結束時間】
        /// </summary>
        public DeviceTimeInterval Interval { get; set; }

        /// <summary>
        /// 機台狀態與時間
        /// </summary>
        public List<devicestatus> devicestatus { get; set; }

    }

    public class devicestatus
    {

        /// <summary>
        /// 狀態:【0:關機:OFF、1:運轉:RUN、2:待機:IDLE、3:警報:ALARM】
        /// </summary>
        public string Devicesattusid { get; set; }
        /// <summary>
        /// 狀態開始時間
        /// </summary>
        public string StartTime { get; set; }
        /// <summary>
        /// 狀態結束時間
        /// </summary>
        public string EndTime { get; set; }
    }


    public class DeviceTimeInterval
    {
        /// <summary>
        /// 區間開始時間
        /// </summary>
        public string ST { get; set; }

        /// <summary>
        /// 區間結束時間
        /// </summary>
        public string ET { get; set; }

    }
}
