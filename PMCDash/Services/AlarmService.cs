using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using PMCDash.Models;
namespace PMCDash.Services
{
    public class AlarmService
    {
        public AlarmService()
        {

        }

        public List<AlarmStatistics> GetAlarm(object requst)
        {
            switch (requst)
            {
                case ActionRequest<Factory> req:
                    break;
                case ActionRequest<RequestFactory> req:
                    break;
            }

            var result = new List<AlarmStatistics>();
            for (int i = 0; i < 10; i++)
            {
                result.Add(new AlarmStatistics
                (
                    alarmMSg: $@"Alarm0059{i,2:00}",
                    times: (i + 7) * 4,
                    totalMin: (i + 13) * 10,
                    display: new TimeSpan(0, (i + 13) * 10, 0).ToString("c")
                ));
            }
            return result;
        }

    }
}
