using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.DTO
{
    public class ScheduleDto
    {
        public ScheduleDto(string orderID, int oPID, string workMachine, string startTime, string endTime, int duration)
        {
            OrderID = orderID;
            OPID = oPID;
            WorkMachine = workMachine;
            StartTime = Convert.ToDateTime(startTime);
            EndTime = Convert.ToDateTime(endTime);
            Duration = new TimeSpan(0, duration, 0);
        }

        public string OrderID { get; set; }

        public int OPID { get; set; }

        public string WorkMachine { get; set; }

        public DateTime StartTime { get; set; }

        public DateTime EndTime { get; set; }

        public TimeSpan Duration { get; set; }
    }

    public class ScheduleChangeDto
    {
        public string OrderID { get; set; }

        public int OPID { get; set; }

        public string WorkMachine { get; set; }

        public DateTime StartTime { get; set; }

        public DateTime EndTime { get; set; }
    }
}
