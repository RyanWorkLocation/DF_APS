using System;

namespace PMCDash.Models
{
    public class LocalMachineSeq
    {
        public string OrderID { get; set; }
        public double OPID { get; set; }
        public TimeSpan Duration { get; set; }
        public DateTime PredictTime { get; set; }
        public int PartCount { get; set; }
        public string Maktx { get; set; }
        public string WorkGroup { get; set; }
        public int EachMachineSeq { get; set; }
    }
}