using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class OEERate
    {
        public OEERate(double oee, double oeeLimit)
        {
            OEE = oee;
            OEELimit = oeeLimit;
        }

        public double OEE { get; set; }
        public double OEELimit { get; set; }
    }

    public class AvailbilityRate
    {
        public AvailbilityRate(double availbility, double availbilityLimit)
        {
            Availbility = availbility;
            AvailbilityLimit = availbilityLimit;
        }

        public double Availbility { get; set; }
        public double AvailbilityLimit { get; set; }
    }

    public class PerformanceRate
    {
        public PerformanceRate(double performance, double performanceLimit)
        {
            Performance = performance;
            PerformanceLimit = performanceLimit;
        }

        public double Performance { get; set; }
        public double PerformanceLimit { get; set; }
    }

    public class YieldRate
    {
        public YieldRate(double yield, double yieldLimit)
        {
            Yield = yield;
            YieldLimit = yieldLimit;
        }

        public double Yield { get; set; }

        public double YieldLimit { get; set; }
    }

    public class DeliveryRate
    {
        public DeliveryRate(double delivery, double deliveryLimit)
        {
            Delivery = delivery;
            DeliveryLimit = deliveryLimit;
        }

        public double Delivery { get; set; }

        public double DeliveryLimit { get; set; }
    }

    public class OEEOverViewHistory
    {
        public OEEOverViewHistory(string date, OEEOverView oeeOverView)
        {
            Date = date;
            this.oeeOverView = oeeOverView;
        }

        public string Date { get; set; }

        public OEEOverView oeeOverView { get; set; }
    }

    public class OEEOverView
    {
        public OEEOverView(OEERate oEE, AvailbilityRate availbility, PerformanceRate performance, YieldRate yield, DeliveryRate delivery)
        {
           
            OEE = oEE;
            Availbility = availbility;
            Performance = performance;
            Yield = yield;
            Delivery = delivery;
        }
        public OEERate OEE { get; set; }

        public AvailbilityRate Availbility { get; set; }

        public PerformanceRate Performance { get; set; }

        public YieldRate Yield { get; set; }

        public DeliveryRate Delivery { get; set; }
    }

    public class YiledDetails
    {
        public YiledDetails(string proudctName, double rateValue)
        {
            ProudctName = proudctName;
            RateValue = rateValue;
        }

        public string ProudctName { get; set; }

        public double RateValue { get; set; }
    }
}
