using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Services
{
    public class RestTime
    {
        private int HoursPerDay { get; set; }
        private int StartHour { get; set; }

        /// <summary>
        /// 預設8點開始上班，上班9小時
        /// </summary>
        public RestTime()
        {
            HoursPerDay = 9;
            StartHour = 8;
        }

        /// <summary>
        /// 設定上班工時與開始上班時間
        /// </summary>
        /// <param name="startHour">幾點開始上班</param>
        /// <param name="hoursPerDay">上班幾小時</param>
        public void SetWorkTimeRang(int startHour, int hoursPerDay)
        {
            StartHour = startHour;
            HoursPerDay = hoursPerDay;
        }

        /// <summary>
        /// 計算補償休息日完工時間
        /// </summary>
        /// <param name="PostST">開始時間</param>
        /// <param name="Duration">花費時間</param>
        /// <returns></returns>
        public DateTime restTimecheck(DateTime PostST, TimeSpan Duration)
        {
            //if (Duration > new TimeSpan(1, 00, 00, 00))
            //{
            //    var days = (int)Duration.TotalDays;
            //    TimeSpan resttime = new TimeSpan((int)(16 * days), 00, 00);
            //    Duration = Duration.Subtract(resttime);
            //}
            int hoursPerDay = HoursPerDay;
            int startHour = StartHour;
            // Don't start counting hours until start time is during working hours
            if (PostST.TimeOfDay.TotalHours > startHour + hoursPerDay)
                PostST = PostST.Date.AddDays(1).AddHours(startHour);
            if (PostST.TimeOfDay.TotalHours < startHour)
                PostST = PostST.Date.AddHours(startHour);
            if (PostST.DayOfWeek == DayOfWeek.Saturday)
                PostST.AddDays(2);
            else if (PostST.DayOfWeek == DayOfWeek.Sunday)
                PostST.AddDays(1);
            // Calculate how much working time already passed on the first day
            TimeSpan firstDayOffset = PostST.TimeOfDay.Subtract(TimeSpan.FromHours(startHour));
            // Calculate number of whole days to add
            var aaa = Duration.Add(firstDayOffset).TotalHours;
            int wholeDays = (int)(Duration.Add(firstDayOffset).TotalHours / hoursPerDay);
            // How many hours off the specified offset does this many whole days consume?
            TimeSpan wholeDaysHours = TimeSpan.FromHours(wholeDays * hoursPerDay);
            // Calculate the final time of day based on the number of whole days spanned and the specified offset
            TimeSpan remainder = Duration - wholeDaysHours;
            // How far into the week is the starting date?
            int weekOffset = ((int)(PostST.DayOfWeek + 7) - (int)DayOfWeek.Monday) % 7;
            // How many weekends are spanned?
            int weekends = (int)((wholeDays + weekOffset) / 5);
            // Calculate the final result using all the above calculated values
            var ans = PostST.AddDays(wholeDays + weekends * 2).Add(remainder);
            return PostST.AddDays(wholeDays + weekends * 2).Add(remainder);
        }

        public DateTime hoildaycheck(DateTime startTime)
        {
            DateTime result = startTime;

            if (!m_IsWorkingDay(startTime))
            {
                if (startTime > DateTime.Parse(startTime.ToShortDateString() + " 08:00"))
                {
                    if (startTime.DayOfWeek == DayOfWeek.Saturday)
                        result = DateTime.Parse(startTime.AddDays(2).ToShortDateString() + " 08:00");
                    else if (startTime.DayOfWeek == DayOfWeek.Friday)
                        result = DateTime.Parse(startTime.AddDays(3).ToShortDateString() + " 08:00");
                    else
                        result = DateTime.Parse(startTime.AddDays(1).ToShortDateString() + " 08:00");
                }
                else
                    result = DateTime.Parse(startTime.ToShortDateString() + " 08:00");
            }
            else
            {
                var s = startTime.DayOfWeek;
            }

            return result;
        }

        private string _strWorkingDayAM = "08:00";//工作时间上午08:00
        private string _strWorkingDayPM = "17:00";
        private string _strRestDay = "6,7";//周几休息日 周六周日为 6,7

        private TimeSpan dspWorkingDayAM;//工作时间上午08:00
        private TimeSpan dspWorkingDayPM;

        private string m_GetWeekNow(DateTime date)
        {
            string strWeek = date.DayOfWeek.ToString();
            switch (strWeek)
            {
                case "Monday":
                    return "1";
                case "Tuesday":
                    return "2";
                case "Wednesday":
                    return "3";
                case "Thursday":
                    return "4";
                case "Friday":
                    return "5";
                case "Saturday":
                    return "6";
                case "Sunday":
                    return "7";
            }
            return "0";
        }


        /// <summary>
        /// 判断是否在工作日内
        /// </summary>
        /// <returns></returns>
        private bool m_IsWorkingDay(DateTime startTime)
        {
            string strWeekNow = this.m_GetWeekNow(startTime);//当前周几
            string[] RestDay = _strRestDay.Split(',');
            if (RestDay.Contains(strWeekNow))
            {
                return false;
            }
            //判断当前时间是否在工作时间段内

            dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;

            TimeSpan dspNow = startTime.TimeOfDay;
            if (dspNow > dspWorkingDayAM && dspNow < dspWorkingDayPM)
            {
                return true;
            }
            return false;
        }
        //初始化默认值
        private void m_InitWorkingDay()
        {
            dspWorkingDayAM = DateTime.Parse(_strWorkingDayAM).TimeOfDay;
            dspWorkingDayPM = DateTime.Parse(_strWorkingDayPM).TimeOfDay;

        }
    }
}
