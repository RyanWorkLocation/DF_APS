using System.ComponentModel.DataAnnotations;

namespace Web.Models
{
    public class Device
    {
        [Display(Name = "機台編號")]
        public int ID { set; get; }

        [Display(Name = "機台名稱")]
        public string MachineName { set; get; }

        [Display(Name = "Remark")]
        public string Remark { set; get; }

        [Display(Name = "群組名稱")]
        public string GroupName { set; get; }
    }
}