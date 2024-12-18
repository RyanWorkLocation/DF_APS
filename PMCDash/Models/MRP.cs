using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace PMCDash.Models
{
    public class MRP
    {
        public class Outsource
        {
            public int ID { get; set; }
            public string remark { get; set; }
            public string isOutsource { get; set; }
        }

        public class ProcessDetial
        {
            public int ID { get; set; }
            public string ProcessID { get; set; }
            public string MachineID { get; set; }
            public string remark { get; set; }
        }
    }
}
