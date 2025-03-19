using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _4337Project
{
    public class Order
    {
        public int Id { get; set; }
        public string CodeOrder { get; set; }
        public string CreateDate { get; set; }
        public string CreateTime { get; set; }
        public string CodeClient { get; set; }
        public string Services { get; set; }
        public string Status { get; set; }
        public string ClosedDate { get; set; }
        public string ProkatTime { get; set; }
    }
}
