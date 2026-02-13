using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System
{
    public class RFQItem
    {
        public int Id { get; set; }
        public int RFQId { get; set; }
        public int ItemNo { get; set; }
        public string ItemDesc { get; set; }
        public int Quantity { get; set; }
        public int DeliveryTime { get; set; }
        public decimal UnitPrice { get; set; }
        public string UnitName { get; set; }

    }
}
