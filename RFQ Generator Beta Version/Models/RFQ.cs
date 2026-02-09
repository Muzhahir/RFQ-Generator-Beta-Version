using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System
{
    public class RFQ
    {
        public int Id { get; set; }
        public int CompanyId { get; set; }
        public int ClientId { get; set; }
        public DateTime CreatedAt { get; set; }
        public string RFQCode { get; set; }
        public string QuoteCode { get; set; }
        public string DeliveryPoint { get; set; }
        public decimal Discount { get; set; }
        public string DeliveryTerm { get; set; }
        public string Validity { get; set; }
        public object QuoteCodeCounter { get; internal set; }
    }
}