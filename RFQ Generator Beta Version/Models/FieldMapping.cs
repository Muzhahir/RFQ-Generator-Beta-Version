using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System
{
    public class FieldMapping
    {
        public int Id { get; set; }
        public int TemplateId { get; set; }
        public string FieldKey { get; set; }
        public string ExcellCell { get; set; }
    }
}
