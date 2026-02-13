using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System
{
    public class Client
    {
        public int Id { get; set; }
        public string ClientName { get; set; }
        public string ClientCode { get; internal set; }

        public string DisplayText
        {
            get
            {
                if (string.IsNullOrEmpty(ClientCode))
                    return ClientName;
                return $"{ClientName} - ({ClientCode})";
            }
        }
    }
}
