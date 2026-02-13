using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RFQ_Generator_System
{
    public class Company
    {
        public int Id { get; set; }
        public string CompanyName { get; set; }
        public string CompanyCode { get; internal set; 
       
    }
    

    public string DispayText
        {
            get
            {
                if (string.IsNullOrEmpty(CompanyCode))
                    return CompanyName;
                return $"{CompanyName} - ({CompanyCode})";
            }
        }
    } 
}
