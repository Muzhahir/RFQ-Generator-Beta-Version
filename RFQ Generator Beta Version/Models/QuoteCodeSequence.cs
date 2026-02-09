using System;

namespace RFQ_Generator_System
{
    public class QuoteCodeSequence
    {
        public int Id { get; set; }
        public int CompanyId { get; set; }
        public int CurrentSequence { get; set; }
        public int LastResetYear { get; set; }
    }
}