using System;
using System.Collections.Generic;
using System.Linq;
using RFQ_Generator_System.Repositories;

namespace RFQ_Generator_System.Services
{
    public class RFQService
    {
        private readonly RFQRepo rfqRepo;
        private readonly RFQItemRepo rfqItemRepo;
        private readonly TemplateRepo templateRepo;
        private readonly FieldMappingRepo fieldMappingRepo;
        private readonly CompanyRepo companyRepo;
        private readonly ClientRepo clientRepo;
        private readonly QuoteCodeSequenceRepo quoteCodeSequenceRepo;

        public RFQService()
        {
            rfqRepo = new RFQRepo();
            rfqItemRepo = new RFQItemRepo();
            templateRepo = new TemplateRepo();
            fieldMappingRepo = new FieldMappingRepo();
            companyRepo = new CompanyRepo();
            clientRepo = new ClientRepo();
            quoteCodeSequenceRepo = new QuoteCodeSequenceRepo();
        }

        // -----------------------------
        // Quote code generation
        // -----------------------------

        public string GenerateQuoteCodePreview(string companyCode, string clientCode)
        {
            if (string.IsNullOrWhiteSpace(companyCode))
                return string.Empty;

            var company = companyRepo
                .GetAllCompanies()
                .FirstOrDefault(c => c.CompanyCode == companyCode);

            if (company == null)
                return string.Empty;

            int sequence = quoteCodeSequenceRepo.PeekNextSequence(company.Id);

            return FormatQuoteCode(companyCode, clientCode, sequence);
        }

        public string GenerateQuoteCode(string companyCode, string clientCode)
        {
            if (string.IsNullOrWhiteSpace(companyCode))
                return string.Empty;

            var company = companyRepo
                .GetAllCompanies()
                .FirstOrDefault(c => c.CompanyCode == companyCode);

            if (company == null)
                return string.Empty;

            int sequence = quoteCodeSequenceRepo.GetNextSequence(company.Id);

            return FormatQuoteCode(companyCode, clientCode, sequence);
        }

        public void UpdateSequenceFromManualEdit(int companyId, string quoteCode)
        {
            quoteCodeSequenceRepo.UpdateSequenceFromQuoteCode(companyId, quoteCode);
        }

        // -----------------------------
        // RFQ save and retrieval
        // -----------------------------

        /// <summary>
        /// Save RFQ header and items in a single transaction.
        /// Returns the new RFQ Id.
        /// </summary>
        public int SaveRFQ(RFQ rfq, List<RFQItem> items)
        {
            return rfqRepo.SaveRFQ(rfq, items);
        }

        public (RFQ rfq, List<RFQItem> items) GetRFQWithItems(int rfqId)
        {
            var rfq = rfqRepo.GetRFQById(rfqId);
            var items = rfqItemRepo.GetRFQItemsByRFQId(rfqId);
            return (rfq, items);
        }

        public List<RFQ> GetAllRFQs()
        {
            return rfqRepo.GetAllRFQs();
        }

        // -----------------------------
        // Supporting data
        // -----------------------------

        public Template GetTemplateByCompanyId(int companyId)
        {
            return templateRepo.GetTemplateByCompanyId(companyId);
        }

        public List<FieldMapping> GetFieldMappingsForTemplate(int templateId)
        {
            return fieldMappingRepo.GetFieldMappingsByTemplateId(templateId);
        }

        public List<Company> GetAllCompanies()
        {
            return companyRepo.GetAllCompanies();
        }

        public List<Client> GetAllClients()
        {
            return clientRepo.GetAllClients();
        }

        // -----------------------------
        // Sequence management
        // -----------------------------

        public void ResetAllQuoteCodeSequences()
        {
            quoteCodeSequenceRepo.ResetAllSequences();
        }

        public void ResetQuoteCodeSequence(int companyId)
        {
            quoteCodeSequenceRepo.ResetSequence(companyId);
        }

        // -----------------------------
        // Private helpers
        // -----------------------------

        private string FormatQuoteCode(string companyCode, string clientCode, int sequence)
        {
            string yearShort = DateTime.Now.Year.ToString().Substring(2, 2);

            switch (companyCode)
            {
                case "CG":
                    return $"CG-{DateTime.Now:MMyy}-{sequence:D6}";

                case "DE":
                    return $"RFP-{sequence:D12}";

                case "GA":
                    return $"GASB-{sequence:D4}-{DateTime.Now:ddMMyy}";

                case "MA":
                    return $"QUO-ML-{sequence:D4}";

                case "OGIT":
                    return $"OGIT{DateTime.Now:ddMM}-{yearShort}-{sequence:D3}";

                case "OP":
                    return $"Q-{sequence:D6}-{clientCode}";

                case "PO":
                    return $"{clientCode}-{sequence:D7}-EPOMS";

                case "SC":
                    return $"{clientCode}-QUO-SC-{sequence:D6}";

                default:
                    return $"{companyCode}-{sequence:D6}";
            }
        }
    }
}
